PROCEDURE re_calcmee

USE pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN' IN 0 ALIAS Tarif SHARED ORDER cod 

SELECT AisOms 
SET ORDER TO mcod 
GO TOP 

m.tmee = 0 
m.tmeebad = 0
m.tsum_mee = 0

SCAN
 WAIT mcod WINDOW NOWAIT 
 m.IsVed   = IIF(LEFT(mcod,1) == '0', .F., .T.)

 lcDir = pBase + '\' + m.gcperiod + '\' + mcod
 IF fso.FileExists(lcDir+'\talon.dbf') AND ;
    fso.FileExists(lcDir+'\e'+mcod+'.dbf')
  tn_result = 0
  tn_result = tn_result + OpenFile("&lcDir\Talon", "Talon", "SHARE", 'sn_pol')
  tn_result = tn_result + OpenFile("&lcDir\people", "people", "SHARE", 'sn_pol')
  tn_result = tn_result + OpenFile("&lcDir\e"+mcod, "Error", "SHARE", 'rid')
  IF tn_result == 0

   CREATE CURSOR curtot (sn_pol c(25))
   SELECT curtot
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol
   
   CREATE CURSOR curbad (sn_pol c(25))
   SELECT curbad
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   SELECT Talon
   SET RELATION TO recid INTO error 
   m.mee = 0 
   m.meebad = 0
   m.sum_mee = 0
   
   m.sum_paz = 0 

   SCAN 
    m.d_type = d_type
    m.sn_pol = sn_pol

    IF !EMPTY(err_mee)
     IF EMPTY(et)
      REPLACE et WITH '2'
     ENDIF 
     m.mee = m.mee + IIF(!SEEK(m.sn_pol, 'curtot'), 1, 0)
     IF  LEFT(UPPER(ALLTRIM(err_mee)),2) != 'W0'
      m.meebad = m.meebad + IIF(!SEEK(m.sn_pol, 'curbad'), 1, 0)
      IF !SEEK(m.sn_pol, 'curbad')
       INSERT INTO curbad (sn_pol) VALUES (m.sn_pol)
      ENDIF 
      IF EMPTY(e_cod) AND EMPTY(e_tip) AND EMPTY(e_ku) && Полное снятие!
       m.sum_mee = m.sum_mee + s_all
      ELSE 
       IF (!EMPTY(e_cod) AND cod != e_cod) OR ;
          (!EMPTY(e_ku) AND k_u != e_ku) OR ;
          (!EMPTY(e_tip) AND e_tip != tip)
        m.ns_all = fsumm(e_cod, e_tip, e_ku, m.IsVed)
        m.delta = s_all - m.ns_all
        m.sum_mee = m.sum_mee + m.delta
       ENDIF 
      ENDIF 
     ELSE 
     ENDIF 
    ENDIF 
    
    IF !SEEK(m.sn_pol, 'curtot')
     INSERT INTO curtot (sn_pol) VALUES (m.sn_pol)
    ENDIF 
    
   ENDSCAN 
   SET RELATION OFF INTO error
   USE
   USE IN error
   USE IN people
   USE IN curtot
   USE IN curbad
   SELECT AisOms

   REPLACE mee WITH m.mee, meebad WITH m.meebad, sum_mee WITH m.sum_mee

   m.tmee     = m.tmee + m.mee
   m.tmeebad  = m.tmeebad + m.meebad
   m.tsum_mee = m.tsum_mee + m.sum_mee

  ENDIF 
 ENDIF 

 ExpView.get_mee.Value     = m.tmee
 ExpView.get_meebad.Value  = m.tmeebad
 ExpView.get_sum_mee.Value = m.tsum_mee
 ExpView.refresh

ENDSCAN 
WAIT CLEAR 
GO TOP

USE IN Tarif
