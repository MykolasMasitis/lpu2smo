PROCEDURE re_calc

SET DELETED ON 

USE pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN' IN 0 ALIAS Tarif SHARED ORDER cod 

SELECT AisOms 
SET ORDER TO 
GO TOP 

m.ts_pred  = 0
m.tsum_flk = 0 
m.tmee = 0 
m.tmeebad = 0
m.tsum_mee = 0

SCAN
 lcDir = pBase + '\' + m.gcperiod + '\' + mcod
 IF fso.FileExists(lcDir+'\talon.dbf') AND ;
    fso.FileExists(lcDir+'\e'+mcod+'.dbf')
  tn_result = 0
  tn_result = tn_result + OpenFile("&lcDir\Talon", "Talon", "SHARE", 'sn_pol')
  tn_result = tn_result + OpenFile("&lcDir\people", "people", "SHARE", 'sn_pol')
  tn_result = tn_result + OpenFile("&lcDir\e"+mcod, "Error", "SHARE", 'rid')
  IF tn_result == 0
   WAIT mcod WINDOW NOWAIT 

   CREATE CURSOR pazamb (sn_pol c(25))
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol
  
   CREATE CURSOR pazppl (sn_pol c(25))
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR pazdst (c_i c(25))
   INDEX on c_i TAG c_i
   SET ORDER TO c_i

   CREATE CURSOR pazst (c_i c(25))
   INDEX ON c_i TAG c_i
   SET ORDER TO c_i
   
   CREATE CURSOR pazambmek (sn_pol c(25))
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR pazdstmek (c_i c(25))
   INDEX on c_i TAG c_i
   SET ORDER TO c_i

   CREATE CURSOR pazstmek (c_i c(25))
   INDEX ON c_i TAG c_i
   SET ORDER TO c_i

   CREATE CURSOR pazambmee (sn_pol c(25))
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR pazdstmee (c_i c(25))
   INDEX on c_i TAG c_i
   SET ORDER TO c_i

   CREATE CURSOR pazstmee (c_i c(25))
   INDEX ON c_i TAG c_i
   SET ORDER TO c_i

   CREATE CURSOR ambchkdmee (sn_pol c(25))
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR dstchkdmee (sn_pol c(25))
   INDEX on sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR stchkdmee (c_i c(25))
   INDEX ON c_i TAG c_i
   SET ORDER TO c_i

   SELECT aisoms

   m.IsVed   = IIF(LEFT(mcod,1) == '0', .F., .T.)
   SELECT Talon
   SET RELATION TO recid INTO error 
   m.s_pred  = 0
   m.sum_flk = 0 
   m.mee = 0 
   m.meebad = 0
   m.sum_mee = 0
   
   m.polis = sn_pol
   m.c_i   = c_i
   m.sum_paz = 0 

   m.paz_amb = 0
*   m.paz_ppl = 0
   m.paz_dst = 0
   m.paz_st  = 0

   m.usl_amb = 0
   m.kd_dst  = 0
   m.ms_st   = 0

   m.usl_ambmek = 0
   m.kd_dstmek  = 0
   m.ms_stmek   = 0

   m.usl_ambmee = 0
   m.kd_dstmee  = 0
   m.ms_stmee   = 0

   m.sum_amb = 0
   m.sum_dst = 0
   m.sum_st  = 0

   m.paz_ambmek = 0
   m.paz_dstmek = 0
   m.paz_stmek  = 0

   m.sum_ambmek = 0
   m.sum_dstmek = 0
   m.sum_stmek  = 0

   m.paz_ambmee = 0
   m.paz_dstmee = 0
   m.paz_stmee = 0

   m.sum_ambmee = 0
   m.sum_dstmee = 0
   m.sum_stmee  = 0

   m.ambchkdmee = 0
   m.dstchkdmee = 0
   m.stchkdmee = 0

   SCAN 
    m.cod = cod
    m.tip = tip
    m.sn_pol = sn_pol
    m.d_type = d_type
    m.k_u = k_u
    m.summa = fsumm(m.cod, m.tip, m.k_u, m.IsVed)
    REPLACE s_all WITH m.summa
    m.s_pred = m.s_pred + s_all
    IF !EMPTY(error.c_err)
     m.sum_flk = m.sum_flk + s_all
    ENDIF 

    IF !EMPTY(err_mee)
     IF EMPTY(et)
      REPLACE et WITH '2'
     ENDIF 
     m.mee = m.mee + 1
     IF  LEFT(UPPER(ALLTRIM(err_mee)),2) != 'W0'
      m.meebad = m.meebad + 1
      IF EMPTY(e_cod) AND EMPTY(e_tip) AND EMPTY(e_ku) && Полное снятие!
       m.sum_mee = m.sum_mee + s_all
      ELSE 
       IF (!EMPTY(e_cod) AND cod != e_cod) OR ;
          (!EMPTY(e_ku) AND k_u != e_ku) OR ;
          (!EMPTY(e_tip) AND e_tip != tip)
        m.ns_all = fsumm(e_cod, e_tip, e_ku, m.IsVed)
        m.delta = m.ns_all - s_all
        m.sum_mee = m.sum_mee + m.delta
       ENDIF 
      ENDIF 
     ELSE 
     ENDIF 
    ENDIF 

    IF sn_pol = m.polis
     m.sum_paz = m.sum_paz + m.summa
    ELSE 
     IF SEEK(m.polis, 'people')
      REPLACE people.s_all WITH m.sum_paz IN people 
     ENDIF 
     m.polis = sn_pol
     m.sum_paz = m.summa
    ENDIF 

    IF IsUsl(m.cod)
     m.sum_amb = m.sum_amb + s_all
     m.usl_amb = m.usl_amb + k_u
     IF !SEEK(m.sn_pol, 'pazamb')
      INSERT INTO pazamb (sn_pol) VALUES (m.sn_pol)
      m.paz_amb = m.paz_amb + 1
     ENDIF 
    ENDIF 
   
    IF !EMPTY(m.tip)
     m.sum_st = m.sum_st + s_all
     m.ms_st = m.ms_st + 1
     IF !SEEK(m.sn_pol, 'pazst')
      INSERT INTO pazst (c_i) VALUES (m.c_i)
      m.paz_st = m.paz_st + 1
     ENDIF 
    ENDIF 

    IF IsKd(m.cod)
     m.sum_dst = m.sum_dst + s_all
     m.kd_dst = m.kd_dst + k_u
     IF !SEEK(m.sn_pol, 'pazppl')
      INSERT INTO pazppl (sn_pol) VALUES (m.sn_pol)
*      m.paz_ppl = m.paz_ppl + 1
     ENDIF 
    ENDIF 

    IF IsKd(m.cod)
     IF !SEEK(m.sn_pol, 'pazdst')
      INSERT INTO pazdst (c_i) VALUES (m.c_i)
      m.paz_dst = m.paz_dst + 1
     ENDIF 
    ENDIF 

    IF IsUsl(m.cod) AND !EMPTY(err_mee)
     IF !SEEK(m.sn_pol, 'ambchkdmee')
      INSERT INTO ambchkdmee (sn_pol) VALUES (m.sn_pol)
      m.ambchkdmee = m.ambchkdmee + 1
     ENDIF 
    ENDIF 

    IF IsKD(m.cod) AND !EMPTY(err_mee)
     IF !SEEK(m.sn_pol, 'dstchkdmee')
      INSERT INTO dstchkdmee (sn_pol) VALUES (m.sn_pol)
      m.dstchkdmee = m.dstchkdmee + 1
     ENDIF 
    ENDIF 

    IF (IsMES(m.cod) OR IsVMP(m.cod)) AND !EMPTY(err_mee)
     IF !SEEK(m.c_i, 'stchkdmee')
      INSERT INTO stchkdmee (c_i) VALUES (m.c_i)
      m.stchkdmee = m.stchkdmee + 1
     ENDIF 
    ENDIF 

   ENDSCAN 
   SET RELATION OFF INTO error
   USE
   USE IN error
   USE IN people
   USE IN pazamb
   USE IN pazdst
   USE IN pazst
   USE IN pazppl
   USE IN pazambmek
   USE IN pazdstmek
   USE IN pazstmek
   USE IN ambchkdmee
   USE IN dstchkdmee
   USE IN stchkdmee

   SELECT AisOms

   REPLACE s_pred WITH m.s_pred, sum_flk WITH m.sum_flk, ;  
    mee WITH m.mee, meebad WITH m.meebad, ;
    sum_mee WITH m.sum_mee, ;
    paz_amb WITH m.paz_amb, paz_dst WITH m.paz_dst, paz_st WITH m.paz_st,;
    sum_amb WITH m.sum_amb, sum_dst WITH m.sum_dst, sum_st WITH m.sum_st,;
    usl_amb WITH m.usl_amb, kd_dst WITH m.kd_dst, ms_st WITH m.ms_st,;
    ambchkdmee WITH m.ambchkdmee, dstchkdmee WITH m.dstchkdmee, stchkdmee WITH m.stchkdmee 

   m.ts_pred  = m.ts_pred + m.s_pred
   m.tsum_flk = m.tsum_flk + m.sum_flk
   m.tmee     = m.tmee + m.mee
   m.tmeebad  = m.tmeebad + m.meebad
   m.tsum_mee = m.tsum_mee + m.sum_mee

  ENDIF 
 ENDIF 

 MailView.get_sum_flk.Value = m.tsum_flk
* MailView.get_mee.Value     = m.tmee
* MailView.get_meebad.Value  = m.tmeebad
 MailView.get_sum_mee.Value = m.tsum_mee
 MailView.refresh

ENDSCAN 
WAIT CLEAR 
GO TOP

USE IN Tarif
