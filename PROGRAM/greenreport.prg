PROCEDURE GreenReport
 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ'+CHR(13)+CHR(10)+;
  'ÑÂÎÄÍÓÞ ÒÀÁËÈÖÓ ÏÐÈ×ÈÍ ÑÍßÒÈÉ?'+CHR(13)+CHR(10),4+32,'ÌÝÝ')==7
  RETURN 
 ENDIF 
 
 m.pgdat1 = m.tdat1
 m.pgdat2 = m.tdat2
 m.ischecked = .f.
 DO FORM SelPeriod
 IF m.ischecked = .f.
  RETURN 
 ENDIF 
 
 CREATE CURSOR curgrrep (osn230 c(5), et c(1), k_u n(5)) 
 INDEX on osn230+et TAG osn230et
 SET ORDER TO osn230et

 m.nmonthes = (MONTH(m.pgdat2) - MONTH(m.pgdat1))

 FOR m.nmonth = 0 TO m.nmonthes
 
  m.pgdat = GOMONTH(m.pgdat2, -m.nmonthes+m.nmonth)
  m.pgperiod = STR(YEAR(m.pgdat),4)+PADL(MONTH(m.pgdat),2,'0')
  WAIT m.pgperiod WINDOW NOWAIT 
  =GrRepOne(m.pgperiod)
  WAIT CLEAR 

 NEXT 
 
 SELECT curgrrep
 COPY TO &pout\GreenReort
 USE 

RETURN 
 
FUNCTION  GrRepOne(_pgperiod)
 
 m.lcPgPeriod = _pgperiod

 IF !fso.FileExists(pBase+'\'+lcPgPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&lcPgPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+lcPgPeriod+'\'+'nsi'+'\TarifN', 'Tarif', 'SHARED', 'cod ') > 0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+lcPgPeriod+'\nsi\sookodxx', 'sookod', 'SHARED', 'er_c') > 0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 
 SCAN 
  m.mcod = mcod
  m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

*  WAIT m.mcod WINDOW NOWAIT 
  
  IF !fso.FolderExists(pbase+'\'+lcPgPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\e'+m.mcod, 'eerror', 'shar', 'rid')>0
   IF USED('eerror')
    USE IN eerror
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 

  SELECT merror
  SCAN 
   m.cod      = cod 
   m.osn230   = osn230
   m.e_period = e_period
   m.et       = et
   
   IF EMPTY(m.osn230)
    LOOP 
   ENDIF 
   IF EMPTY(m.et)
    m.et = '2'
   ENDIF 
   IF EMPTY(e_period)
    LOOP 
   ENDIF 

   m.eperiod = '01.'+SUBSTR(m.e_period,5,2)+'.'+LEFT(m.e_period,4)
   TRY 
    m.eperiod = CTOD(m.eperiod)
   CATCH 
    LOOP 
   ENDTRY 
   
   IF !BETWEEN(m.eperiod,m.pgdat1,m.pgdat2)
    LOOP 
   ENDIF 
   
   m.vir = m.osn230+m.et
   IF !SEEK(m.vir, 'curgrrep')
    INSERT INTO curgrrep (osn230,et,k_u) VALUES (m.osn230,m.et,1) 
   ELSE 
    m.o_ku = curgrrep.k_u
    m.n_ku = m.o_ku + 1
    UPDATE curgrrep SET k_u=m.n_ku WHERE osn230=m.osn230 AND et=m.et
   ENDIF 
   
  ENDSCAN 
  USE IN merror

  SELECT eerror
  SET RELATION TO LEFT(c_err,2) INTO sookod
  SCAN 
   m.osn230   = LEFT(sookod.osn230,5)
   m.et       = '1'
   
   IF EMPTY(m.osn230)
    LOOP 
   ENDIF 

   m.vir = m.osn230+m.et
   IF !SEEK(m.vir, 'curgrrep')
    INSERT INTO curgrrep (osn230,et,k_u) VALUES (m.osn230,m.et,1) 
   ELSE 
    m.o_ku = curgrrep.k_u
    m.n_ku = m.o_ku + 1
    UPDATE curgrrep SET k_u=m.n_ku WHERE osn230=m.osn230 AND et=m.et
   ENDIF 
   
  ENDSCAN 
  SET RELATION OFF INTO sookod
  USE IN eerror

*  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE IN aisoms
 USE IN tarif
 USE IN sookod

RETURN 

