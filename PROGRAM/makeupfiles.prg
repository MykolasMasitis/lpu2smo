PROCEDURE MakeUPFiles
 IF MESSAGEBOX(CHR(13)+	CHR(10)+'¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ UP-‘¿…À€?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\upqqxxxx.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—“”“—“¬”≈“ ÿ¿¡ÀŒÕ UPQQXXXX.DBF!'+CHR(13)+CHR(10),0+16,'Templates')
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\lputpn', 'lputpn', 'shar', 'lpu_id')>0
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN lputpn
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')>0
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  USE IN tarif
  USE IN lputpn
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN horlpu
  USE IN tarif
  USE IN lputpn
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 
 SELECT aisoms
 SCAN FOR !DELETED()

  m.lpuid = lpuid
  m.mcod  = mcod 
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  
  IF !SEEK(m.lpuid, 'pilot')
   IF !SEEK(m.lpuid, 'horlpu')
    LOOP 
   ENDIF 
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('error')
    USE IN error 
   ENDIF 
   IF USED('people')
    USE IN people 
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 

  m.upfile = 'UP'+UPPER(m.qcod)+PADL(m.lpuid,4,'0')
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile+'.dbf')
  ENDIF 
  fso.CopyFile(ptempl+'\upqqxxxx.dbf', pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile+'.dbf')
  =OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.upfile, 'upfile', 'shar')
  
  SELECT talon 
  SET RELATION TO recid INTO error
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  SCAN 
   IF m.IsLpuTpn = .T.
    m.fil_id = fil_id
    IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
     LOOP 
    ENDIF 
   ENDIF 
   m.prmcod  = people.prmcod
   IF !EMPTY(error.c_err)
    LOOP 
   ENDIF 
   IF EMPTY(m.prmcod)
    LOOP 
   ENDIF 
   IF m.prmcod=m.mcod
    LOOP 
   ENDIF 
   
   m.prlpuid = IIF(SEEK(m.prmcod, 'pilot', 'mcod'), pilot.lpu_id, 0)
   IF !SEEK(m.prlpuid, 'pilot')
    LOOP 
   ENDIF 

   m.cod     = cod
   m.otd = SUBSTR(otd,2,2)
   IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) OR IsEKO(m.cod) OR IsPat(m.cod) 
    LOOP 
   ENDIF 
   IF INLIST(m.cod, 97003,97010,97011,197010,197011,149017,97007) AND BETWEEN(m.tdat1,{01.01.2015},{01.02.2015})
    LOOP 
   ENDIF 
   IF INLIST(m.cod, 97010,97011,197010,197011,97007) AND m.tdat1>{01.02.2015}
    LOOP 
   ENDIF 
   IF INLIST(m.otd,'01','70','73')
    LOOP 
   ENDIF 
   IF IsSimult(m.cod)
    LOOP 
   ENDIF 
   
   m.cod     = cod
   m.k_u     = k_u
   m.s_all   = s_all
   m.pr_all  = m.s_all
   m.lpu_ord = lpu_ord
   m.otd     = SUBSTR(otd,2,2)
   
   m.lIs02   = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
   m.prlpuid = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)

*   IF (!EMPTY(m.lpu_ord) AND m.lpu_ord=m.prlpuid) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(m.otd,'08','92')))
   IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(m.otd,'08','92')))
    m.vz = 1 
   ELSE && IF EMPTY(m.lpu_ord) AND !(m.lIs02=.T. OR INLIST(m.otd,'08','92'))
    m.vz = 2
   ENDIF 
   
   IF m.vz=2
    INSERT INTO upfile (period,lpu_id,cod,k_u,pr_all) VALUES ;
    (m.gcperiod,m.prlpuid,m.cod,m.k_u,m.s_all)
   ENDIF 

  ENDSCAN 

  SELECT upfile
  REPLACE ALL recid WITH PADL(RECNO(),7,'0')
  USE IN upfile

  SELECT talon 
  SET RELATION OFF INTO people
  SET RELATION OFF INTO error
  USE IN people
  USE IN talon 
  USE IN error
  
  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 USE IN horlpu
 USE IN tarif
 USE IN lputpn
 USE IN pilot
 USE IN sprlpu
 
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 