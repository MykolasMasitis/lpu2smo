PROCEDURE MakeUDFiles
 IF MESSAGEBOX(CHR(13)+	CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ UD-ÔÀÉËÛ?'+CHR(13)+CHR(10),4+32,'')=7
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

 
 dufile = pbase+'\'+m.gcperiod+'\ud'+m.qcod+m.gcperiod
 IF fso.FileExists(dufile + '.dbf')
  fso.DeleteFile(dufile + '.dbf')
 ENDIF 
 
 CREATE TABLE &dufile ;
  (mcod c(7), lpu_id n(6), prmcod c(7), prlpu_id n(6), cod n(6), k_u n(3), pr_all n(12,2), vz n(1))
 USE 
 =OpenFile(dufile, 'dufile', 'shar')
 
 SELECT aisoms
 SCAN FOR !DELETED()

  m.lpuid = lpuid
  m.mcod  = mcod 
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  
  IF !SEEK(m.lpuid, 'pilot')
   IF !SEEK(m.lpuid, 'horlpu')
    LOOP 
   ENDIF 
*   LOOP 
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

  SELECT talon 
  SET RELATION TO recid INTO error
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  SCAN 
*   IF m.IsLpuTpn = .T.
*    m.fil_id = fil_id
*    IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*     LOOP 
*    ENDIF 
*   ENDIF 
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
*    IF !SEEK(m.prlpuid, 'horlpu')
*     LOOP 
*    ENDIF 
    LOOP 
   ENDIF 

   m.cod     = cod
   m.otd     = SUBSTR(otd,2,2)
   m.d_type  = d_type
*   m.IsTpnR  = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod) OR IsEko(m.cod)), .T., .F.)
   m.IsTpnR  = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
   m.ord     = ord
   m.lpu_ord = lpu_ord

*   IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) OR IsEKO(m.cod)
   IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod)
    LOOP 
   ENDIF 
   IF m.IsTpnR OR m.d_type='s'
    LOOP 
   ENDIF 
   IF INLIST(m.otd,'70','73') AND IsStac(m.mcod)
    LOOP 
   ENDIF 
   IF INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
    LOOP 
   ENDIF 
   IF m.ord=7 AND m.lpu_ord=7665
    LOOP 
   ENDIF 

*   IF INLIST(m.otd,'70','73','93')
*    LOOP 
*   ENDIF 
*   IF INLIST(m.otd,'01','90') AND !IsStac(m.mcod)
*    LOOP 
*   ENDIF 
   
   m.cod     = cod
   m.k_u     = k_u
   m.s_all   = s_all
   m.lpu_ord = lpu_ord
   m.otd     = SUBSTR(otd,2,2)
   
   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)

*   IF EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(m.otd,'08','92'))
   IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(m.otd,'08','92')))
    m.vz = 1
   ELSE 
    m.vz = 2
   ENDIF 
*   m.vz = IIF(!EMPTY(m.lpu_ord), 1, 2)
   
   INSERT INTO dufile (mcod,lpu_id,prmcod,prlpu_id,cod,k_u,pr_all,vz) VALUES ;
    (m.mcod,m.lpuid,m.prmcod,m.prlpuid,m.cod,m.k_u,m.s_all,m.vz)
   
  ENDSCAN 

  SET RELATION OFF INTO people
  SET RELATION OFF INTO error
  USE IN people
  USE IN talon 
  USE IN error
  
  WAIT CLEAR 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms
 
 SELECT dufile
 SELECT prmcod DISTINCT FROM dufile INTO CURSOR svcur
 IF _tally<=0
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\udqqxxxx.dbf')
  USE IN svcur
  USE IN dufile
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÒÓÒÑÒÂÓÅÒ ØÀÁËÎÍ UDQQXXXX.DBF!'+CHR(13)+CHR(10),0+16,'Templates')
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar', 'mcod')<=0
   SELECT prmcod as mcod, SUM(pr_all) as s_all WHERE vz=1 FROM dufile GROUP BY prmcod INTO CURSOR curstat
   SELECT curstat
   SET RELATION TO mcod INTO pr4
   m.sumstat = 0
   m.sumpr4  = 0
   SCAN 
    m.sumstat = m.sumstat + s_all
    m.sumpr4  = m.sumpr4 + pr4.s_others
   ENDSCAN 
   SET RELATION OFF INTO pr4
   USE IN pr4
   USE IN curstat
   IF m.sumstat!=m.sumpr4
    MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÍÀÐÓÆÅÍÎ ÐÀÑÕÎÆÄÅÍÈÅ ÌÅÆÄÓ ÔÀÉËÎÌ PR4 È'+CHR(13)+CHR(10)+;
     'È ÑÔÎÐÌÈÐÎÂÀÍÍÛÌÈ UD-ÔÀÉËÀÌÈ!'+CHR(13)+CHR(10)+;
     'ÏÎ ÄÀÍÍÛÌ PR4 ÑÓÌÌÀ S_OTHERS='+TRANSFORM(m.sumpr4,'99999999.99')+CHR(13)+CHR(10)+;
     'ÏÎ ÄÀÍÍÛÌ UD-ÔÀÉËÎÂ ÑÓÌÌÀ='+TRANSFORM(m.sumstat,'99999999.99')+CHR(13)+CHR(10),0+48,'')
   ELSE 
    MESSAGEBOX(CHR(13)+CHR(10)+'ÑÓÌÌÀ S_OTHERS ÔÀÉËÀ PR4'+CHR(13)+CHR(10)+;
     'ÑÎÎÒÂÅÒÑÒÂÓÅÒ ÑÓÌÌÅ UD-ÔÀÉËÎÂ'+CHR(13)+CHR(10)+;
     'È ÑÎÑÒÀÂËßÅÒ '+TRANSFORM(m.sumstat,'99999999.99')+CHR(13)+CHR(10),0+64,'')
   ENDIF 
  ELSE 
   IF USED('pr4')
    USE 
   ENDIF 
   MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÓÄÀËÎÑÜ ÎÒÊÐÛÒÜ ÔÀÉË PR4,'+CHR(13)+CHR(10)+;
    'ÑÂÅÐÊÀ ÄÀÍÍÛÕ ÍÅ ÏÐÎÈÇÂÎÄÈÒÑß!',0+64,'')
  ENDIF 
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'ÔÀÉË PR4 ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ,'+CHR(13)+CHR(10)+;
   'ÑÂÅÐÊÀ ÄÀÍÍÛÕ ÍÅ ÏÐÎÈÇÂÎÄÈÒÑß!',0+64,'')
 ENDIF 
 
 SELECT svcur
 SCAN 
  m.prmcod = prmcod 
  m.lppid  = IIF(SEEK(m.prmcod, 'pilot', 'mcod'), pilot.lpu_id, 0)
  m.uddfile = 'UD'+UPPER(m.qcod)+PADL(m.lppid,4,'0')
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
  ENDIF 

  fso.CopyFile(ptempl+'\udqqxxxx.dbf', pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile+'.dbf')
  
  =OpenFile(pbase+'\'+m.gcperiod+'\'+m.prmcod+'\'+m.uddfile, 'udfile', 'shar')
  
  SELECT * FROM dufile WHERE prmcod = m.prmcod INTO CURSOR curmmm
  SELECT curmmm
  m.mcod   = mcod
  m.lpu_id = lpu_id
*  m.period = m.gcperiod
  SCAN 
   m.cod    = cod 
   m.k_u    = k_u
   m.pr_all = pr_all
   m.lpu_id = lpu_id
   m.vz     = vz
   IF m.vz=1
    INSERT INTO udfile FROM MEMVAR 
   ENDIF 
  ENDSCAN 
  USE IN curmmm
  SELECT udfile
  REPLACE ALL recid WITH PADL(RECNO(),7,'0')
  USE IN udfile
  
  SELECT svcur
  
 ENDSCAN 
 USE IN svcur
 USE IN horlpu 
 USE IN pilot 
 USE IN lputpn
 USE IN tarif
 USE IN dufile
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 