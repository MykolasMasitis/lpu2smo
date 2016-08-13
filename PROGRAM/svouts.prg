PROCEDURE SvOutS
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÑÂÅÐÈÒÜ ÐÅÃÈÑÒÐ Ñ ÍÎÌÅÐÍÈÊÎÌ?'+CHR(13)+CHR(10),4+32, '')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\outs.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÍÎÌÅÐÍÈÊ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\pilot.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË PILOT.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\outs', 'outs', 'shar')>0
  IF USED('outs')
   USE IN outs
  ENDIF 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('outs')
   USE IN outs
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('outs')
   USE IN outs
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('outs')
   USE IN outs
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR cursv (lpu_id n(4), mcod c(7), totrecs n(7), qrecs n(7), fndrecs n(7))
 CREATE CURSOR notfnd (sn_pol c(25))
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  m.lpuid = lpuid
  IF !SEEK(m.lpuid, 'pilot')
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
  ENDIF 
  SELECT people
  IF RECCOUNT()<=0
   IF USED('people')
    USE IN people
   ENDIF
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  m.totrecs = 0
  m.qrecs   = 0
  m.fndrecs = 0
  
  SCAN 
   m.totrecs = m.totrecs + 1
   m.sn_pol  = sn_pol
   m.qq      = qq
   IF m.qq != m.qcod
    LOOP 
   ENDIF 
   m.qrecs = m.qrecs + 1

   IF SEEK(LEFT(m.sn_pol,16), 'outs', 'enp')
    m.fndrecs = m.fndrecs + 1
    m.lpid  = outs.lpu_id
    m.lpcod = IIF(SEEK(m.lpid, 'sprlpu'), sprlpu.mcod, '')
*    REPLACE prmcod WITH m.lpcod
    LOOP 
   ENDIF 

   IF SEEK(LEFT(m.sn_pol,17), 'outs', 'kms')
    m.fndrecs = m.fndrecs + 1
    m.lpid  = outs.lpu_id
    m.lpcod = IIF(SEEK(m.lpid, 'sprlpu'), sprlpu.mcod, '')
*    REPLACE prmcod WITH m.lpcod
    LOOP 
   ENDIF 

   IF LEFT(m.sn_pol,2)=m.qcod
    m.vs = SUBSTR(m.sn_pol,7,9)
    IF SEEK(m.vs, 'outs', 'vsn')
     m.fndrecs = m.fndrecs + 1
     m.lpid  = outs.lpu_id
     m.lpcod = IIF(SEEK(m.lpid, 'sprlpu'), sprlpu.mcod, '')
     LOOP 
    ENDIF 
   ENDIF 

   IF LEN(ALLTRIM(m.sn_pol))=9
    m.vs = ALLTRIM(m.sn_pol)
    IF SEEK(m.vs, 'outs', 'vsn')
     m.fndrecs = m.fndrecs + 1
     m.lpid  = outs.lpu_id
     m.lpcod = IIF(SEEK(m.lpid, 'sprlpu'), sprlpu.mcod, '')
     LOOP 
    ENDIF 
   ENDIF 
   
   INSERT INTO notfnd (sn_pol) VALUES (m.sn_pol)
   
*   REPLACE prmcod WITH ''
  ENDSCAN 

  IF USED('people')
   USE IN people
  ENDIF

  INSERT INTO cursv (lpu_id, mcod, totrecs, qrecs, fndrecs) VALUES ;
   (m.lpuid, m.mcod, m.totrecs, m.qrecs, m.fndrecs)

  SELECT aisoms
 ENDSCAN 

 WAIT CLEAR 

 crname = 'svrslt'+m.qcod
 crpath = pbase+'\'+m.gcperiod
 SELECT cursv
 COPY TO &crpath\&crname
 USE 
 
 SELECT notfnd
 crname = 'nfnd'+m.qcod
 COPY TO &crpath\&crname
 USE 

 IF USED('pilot')
  USE IN pilot
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('outs')
  USE IN outs
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ÐÀÑ×ÅÒ ÇÀÊÎÍ×ÅÍ!'+CHR(13)+CHR(10),0+64,'')
 
 
RETURN 