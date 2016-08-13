PROCEDURE DelPPA
 IF MESSAGEBOX('¬€ ’Œ“»“≈ —Õﬂ“‹ Œÿ»¡ » "PPA"'+CHR(13)+CHR(10)+;
  '—Œ "—¬Œ»’" œ¿÷»≈Õ“Œ¬?'+CHR(13)+CHR(10),4+16,'')=7
  RETURN 
 ENDIF 
 
 ppath = pbase+'\'+m.gcperiod
 IF OpenFile(ppath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\row2codes', 'rcodes', 'shar', 'cod')>0
  IF USED('rcodes')
   USE IN rcodes
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 SELECT aisoms
 SCAN 
  m.mcod = mcod
  WAIT m.mcod+'...' WINDOW NOWAIT 
  IF !fso.FolderExists(ppath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppath+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(ppath+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar')>0
   IF USED('error')
    USE IN error
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('error')<=0
   IF USED('error')
    USE IN error
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  SELECT error
  COUNT FOR c_err='PPA' TO m.nerrs
  IF m.nerrs<=0
   IF USED('error')
    USE IN error
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  && ≈ÒÚ¸ Ú‡ÍËÂ Ó¯Ë·ÍË!
  
  IF OpenFile(ppath+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(ppath+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  CREATE CURSOR curpol (recid i, sn_pol c(25))	
  INDEX on recid TAG recid
  INDEX on sn_pol TAG sn_pol
  
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SELECT error
  SET ORDER TO rid 
  SET RELATION TO rid INTO talon 
  
  SCAN 
   m.c_err = c_err
   m.cod   = talon.cod
   IF m.c_err!='PPA'
    LOOP 
   ENDIF 
   IF people.prmcod!=m.mcod
    LOOP 
   ENDIF 
   IF SEEK(m.cod, 'rcodes')
    LOOP 
   ENDIF 
   IF IsKd(m.cod)
    LOOP 
   ENDIF 
   
   m.sn_pol = people.sn_pol 
   m.recid  = people.recid
   IF !SEEK(m.sn_pol, 'curpol')
    INSERT INTO curpol FROM MEMVAR 
   ENDIF 
   DELETE 
   
  ENDSCAN 
  
  SET ORDER TO rrid
  SET ORDER TO recid IN curpol 
  
  SCAN 
   m.c_err = c_err
   IF m.c_err!='PNA'
    LOOP 
   ENDIF 
   m.rid = rid
   IF !SEEK(m.rid, 'curpol')
    LOOP 
   ENDIF 
   
   DELETE 
   
  ENDSCAN 
  
  SET RELATION OFF INTO talon 
  USE 
  SELECT talon 
  SET RELATION OFF INTO people
  USE 
  USE IN people 
  
  USE IN curpol 
  
  SELECT aisoms
  
 ENDSCAN 
 USE IN aisoms

 USE IN rcodes
 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 