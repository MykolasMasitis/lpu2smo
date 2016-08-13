PROCEDURE CorrMcod
 IF MESSAGEBOX('¬€ ’Œ“»“≈ «¿Ã≈Õ»“‹ MCOD?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 DIMENSION dimcorr(2,2)
 dimcorr(1,1) = '0106002'
 dimcorr(1,2) = '0306002'
 dimcorr(2,1) = '0106003'
 dimcorr(2,2) = '0306003'
 
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
 
 SELECT aisoms
 SCAN
  m.mcod = mcod
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
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT people
  SCAN 
   m.prmcod = prmcod
   IF m.prmcod=dimcorr(1,1)
    REPLACE prmcod WITH dimcorr(1,2)
   ENDIF 
   IF m.prmcod=dimcorr(2,1)
    REPLACE prmcod WITH dimcorr(2,2)
   ENDIF 
  ENDSCAN 
  USE IN people
  
  SELECT aisoms
  
 ENDSCAN
 WAIT CLEAR 
 USE 

 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ Œ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')

RETURN 