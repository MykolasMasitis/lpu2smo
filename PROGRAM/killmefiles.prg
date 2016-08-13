PROCEDURE KillMeFiles

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÍÓËÈÒÜ M-ÔÀÉËÛ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÏÅÐÅÄÓÌÀÅÒÅ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 FOR m.nmm=0 TO 24
  m.lcperiod = LEFT(DTOS(GOMONTH(m.tdat2,-m.nmm)),6)
  m.lpath = pbase+'\'+m.lcperiod
  IF !fso.FolderExists(m.lpath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lpath+'\aisoms.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.lcperiod+'...' WINDOW NOWAIT 
  =KillOnePeriod(m.lpath)
  WAIT CLEAR 

 NEXT 

RETURN 

FUNCTION KillOnePeriod
 PARAMETERS m.lpath
 PRIVATE m.llcpath
 m.llcpath = m.lpath
 IF OpenFile(m.llcpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.llcpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.llcpath+'\'+m.mcod+'\m'+m.mcod, 'merror', 'excl')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT merror
  ZAP 
  USE IN merror 

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 