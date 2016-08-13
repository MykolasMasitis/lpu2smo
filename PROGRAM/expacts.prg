PROCEDURE ExpActs
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ «¿√–”«»“‹ ¿ “€ › —œ≈–“»«?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ › —œŒ–“¿-»ÃœŒ–“¿'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pExpImp+'\svactsimp.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À »ÃœŒ–“¿ SVACTSIMP.DBF Õ≈ Œ¡Õ¿–”∆≈Õ'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pExpImp+'\svactsimp.dbf', 'svactsimp', 'shar')>0
  IF USED('svactsimp')
   USE IN svactsimp
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pmee+'\svacts\svacts.dbf', 'svacts', 'shar')>0
  IF USED('svactsimp')
   USE IN svactsimp
  ENDIF 
  IF USED('svacts')
   USE IN svacts
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT svactsimp
 m.nActsAdded = 0
 SCAN 
  m.period = period
  m.mcod   = mcod 
  m.codexp = codexp
  m.actname = ALLTRIM(actname)

  ooal = ALIAS()
  SELECT recid FROM svacts WHERE period=m.period AND mcod=m.mcod AND codexp=m.codexp ;
  INTO CURSOR rqwest NOCONSOLE  
  m.nfileid = recid
  USE 
  SELECT (ooal)
 
  IF m.nfileid>0
*   MESSAGEBOX('!',0+64,'')
  ELSE 
   IF !fso.FileExists(pexpimp+'\'+m.actname)
    LOOP 
   ENDIF 
   m.nActsAdded = m.nActsAdded + 1
   INSERT INTO svacts (period,mcod,codexp) VALUES (m.period,m.mcod,m.codexp)
   m.nfileid = GETAUTOINCVALUE()
   DocName = pmee+'\svacts\'+PADL(m.nfileid,6,'0')
   DO WHILE fso.FileExists(docname+'.doc')
    m.nfileid = m.nfileid + 1
    DocName = pmee+'\svacts\'+PADL(m.nfileid,6,'0')
   ENDDO 
   UPDATE svacts SET actname=PADL(m.nfileid,6,'0')+'.doc', actdate=DATETIME() WHERE recid = m.nfileid
   fso.CopyFile(pexpimp+'\'+m.actname, DocName+'.doc')
  ENDIF 
  
 ENDSCAN 
 USE IN svactsimp
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ƒŒ¡¿¬À≈ÕŒ '+ALLTRIM(STR(m.nActsAdded))+'¿ “Œ¬!'+CHR(13)+CHR(10),0+64,'')

RETURN 