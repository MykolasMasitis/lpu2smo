PROCEDURE ImpActs
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ »ÃœŒ–“»–Œ¬¿“‹ ¿ “€ › —œ≈–“»«?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ › —œŒ–“¿-»ÃœŒ–“¿'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 

 m.lIsNeedOldRecs=.f.
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€√–”∆¿“‹ –¿Õ≈≈ »ÃœŒ–“»–Œ¬¿ÕÕ€≈ ¿ “€?'+CHR(13)+CHR(10),4+32+256,'')==6
  m.lIsNeedOldRecs=.t.
 ENDIF 
 
 IF !fso.FolderExists(pMee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ Ã››'+CHR(13)+CHR(10)+'('+pmee+')!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\svacts')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ —¬ŒƒÕ€’ ¿ “Œ¬'+CHR(13)+CHR(10)+'('+pmee+'\SVACTS)!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\svacts\svacts.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ —œ–¿¬Œ◊Õ»  —¬ŒƒÕ€’ ¿ “Œ¬'+CHR(13)+CHR(10)+'('+pmee+'\SVACTS\SvActs.dbf)!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pMee+'\svacts\svacts', 'svacts', 'shar')>0
  IF USED('svacrs')
   USE IN svacts
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT svacts
 IF UPPER(FIELD('IsImp')) !=  'ISIMP'
  USE IN svacts
  IF OpenFile(pMee+'\svacts\svacts', 'svacts', 'excl')>0
   IF USED('svacts')
    USE IN svacts
   ENDIF 
   RETURN 
  ENDIF 
  SELECT svacts
  ALTER TABLE svacts ADD COLUMN IsImp L
  USE IN svacts
  IF OpenFile(pMee+'\svacts\svacts', 'svacts', 'shar')>0
   IF USED('svacts')
    USE IN svacts
   ENDIF 
   RETURN 
  ENDIF 
 ENDIF 
 
 CREATE TABLE &pExpImp\svactsimp ;
  (period c(6), mcod c(7), codexp n(1), actname c(25), actdate t, ok l)
 USE 
 =OpenFile(pExpImp+'\svactsimp.dbf', 'svactsimp', 'shar')>0

 SELECT svacts
 m.nCopiedActs = 0
 SCAN FOR IsImp = IIF(m.lIsNeedOldRecs=.f., .f., .t.)
  SCATTER MEMVAR 
  m.actname = ALLTRIM(m.actname)
  IF !fso.FileExists(pMee+'\svacts\'+m.actname)
   LOOP
  ENDIF 
  m.newname = ALLTRIM(SYS(2015))
  DO WHILE fso.FileExists(pExpImp+'\'+m.newname)
   m.newname = SYS(2015)
  ENDDO 
  fso.CopyFile(pMee+'\svacts\'+m.actname, pExpImp+'\'+m.newname, .t.)
  m.nCopiedActs = m.nCopiedActs + 1
  REPLACE IsImp WITH .t.
  
  m.actname = m.newname
  INSERT INTO svactsimp FROM MEMVAR 
  
 ENDSCAN 
 USE IN svacts
 
 m.nRecsInImpFile = RECCOUNT('svactsimp')
 
 USE IN svactsimp
 
 IF m.nRecsInImpFile==0
  fso.DeleteFile(pExpImp+'\svactsimp.dbf')
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'— Œœ»–Œ¬¿ÕŒ '+ALLTRIM(STR(m.nRecsInImpFile))+' ¿ “Œ¬'+CHR(13)+CHR(10),0+64,'')

RETURN 