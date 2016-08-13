PROCEDURE ExpExp
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ «¿√–”«»“‹ –≈«”À‹“¿“€ › —œ≈–“»«?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ › —œŒ–“¿-»ÃœŒ–“¿'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ Ã››'+CHR(13)+CHR(10)+pMee+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pMee+'\SVACTS')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ SVACTS'+CHR(13)+CHR(10)+pMee+'\SVACTS'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pMee+'\SVACTS\svacts.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À SVACTS.DBF'+CHR(13)+CHR(10)+pMee+'\SVACTS\svacts.dbf'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'shar', 'unik')>0
  IF USED('svacts')
   USE IN svacts
  ENDIF 
  RETURN 
 ENDIF 

 oMailDir        = fso.GetFolder(pExpImp)
 MailDirName     = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count
 
 m.nGoodFiles = 0
 m.nGoodRecs  = 0 
 m.nProcFiles = 0
 m.nProcRecs  = 0

 FOR EACH oFileInMailDir IN oFilesInMailDir
  m.BFullName = oFileInMailDir.Path
  m.bname     = oFileInMailDir.Name
  m.recieved  = oFileInMailDir.DateLastModified
  
  IF LEN(m.bname)!=18
   LOOP 
  ENDIF 
  
  m.part01 = LEFT(m.bname,1)
  m.part02 = SUBSTR(m.bname,2,6)
  m.part03 = SUBSTR(m.bname,8,7)
  m.ext    = LOWER(RIGHT(m.bname,3))

  IF part01 != 'i'
   LOOP 
  ENDIF 
  IF !INLIST(LEFT(part02,4), STR(tYear-1,4), STR(tYear,4))
   LOOP 
  ENDIF 
  IF !INLIST(SUBSTR(part02,5,2), '01','02','03','04','05','06','07','08','09','10','11','12')
   LOOP 
  ENDIF 
  IF !INLIST(ext, 'dbf')
   LOOP 
  ENDIF 
  
  curperiod = part02
  curmcod   = part03
  
  IF !fso.FolderExists(pbase+'\'+curperiod)
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+curperiod+'\nsi')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+curperiod+'\nsi\sprlpuxx.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+curperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   LOOP 
  ENDIF 
  IF !SEEK(curmcod, 'sprlpu')
   USE IN sprlpu
   LOOP 
  ENDIF 
  USE IN sprlpu 
  
  IF !fso.FolderExists(pbase+'\'+curperiod+'\'+curmcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+curperiod+'\'+curmcod+'\m'+curmcod+'.'+m.ext)
   LOOP 
  ENDIF 
  
  m.nGoodFiles = m.nGoodFiles + 1
  
  IF OpenFile(m.BFullName, 'impfile', 'shar')>0
   IF USED('impfile')
    USE IN impfile
   ENDIF
   LOOP  
  ENDIF 
  
  IF OpenFile(pbase+'\'+curperiod+'\'+curmcod+'\m'+curmcod, 'merror', 'shar', 'unik')>0
   IF USED('merror')
    USE IN merror 
   ENDIF 
   USE IN impfile
   LOOP 
  ENDIF 
  
  SELECT impfile
  SCAN 
   SCATTER FIELDS EXCEPT rid MEMVAR 
   
   m.unvir = m.curperiod + m.e_period + m.curmcod + m.et + m.docexp
   IF !SEEK(m.unvir, 'svacts')
    INSERT INTO svacts (period,e_period,mcod,codexp,docexp) VALUES (m.curperiod,m.e_period,m.curmcod,INT(VAL(m.et)),m.docexp)
   ENDIF 
   
   m.vvir = PADL(m.recid,6,"0")+m.et+LEFT(m.err_mee,2)
   IF SEEK(m.vvir, 'merror')
    IF merror.e_period!=m.e_period
     UPDATE merror SET e_period=m.e_period WHERE PADL(merror.recid,6,"0")+merror.et+LEFT(merror.err_mee,2)=m.vvir
    ENDIF 

    LOOP 

   ENDIF 
   
   INSERT INTO merror FROM MEMVAR 
     
  ENDSCAN 
  
  USE IN merror
  USE IN impfile

 ENDFOR 

 USE IN svacts
 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“¿ÕŒ '+ALLTRIM(STR(m.nGoodFiles))+' ‘¿…ÀŒ¬'+;
 CHR(13)+CHR(10),0+64,'')

RETURN 