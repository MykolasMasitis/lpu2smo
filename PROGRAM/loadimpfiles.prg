PROCEDURE LoadImpFiles

 IF MESSAGEBOX(CHR(13)+CHR(10)+'«¿√–”«»“‹ ƒ¿ÕÕ€≈ œŒ —Õﬂ“»ﬂÃ ¬ ¿œ—‘?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(13)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+pbase+'\'+gcperiod+'!'+CHR(13)+CHR(10),0+16,'') 
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(13)+'Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'') 
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ › —œŒ–“¿-»ÃœŒ–“¿'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
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

  m.nGoodFiles = m.nGoodFiles + 1
  
  IF OpenFile(m.BFullName, 'impfile', 'shar')>0
   IF USED('impfile')
    USE IN impfile
   ENDIF
   LOOP  
  ENDIF 
  
  m.emee  = 0
  m.eekmp = 0
  SELECT impfile
  SCAN 
   m.e_period = e_period
   IF m.e_period!=m.gcperiod
    LOOP 
   ENDIF 
   m.et   = et
   m.sexp = s_1
   m.emee  = m.emee  + IIF(INLIST(m.et,'2','3','7'), m.sexp, 0)
   m.eekmp = m.eekmp + IIF(INLIST(m.et,'4','5','6'), m.sexp, 0)
  ENDSCAN 
  USE IN impfile

  SELECT aisoms
  SEEK curmcod
  REPLACE e_mee WITH m.emee, e_ekmp WITH m.eekmp

 ENDFOR 

 WAIT CLEAR 
 
 USE IN aisoms

 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“¿ÕŒ '+ALLTRIM(STR(m.nGoodFiles))+' ‘¿…ÀŒ¬'+;
 CHR(13)+CHR(10),0+64,'')
  
RETURN 