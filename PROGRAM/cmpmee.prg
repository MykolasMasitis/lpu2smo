PROCEDURE CmpMee
 IF MESSAGEBOX('ÏÐÎÂÅÐÈÒÜ ÊÎÐÐÅÊÒÍÎÑÒÜ ÑÍßÒÈÉ?'+CHR(13)+CHR(10),4+16,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pOut+'\'+m.gcPeriod)
  MESSAGEBOX('ME-ÔÀÉËÛ ÍÅ ÑÔÎÐÌÈÐÎÂÀÍÛ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR cmpres (lpuid n(4), mcod c(7), apsf n(11,2), mesum n(11,2))
 INDEX ON lpuid TAG lpuid
 INDEX ON mcod TAG mcod 

 oMailDir        = fso.GetFolder(pOut+'\'+m.gcPeriod)
 MailDirName     = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count

 FOR EACH oFileInMailDir IN oFilesInMailDir
  m.BFullName = oFileInMailDir.Path
  m.bname     = oFileInMailDir.Name
  m.recieved  = oFileInMailDir.DateLastModified
  
  IF LEN(m.bname)!=12
   LOOP 
  ENDIF 
  
  m.part01 = UPPER(LEFT(m.bname,2))
  m.part02 = UPPER(SUBSTR(m.bname,3,2))
  m.part03 = SUBSTR(m.bname,5,4)
  m.ext    = LOWER(RIGHT(m.bname,3))

  IF part01 != 'ME'
   LOOP 
  ENDIF 
  IF part02 != m.qcod
   LOOP 
  ENDIF 
  IF !INLIST(ext, 'dbf', 'zip')
   LOOP 
  ENDIF 
  IF !SEEK(INT(VAL(m.part03)),'aisoms')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.BFullName, 'mfile', 'shar')>0
   IF USED('mfile')
    USE IN mfile
   ENDIF 
   LOOP 
  ELSE 
   SELECT mfile 
   SUM s_opl_e TO m.mesum
   USE IN mfile 
  ENDIF 

  m.lpuid = INT(VAL(m.part03))
  m.mcod  = aisoms.mcod
  m.apsf  = aisoms.e_mee + aisoms.e_ekmp
  
  INSERT INTO cmpres (lpuid,mcod,apsf,mesum) VALUES (m.lpuid,m.mcod,m.apsf,m.mesum)
  
 ENDFOR 
 
 SELECT cmpres
 IF RECCOUNT()<=0 
  USE IN aisoms
  USE IN cmpres
  MESSAGEBOX('ÍÅ ÎÁÍÀÐÓÆÅÍÎ ÍÈ ÎÄÍÎÃÎ ME-ÔÀÉËÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ELSE 
  SELECT cmpres
  COPY TO &pout\&gcperiod\cmpfile
  USE 
  USE IN aisoms
  MESSAGEBOX('ÔÀÉË ÑÐÀÂÍÅÍÈß ÑÔÎÐÌÈÐÎÂÀÍ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
  
RETURN 