PROCEDURE ImpExp
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÈÌÏÎÐÒÈÐÎÂÀÒÜ ÐÅÇÓËÜÒÀÒÛ ÝÊÑÏÅÐÒÈÇ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pExpImp)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÝÊÑÏÎÐÒÀ-ÈÌÏÎÐÒÀ'+CHR(13)+CHR(10)+pExpImp+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 

 iii = 1
 expimpbak = 'expimp'+PADL(iii,3,'0')
 DO WHILE fso.FileExists(pexpimp+'\'+expimpbak+'.zip')
  iii = iii + 1
  expimpbak = 'expimp'+PADL(iii,3,'0')
 ENDDO 
 ZipOpen(pexpimp+'\'+expimpbak+'.zip')
 m.nFilesInZip = 0

 FOR nmn=0 TO 12
  curmonth  = IIF(tmonth-nmn>0, tmonth-nmn, 12+tmonth-nmn)
  curyear   = IIF(tmonth-nmn>0, tYear, tYear-1)
  curperiod = STR(curyear,4)+PADL(curmonth,2,'0')

  IF !fso.FolderExists(pbase+'\'+curperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+curperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+curperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  WAIT curperiod WINDOW NOWAIT 

  SELECT aisoms
  SCAN 
   m.mcod = mcod 
   m.IsVed = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
   IF !fso.FolderExists(pbase+'\'+curperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   mfile = pbase+'\'+curperiod+'\'+m.mcod+'\m'+m.mcod
   IF !fso.FileExists(mfile+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(mfile, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    USE IN merror
    LOOP 
   ENDIF 
   USE IN merror
   
   m.impfile = 'i'+curperiod+m.mcod
   IF fso.FileExists(pExpImp+'\'+m.impfile+'.dbf')
    ZipFile(pExpImp+'\'+m.impfile+'.dbf')
    m.nFilesInZip = m.nFilesInZip + 1
    fso.DeleteFile(pExpImp+'\'+m.impfile+'.dbf')
   ENDIF 
   IF fso.FileExists(pExpImp+'\'+m.impfile+'.cdx')
    ZipFile(pExpImp+'\'+m.impfile+'.cdx')
    m.nFilesInZip = m.nFilesInZip + 1
    fso.DeleteFile(pExpImp+'\'+m.impfile+'.cdx')
   ENDIF 
   
   fso.CopyFile(mfile+'.dbf', pExpImp+'\'+m.impfile+'.dbf')
   fso.CopyFile(mfile+'.cdx', pExpImp+'\'+m.impfile+'.cdx')

   SELECT aisoms 
  ENDSCAN 
  USE IN aisoms
  
  WAIT CLEAR 
 
 NEXT 

 ZipClose()
 IF m.nFilesInZip=0
  fso.DeleteFile(pexpimp+'\'+expimpbak+'.zip')
 ENDIF 
 
RETURN 

