PROCEDURE FillFilId
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ «¿œŒÀÕ»“‹ œŒÀ≈ LPU_ID?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF tmonth!=12 OR tyear!=2012
  MESSAGEBOX(CHR(13)+CHR(10)+'ÃŒƒ”À‹ œ–≈ƒÕ¿«Õ¿◊≈Õ “ŒÀ‹ Œ ƒÀﬂ ŒƒÕŒ√Œ œ≈–»Œƒ¿!'+CHR(13)+CHR(10);
   ,0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(10)+'ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿ Œ“—”“—“¬”≈“!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿ Œ“—”“—“¬”≈“!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF  
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF
 ENDIF 
 
 m.mmy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\b'+m.mcod+'.'+m.mmy)
   LOOP 
  ENDIF 

  ffile = fso.GetFile(pbase+'\'+gcperiod+'\'+m.mcod+'\b'+m.mcod+'.'+m.mmy)
  IF ffile.size >= 2
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(2)
   fhandl.Close
  ELSE 
   lcHead = ''
   LOOP 
  ENDIF 

  IF lcHead!='PK' && ›ÚÓ zip-Ù‡ÈÎ!
   LOOP 
  ENDIF 

  ZipName = pbase+'\'+gcperiod+'\'+m.mcod+'\b'+m.mcod+'.'+m.mmy
  IsSuccess=UnzipOpen(ZipName)
  IF IsSuccess
   llIsOneZip = .t.
   UnzipClose()
  ELSE 
   UnzipClose()
   LOOP 
  ENDIF 

  lIsFileInArc = .f.
  UnzipOpen(ZipName)
  sItem   = 'D79S' + m.qcod + '.' + m.mmy
  IF UnzipGotoFileByName(sItem)
   lIsFileInArc = .t.
   UnzipClose()
  ELSE 
   UnzipClose()
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP
  ENDIF 
  IF OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT talon
  IF FIELD('fil_id')!='FIL_ID'
   USE IN talon
   OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\talon', 'talon', 'excl')
   SELECT talon
   ALTER TABLE talon ADD COLUMN fil_id n(6)
   USE 
   OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')
  ENDIF 
  
  inDir   = pbase+'\'+gcperiod+'\'+m.mcod
  ZipDir  = InDir + '\'

  UnzipOpen(ZipName)
  UnzipGotoFileByName(sItem)
  UnzipFile(ZipDir)
  UnzipClose()
  
  =OpenFile(pBase+'\'+gcperiod+'\'+m.mcod+'\'+sItem, 'sItem', 'excl')
  SELECT sItem
  INDEX on recid TAG recid 
  SET ORDER TO recid

  SELECT talon 
  SET RELATION TO LEFT(recid_lpu,6) INTO sItem
  SCAN 
   IF !EMPTY(sItem.fil_id)
    REPLACE fil_id WITH sItem.fil_id
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO sItem
  USE
  
  SELECT sItem
  SET ORDER TO 
  DELETE TAG ALL 
  USE 
  
  fso.DeleteFile(pBase+'\'+gcperiod+'\'+m.mcod+'\'+sItem)
  
*  MESSAGEBOX(m.mcod,0+64,'')
  
 ENDSCAN 
 USE 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 