PROCEDURE ImpS6Me
 IF MESSAGEBOX('¬€ ’Œ“»“≈ »ÃœŒ–“»–Œ¬¿“‹ ME-Ù‡ÈÎ˚?'+CHR(13)+CHR(10),4+32,'ImpS6Me')=7
  RETURN 
 ENDIF 

 pUpdDir = fso.GetParentFolderName(pbin)+'\MEFILES'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 
 
 IF !fso.FileExists(pcommon+'\pnyear.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À '+UPPER(pcommon+'\pnyear.dbf'),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\pnyear', 'pnyear', 'shar', 'period')>0
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
  RETURN 
 ENDIF 

 oMailDir        = fso.GetFolder(pUpdDir)
 MailDirName     = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count
 
 IF nFilesInMailDir<=0
  RELEASE oMailDir, MailDirName, oFilesInMailDir, nFilesInMailDir
  MESSAGEBOX('¬ ƒ»–≈ “Œ–»» '+UPPER(ALLTRIM(pUpdDir))+CHR(13)+CHR(10)+'Õ≈ Œ¡Õ¿–”∆≈ÕŒ Õ» ŒƒÕŒ√Œ ‘¿…À¿!',0+64,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curmes (fname c(12))

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
  IF !INLIST(ext, 'zip')
   LOOP 
  ENDIF 
  
  oEFile = fso.GetFile(m.BFullName)
  IF oEFile.size >= 2
   fhandl = oEFile.OpenAsTextStream
   lcHead = fhandl.Read(2)
   fhandl.Close
  ELSE 
   lcHead = ''
  ENDIF 

  IF lcHead != 'PK' && ›ÚÓ zip-Ù‡ÈÎ!
   LOOP 
  ENDIF 

  UnzipOpen(m.BFullName)
  FilesInZip = UnZipFileCount()
  UnzipClose()
 
  IF FilesInZip != 1
   LOOP 
  ENDIF 

  INSERT INTO curmes (fname) VALUES (m.bname)

 ENDFOR 
 
 SELECT curmes
 IF RECCOUNT('curmes')<=0
  USE IN curmes
  MESSAGEBOX('»« '+ALLTRIM(STR(nFilesInMailDir))+' Œ¡Õ¿–”∆≈ÕÕ€’ ¬ ƒ»–≈ “Œ–»»'+CHR(13)+CHR(10)+;
   UPPER(ALLTRIM(pUpdDir))+' ‘¿…ÀŒ¬'+CHR(13)+CHR(10)+'Õ≈“ Õ» ŒƒÕŒ√Œ ‘¿…À¿ ME!',0+64,'')
  RELEASE oMailDir, MailDirName, oFilesInMailDir, nFilesInMailDir
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdbf (fname c(12))
 
 SELECT curmes
 SCAN 
  m.fname = ALLTRIM(fname)
  m.fpath = pUpdDir+'\'+m.fname
  
  IF UnzipOpen(m.fpath)==.T.
   FilesInZip = UnZipFileCount()
   DIMENSION ZipArray(13)
   ZipArray = ''
   UnZipGotoTopFile()
   FOR FileInZip=0 TO FilesInZip-1
    UnzipAFileInfo("ZipArray")
    m.FileInZipName = ALLTRIM(ZipArray(1))
    IF STRTRAN(UPPER(ALLTRIM(m.FileInZipName)),'.DBF','.ZIP') != UPPER(ALLTRIM(m.fname))
     LOOP 
    ENDIF 

    INSERT INTO curdbf (fname) VALUES (m.FileInZipName)
    
    UnzipSetFolder(pUpdDir)
    IF fso.FileExists(pUpdDir+'\'+m.FileInZipName)
     fso.DeleteFile(pUpdDir+'\'+m.FileInZipName)
    ENDIF 
    UnzipByIndex(FileInZip)
    UnzipGotoNextFile()
   ENDFOR  
   UnzipClose()
  ENDIF 
  
 ENDSCAN 
 USE
 
 CREATE CURSOR curperiod (period c(6), ss_all n(11,2))
 INDEX on period TAG period 
 SET ORDER TO period 
 
 SELECT curdbf
 SCAN 
  m.fname = ALLTRIM(fname)
  m.fpath = pUpdDir+'\'+m.fname
  
  IF !fso.FileExists(m.fpath)
   LOOP 
  ENDIF 
  
  IF OpenFile(m.fpath, 'mefile', 'shar')>0
   IF USED('mefile')
    USE IN mefile
   ENDIF 
   SELECT curdbf
   LOOP 
  ENDIF 
  
  IF !USED('svmee')
   nCounts = AFIELDS(svm, 'mefile')
   CREATE CURSOR svmee FROM ARRAY svm
   RELEASE svm 
  ENDIF 
  
  SELECT mefile 
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   IF !SEEK(m.period_e, 'curperiod')
    INSERT INTO curperiod (period) VALUES (m.period_e)
   ENDIF 
   
   INSERT INTO svmee FROM MEMVAR 
   
  ENDSCAN 
  USE 
  SELECT curdbf
  
 ENDSCAN 
 USE

 RELEASE oMailDir, MailDirName, oFilesInMailDir, nFilesInMailDir
 
 SELECT curperiod
 
* BROWSE 
 
 IF RECCOUNT('curperiod')<=0
  USE IN curperiod
  USE IN svmee 
  MESSAGEBOX('Õ≈ –¿—œŒ«Õ¿Õ Õ» Œƒ»Õ œ≈–»Œƒ!',0+16,'')
  RETURN 
 ENDIF 
 
 SCAN 
  m.lcperiod = period
  m.ss_all = 0 
  IF !SEEK(LEFT(m.lcperiod,4), 'pnyear')
   MESSAGEBOX('¬ ‘¿…À≈ '+UPPER(pcommon+'\pnyear.dbf')+CHR(13)+CHR(10)+;
    'Õ≈ Õ¿…ƒ≈ÕŒ «Õ¿◊≈Õ»≈ ÕŒ–Ã¿“»¬¿ Õ¿'+m.lcperiod+'!'+CHR(13)+CHR(10)+;
    'œ≈–»Œƒ œ–Œ»√ÕŒ–»–Œ¬¿Õ!'+CHR(13)+CHR(10),0+64,'')
   LOOP 
  ELSE 
   m.pnorm = pnyear.pnorm
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\nsi\sprlpuxx.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   SELECT curperiod
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   SELECT curperiod
   LOOP 
  ENDIF 
  
  WAIT m.lcperiod +'...' WINDOW NOWAIT 

  SELECT * FROM svmee WHERE period_e=m.lcperiod ORDER BY lpu_id INTO CURSOR onemee
  
*  IF m.lcperiod = '201402'
*   SET STEP ON ON 
*  ENDIF 
  
  SELECT onemee
  m.omcod = ''
  SCAN 
   m.lpu_id = lpu_id
   m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')

   IF m.omcod != m.mcod
    m.omcod = m.mcod 

    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('merror')
     USE IN merror
    ENDIF 

    IF EMPTY(m.mcod)
     m.omcod = ''
     LOOP 
    ENDIF 
    IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\'+m.mcod)
     m.omcod = ''
     LOOP 
    ENDIF 
    IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
     m.omcod = ''
     LOOP 
    ENDIF 
    IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
     m.omcod = ''
     LOOP 
    ENDIF 
    IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'unik')>0
     IF USED('talon')
      USE IN talon
     ENDIF 
     SELECT onemee
     m.omcod = ''
     LOOP 
    ENDIF 
    IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'recid')>0
     IF USED('talon')
      USE IN talon
     ENDIF 
     IF USED('merror')
      USE IN merror
     ENDIF 
     SELECT onemee
     m.omcod = ''
     LOOP 
    ENDIF 

   ENDIF 
   
   SELECT onemee
   SCATTER MEMVAR 
   RELEASE recid, period 

   m.otd     = m.iotd 
   m.vir     = m.sn_pol+m.otd+m.ds+PADL(m.cod,6,"0")+DTOC(m.d_u)
   m.et      = IIF(EMPTY(m.et),'2',m.et)
      
   IF SEEK(m.vir, 'talon')
    m.recid = talon.recid

    m.e_cod    = m.cod_e
    m.e_ku     = m.k_u_e
    m.e_tip    = m.tip_e
    m.err_mee  = IIF(m.er_c='99','W0',m.er_c)
    m.osn230   = ALLTRIM(m.osn230)
*    m.e_period = m.lcperiod
    m.e_period = period
    m.koeff    = 0 
    m.s_all    = talon.s_all
    m.s_1      = m.s_opl_e
    m.s_2      = m.s_sank
    m.straf    = ROUND(m.s_2/m.pnorm,2)
    m.impdata  = DATE()

    IF m.err_mee = 'W0'
     m.e_cod = 0 
     m.e_ku  = 0
     m.e_tip = ''
    ELSE 
     IF m.cod = m.e_cod AND m.k_u = m.e_ku AND m.tip = m.e_tip
      m.ee    = '1'
      m.koeff = ROUND(m.s_1/m.s_all,2)
      m.e_cod = 0 
      m.e_ku  = 0
      m.e_tip = ''
     ELSE 
      m.ee    = '2'
      m.koeff = 0 
     ENDIF 
    ENDIF 

    IF !SEEK(m.recid, 'merror')
     && ‰Ó·‡‚ÎˇÂÏ ‚ Ù‡ÈÎ Ó¯Ë·ÓÍ Á‡ÔËÒË!!!
     INSERT INTO merror FROM MEMVAR 
     m.ss_all = m.ss_all + m.s_1
    ELSE 
     UPDATE merror WHERE recid=m.recid SET e_period=m.e_period
     BROWSE 
    ENDIF 
   ENDIF 

  ENDSCAN 

  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 

  USE IN onemee
  USE IN aisoms 
  USE IN sprlpu 

  SELECT curperiod
  IF ss_all != m.ss_all
   REPLACE ss_all WITH m.ss_all
  ENDIF 
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 SELECT svmee 
 USE 
 
 USE IN pnyear
 
 SELECT curperiod 
 GO TOP 
 BROWSE 
 USE IN curperiod
 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!',0+64,'')

RETURN 
