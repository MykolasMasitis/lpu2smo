PROCEDURE ImpS2Me
 IF MESSAGEBOX('¬€ ’Œ“»“≈ »ÃœŒ–“»–Œ¬¿“‹ ME-Ù‡ÈÎ˚?'+CHR(13)+CHR(10),4+32,'ImpS2Me')=7
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
 IF !fso.FileExists(pmee+'\svacts\svacts.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À '+UPPER(pmee+'\svacts\svacts.dbf'),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\mee2mgf.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À '+UPPER(pcommon+'\mee2mgf.dbf'),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\pnyear', 'pnyear', 'shar', 'period')>0
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pmee+'\svacts\svacts', 'svacts', 'shar', 'unik')>0
  IF USED('svacts')
   USE IN svacts
  ENDIF 
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\mee2mgf', 'mee2mgf', 'shar', 'mgf_et')>0
  IF USED('mee2mgf')
   USE IN mee2mgf
  ENDIF 
  IF USED('svacts')
   USE IN svacts
  ENDIF 
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
  IF !INLIST(ext, 'dbf')
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
  m.fname = fname 
  INSERT INTO curdbf (fname) VALUES (m.fname)
 ENDSCAN 
 USE 
 
 CREATE CURSOR curperiod (period c(6), me_s_all n(11,2), ss_all n(11,2))
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
   CREATE CURSOR curdefs FROM ARRAY svm
   ALTER TABLE curdefs ADD COLUMN "comment" c(100)
   RELEASE svm 
  ENDIF 
  
  SELECT mefile 
  SCAN 
*   SCATTER FIELDS EXCEPT recid MEMVAR 
   SCATTER MEMVAR 
   IF !SEEK(m.period_e, 'curperiod')
    INSERT INTO curperiod (period) VALUES (m.period_e)
   ENDIF 
   
   INSERT INTO svmee FROM MEMVAR 
   
  ENDSCAN 
  USE 
  SELECT curdbf
  
 ENDSCAN 
 USE

 SELECT curperiod
 
 SCAN 
  m.lcperiod = period
  m.ss_all   = 0 
  m.me_s_all = 0
  IF !SEEK(LEFT(m.lcperiod,4), 'pnyear')
   MESSAGEBOX('¬ ‘¿…À≈ '+UPPER(pcommon+'\pnyear.dbf')+CHR(13)+CHR(10)+;
    'Õ≈ Õ¿…ƒ≈ÕŒ «Õ¿◊≈Õ»≈ ÕŒ–Ã¿“»¬¿ Õ¿'+m.lcperiod+'!'+CHR(13)+CHR(10)+;
    'œ≈–»Œƒ œ–Œ»√ÕŒ–»–Œ¬¿Õ!'+CHR(13)+CHR(10),0+64,'')
   LOOP 
  ELSE 
   m.pnorm = pnyear.pnorm
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   SELECT * FROM svmee WHERE period_e=m.lcperiod ORDER BY lpu_id INTO CURSOR onemee
   m.comment = 'ŒÚÒÛÚÒÚ‚ÛÂÚ ‰ËÂÍÚÓËˇ ÔÂËÓ‰‡ '+m.lcperiod
   SELECT onemee
   SCAN 
    SCATTER MEMVAR 
    INSERT INTO curdefs FROM MEMVAR 
   ENDSCAN 
   USE IN onemee
   SELECT curperiod
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   SELECT * FROM svmee WHERE period_e=m.lcperiod ORDER BY lpu_id INTO CURSOR onemee
   m.comment = 'ŒÚÒÛÚÒÚ‚ÛÂÚ Ù‡ÈÎ '+m.lcperiod+'\aisoms.dbf'
   SELECT onemee
   SCAN 
    SCATTER MEMVAR 
    INSERT INTO curdefs FROM MEMVAR 
   ENDSCAN 
   USE IN onemee
   SELECT curperiod
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
  
  SELECT onemee
  m.omcod = ''
  m.lnsum = 0 
  SCAN 
   m.lpu_id = lpu_id
   m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')

   IF m.omcod != m.mcod
    m.omcod = m.mcod 
    m.lnsum = 0 

    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('merror')
     USE IN merror
    ENDIF 
    IF USED('eerror')
     USE IN eerror
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
    IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'eerror', 'shar', 'rid')>0
     IF USED('talon')
      USE IN talon
     ENDIF 
     IF USED('eerror')
      USE IN eerror
     ENDIF 
     SELECT onemee
     m.omcod = ''
     LOOP 
    ENDIF 
    IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'excl', 'recid')>0
     IF USED('talon')
      USE IN talon
     ENDIF 
     IF USED('merror')
      USE IN merror
     ENDIF 
     IF USED('eerror')
      USE IN eerror
     ENDIF 
     SELECT onemee
     m.omcod = ''
     LOOP 
    ENDIF 
    
    SELECT merror
*    ZAP 
    
    SELECT talon 
    SET RELATION TO recid INTO eerror

   ENDIF 
   
   SELECT onemee
   m.act = ''
   m.d_a = DATETIME()
   SCATTER MEMVAR 
   RELEASE recid, period 

   m.otd     = m.iotd 
   m.cod     = IIF(EMPTY(m.cod), m.cod_e, m.cod)
   m.k_u     = IIF(EMPTY(m.k_u), m.k_u_e, m.k_u)
   m.vir     = m.sn_pol+m.otd+m.ds+PADL(m.cod,6,"0")+DTOC(m.d_u)
   m.et      = IIF(EMPTY(m.et),'2',m.et)
      
   IF !SEEK(m.vir, 'talon')
    SCATTER MEMVAR 
*    MESSAGEBOX(m.vir,0+64,'')
    m.comment = '«‡ÔËÒ¸ ÌÂ Ò‚ˇÁ‡Î‡Ò¸ Ò talon.dbf'
    INSERT INTO curdefs FROM MEMVAR 
    LOOP 
   ENDIF 
   
   IF !EMPTY(eerror.c_err) AND !INLIST(m.et,'F','R')
    SCATTER MEMVAR 
    m.comment = '«‡ÔËÒ Á‡·‡ÍÓ‚‡Ì‡ ÔÓ Ã›  '+eerror.c_err
    INSERT INTO curdefs FROM MEMVAR 
    LOOP 
   ENDIF 
      
    m.recid = talon.recid

    m.e_cod    = m.cod_e
    m.e_ku     = m.k_u_e
    m.e_tip    = m.tip_e
    m.err_mee  = IIF(m.er_c='99','W0',m.er_c)
    m.osn230   = ALLTRIM(m.osn230)
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
*     IF m.cod = m.e_cod AND m.k_u = m.e_ku AND m.tip = m.e_tip
      m.ee    = '1'
      m.koeff = ROUND(m.s_1/m.s_all,2)
      m.e_cod = 0 
      m.e_ku  = 0
      m.e_tip = ''
*     ELSE 
*      m.ee    = '2'
*      m.koeff = 0 
*     ENDIF 
    ENDIF 
    
    m.codexp = IIF(SEEK(m.et, 'mee2mgf'), INT(VAL(mee2mgf.my_et)), 0)

    IF !ISALPHA(m.et)
    m.svactid = m.lcperiod + m.e_period + m.mcod+STR(m.codexp,1)+SPACE(7)
    IF !SEEK(m.svactid, 'svacts')
     INSERT INTO svacts (period,mcod,lpu_id,codexp,e_period,et,actname,actdate,s_all,s_exp) VALUES ;
      (m.lcperiod,m.mcod,m.lpu_id,m.codexp,m.e_period,m.et,m.act,m.d_a,m.s_all,m.s_1)
    ELSE 
     m.rriid = svacts.recid
     IF m.s_1>0
      m.o_s_1 = svacts.s_exp
      m.n_s_1 = m.o_s_1 + m.s_1
      m.o_s_all = svacts.s_all
      m.n_s_all = m.o_s_all + m.s_all
      UPDATE svacts SET s_exp = m.n_s_1, s_all = m.n_s_all WHERE recid = m.rriid
     ENDIF 
    ENDIF 
    ENDIF 

    IF !SEEK(m.recid, 'merror')
     INSERT INTO merror FROM MEMVAR 
     m.ss_all = m.ss_all + m.s_1
    ELSE  
     SKIP IN talon 
     IF talon.sn_pol+talon.otd+talon.ds+PADL(talon.cod,6,"0")+DTOC(talon.d_u) = m.vir
      m.recid = talon.recid
      INSERT INTO merror FROM MEMVAR 
      m.ss_all = m.ss_all + m.s_1
     ELSE 
      INSERT INTO merror FROM MEMVAR 
      m.ss_all = m.ss_all + m.s_1
*      SCATTER MEMVAR 
*      m.comment = 'ƒÛ·Î¸ sn_pol+otd+ds+PADL(cod,6,"0")+DTOC(d_u)'
*      INSERT INTO curdefs FROM MEMVAR 
     ENDIF 
    ENDIF 


  ENDSCAN 

  IF USED('talon')
   SET RELATION OFF INTO eerror
   USE IN talon
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('eerror')
   USE IN eerror
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
 
 SELECT svmee 
 USE 
 
 USE IN pnyear
 USE IN svacts
 USE IN mee2mgf
 
 SELECT curperiod 
 m.cfile = 'mee'+m.e_period
 COPY TO pUpdDir+'\'+m.cfile
 USE IN curperiod

 SELECT curdefs
 COPY TO pUpdDir+'\curdefs'
 USE 

 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!',0+64,'')

RETURN 
