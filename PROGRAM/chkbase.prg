FUNCTION chkBase

IsArcExists = fso.FolderExists(pArc)
IF IsArcExists == .F.
 IF MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ" + CHR(13) + "&pArc!" + CHR(13) + "—Œ«ƒ¿“‹?",4+32, "") == 6
  fso.CreateFolder(pArc)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase)
IF IsDirExists == .F.
 IF MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ" + CHR(13) + "&pBase!" + CHR(13) + "—Œ«ƒ¿“‹?",4+32, "") == 6
  fso.CreateFolder(pBase)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase+'\'+gcPeriod)
IF IsDirExists == .F.
 IF MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ" + CHR(13) + "&pBase\&gcPeriod!" + CHR(13) + "—Œ«ƒ¿“‹?",4+32, "") == 6
  fso.CreateFolder(pBase+'\'+gcPeriod)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pDouble)
IF IsDirExists == .F.
 fso.CreateFolder(pDouble)
ENDIF 

IsDirExists = fso.FolderExists(pMee)
IF IsDirExists == .F.
 fso.CreateFolder(pMee)
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\SVACTS')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\SVACTS')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\SSACTS')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\SSACTS')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\REQUESTS')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\REQUESTS')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\TUNES')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\TUNES')
ENDIF 

IsDirExists = fso.FolderExists(pMee+'\BLANK')
IF IsDirExists == .F.
 fso.CreateFolder(pMee+'\BLANK')
ENDIF 

IsDirExists = fso.FolderExists(pOut)
IF IsDirExists == .F.
 IF MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ" + CHR(13) + "&pOut!" + CHR(13) + "—Œ«ƒ¿“‹?",4+32, "") == 6
  fso.CreateFolder(pOut)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pTempl)
IF IsDirExists == .F.
 IF MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ" + CHR(13) + "&pTempl!" + CHR(13) + "—Œ«ƒ¿“‹?",4+32, "") == 6
  fso.CreateFolder(pTempl)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pTrash)
IF IsDirExists == .F.
 fso.CreateFolder(pTrash)
ENDIF 

IsDirExists = fso.FolderExists(pExpImp)
IF IsDirExists == .F.
 fso.CreateFolder(pExpImp)
ENDIF 

*IF !fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee.dbf')
* MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À '+CHR(13)+CHR(10)+pbase+'\'+gcperiod+'\'+'NSI'+'\ERRSMEE.DBF'+CHR(13)+CHR(10),0+64,'')
*ENDIF 

IF !fso.FileExists(pcommon+'\dspcodes.dbf')
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À '+CHR(13)+CHR(10)+pcommon+'\DSPCODES.DBF'+CHR(13)+CHR(10),0+64,'')
ENDIF 

IF !fso.FileExists(pcommon+'\mo_vmp_2014.dbf')
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À '+CHR(13)+CHR(10)+pcommon+'\MO_VMP_2014.DBF'+CHR(13)+CHR(10),0+64,'')
ENDIF 

IF !fso.FileExists(pcommon+'\spi_lpu_dd_2014.dbf')
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À '+CHR(13)+CHR(10)+pcommon+'\SPI_LPU_DD_2014.DBF'+CHR(13)+CHR(10),0+64,'')
ENDIF 

*IF !fso.FileExists(pcommon+'\tpn.dbf')
* MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À '+CHR(13)+CHR(10)+pcommon+'\TPN.DBF'+CHR(13)+CHR(10),0+64,'')
*ENDIF 

IsFileExists = fso.FileExists(pCommon+'\UsrLpu.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\UsrLpu (mcod c(7), lpu_id n(4), cokr c(2), usr n(2)) 
 APPEND FROM &pCommon\sprlpuxx 
 REPLACE ALL usr WITH 1
* INDEX ON mcod TAG mcod 
* INDEX ON lpu_id TAG lpu_id
 USE 
ENDIF 

IsFileExists = fso.FileExists(pCommon+'\Users.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\Users (RecId i AUTOINC NEXTVALUE 1 STEP 1, "name" c(6), ;
  fam c(25), im c(25), ot c(25), fio c(40), super l, usrmail c(6)) 
 INDEX ON recid TAG recid CANDIDATE 
 INDEX on name TAG name 
 
 INSERT INTO Users ("name",fam,im,ot,fio,super,usrmail) VALUES ;
  ('OMS','–ˇ·Ó‚','ÃËı‡ËÎ','—Ú‡ÌËÒÎ‡‚Ó‚Ë˜','–ˇ·Ó‚ Ã.—.',.t.,'USR010')
 USE
ENDIF 

IF fso.FileExists(pCommon+'\prv002xx.dbf')
 IF OpenFile(pCommon+'\prv002xx', 'profss', 'shar')>0
  IF USED('profss')
   USE IN profss
  ENDIF 
 ELSE 
  SELECT profss
  IF FIELD('isoms')!='ISOMS'
   USE IN profss
   IF OpenFile(pCommon+'\prv002xx', 'profss', 'excl')>0
   ELSE 
    SELECT profss
    ALTER table profss ADD COLUMN isoms l
    REPLACE ALL isoms WITH .t.
   ENDIF 
  ENDIF 
  IF USED('profss')
   USE IN profss
  ENDIF 
 ENDIF 
ENDIF 

IF fso.FileExists(pMee+'\SVACTS\svacts.dbf')
 IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'shar')>0
  IF USED('svacts')
   USE IN svacts
  ENDIF 
 ELSE 
  SELECT svacts
  IF FIELD('flcod')!='FLCOD'
   USE IN svacts
   IF OpenFile(pMee+'\SVACTS\svacts', 'svacts', 'excl')>0
   ELSE 
    SELECT svacts
    ALTER TABLE svacts ADD COLUMN flcod c(12)
   ENDIF 
  ENDIF 
  IF USED('svacts')
   USE IN svacts
  ENDIF 
 ENDIF 
ENDIF 

IF fso.FileExists(pMee+'\ssacts\ssacts.dbf')
 IF OpenFile(pMee+'\ssacts\ssacts', 'ssacts', 'shar')>0
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
 ELSE 
  SELECT ssacts
  IF FIELD('flcod')!='FLCOD'
   USE IN ssacts
   IF OpenFile(pMee+'\ssacts\ssacts', 'ssacts', 'excl')>0
   ELSE 
    SELECT ssacts
    ALTER TABLE ssacts ADD COLUMN flcod c(12)
   ENDIF 
  ENDIF 
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
 ENDIF 
ENDIF 

IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee.dbf')
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee', 'errs', 'shar')>0
  IF USED('errs')
   USE IN errs
  ENDIF 
 ELSE 
  SELECT errs
  IF FIELD('vp')!='VP'
   USE IN errs
   IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\errsmee', 'errs', 'excl')>0
   ELSE 
    SELECT errs
    ALTER table errs ADD COLUMN vp n(1)
   ENDIF 
  ENDIF 
  IF USED('errs')
   USE IN errs
  ENDIF 
 ENDIF 
ENDIF 

IsFileExists = fso.FileExists(pCommon+'\pnorm.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\pnorm (period c(6), chnorm n(13,2), adnorm n(13,2), ;
 m0001 n(13,2), f0001 n(13,2), m0104 n(13,2), f0104 n(13,2), m0514 n(13,2), f0514 n(13,2),;
 m1517 n(13,2), f1517 n(13,2), m1824 n(13,2), f1824 n(13,2), m2534 n(13,2), f2534 n(13,2),;
 m3544 n(13,2), f3544 n(13,2), m4559 n(13,2), f4554 n(13,2), m6068 n(13,2), f5564 n(13,2),;
 m6999 n(13,2), f6599 n(13,2))
 INDEX ON period TAG period
 FOR jjj=1 TO 12
  m.period = '2014'+PADL(jjj,2,'0')
  INSERT INTO pnorm (period, chnorm, adnorm, m0001, f0001, m0104, f0104, m0514, f0514,;
   m1517, f1517, m1824, f1824, m2534, f2534, m3544, f3544, m4559, f4554, m6068, ;
   f5564, m6999, f6599) ;
  VALUES ;
  (m.period,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0) 
 ENDFOR 
 USE 
ELSE 
 IF OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')=0
  SELECT pnorm
  IF FIELD('m0001')!='M0001'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m0001 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f0001')!='F0001'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f0001 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m0104')!='M0104'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m0104 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f0104')!='F0104'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f0104 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m0514')!='M0514'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m0514 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f0514')!='F0514'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f0514 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m1517')!='M1517'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m1517 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f1517')!='F1517'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f1517 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m1824')!='M1824'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m1824 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f1824')!='F1824'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f1824 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m2534')!='M2534'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m2534 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f2534')!='F2534'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f2534 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m3544')!='M3544'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m3544 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f3544')!='F3544'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f3544 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m4559')!='M4559'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m4559 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f4554')!='F4554'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f4554 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m6068')!='M6068'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m6068 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f5564')!='F5564'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f5564 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('m6999')!='M6999'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN m6999 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF FIELD('f6599')!='F6599'
   USE IN pnorm
   IF OpenFile(pCommon+'\pnorm', 'pnorm', 'excl')=0
    ALTER table pnorm ADD COLUMN f6599 n(13,2)
   ENDIF
   IF USED('pnorm')
    USE IN pnorm
   ENDIF 
   =OpenFile(pCommon+'\pnorm', 'pnorm', 'shar')
  ENDIF 
  IF USED('pnorm')
   USE IN pnorm
  ENDIF 
 ELSE 
  IF USED('pnorm')
   USE IN pnorm
  ENDIF 
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase+'\'+gcPeriod+'\NSI')
IF IsDirExists == .F.
 IF MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ" + CHR(13) + "&pBase\&gcPeriod\NSI!" + CHR(13) + "—Œ«ƒ¿“‹?",4+32, "") == 6
  WAIT "—Œ«ƒ¿≈“—ﬂ ÀŒ ¿À‹Õ€… Õ—»..." WINDOW NOWAIT 

 fso.CreateFolder(pBase+'\'+gcPeriod+'\NSI')

tyu = pcommon+'\admokrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\admokrxx.dbf', pBase+'\'+gcPeriod+'\NSI\admokrxx.dbf')
  
tyu = pcommon+'\codku_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codku_xx.dbf', pBase+'\'+gcPeriod+'\NSI\codku.dbf')

tyu = pcommon+'\codotdxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codotdxx.dbf', pBase+'\'+gcPeriod+'\NSI\codotd.dbf')

tyu = pcommon+'\codwdrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\codwdrxx.dbf', pBase+'\'+gcPeriod+'\NSI\codwdr.dbf')

tyu = pcommon+'\hopff_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\hopff_xx.dbf', pBase+'\'+gcPeriod+'\NSI\hopff.dbf')

tyu = pcommon+'\isv012xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\isv012xx.dbf', pBase+'\'+gcPeriod+'\NSI\isv012.dbf')

tyu = pcommon+'\kdolgxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\kdolgxx.dbf', pBase+'\'+gcPeriod+'\NSI\kdolgxx.dbf')

tyu = pcommon+'\kpreslxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\kpreslxx.dbf', pBase+'\'+gcPeriod+'\NSI\kpresl.dbf')

*tyu = pcommon+'\kspecxx.dbf'
*oSettings.CodePage('&tyu', 866, .t.)
*fso.CopyFile(pcommon+'\kspecxx.dbf', pBase+'\'+gcPeriod+'\NSI\kspec.dbf')

tyu = pcommon+'\mkb10_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\mkb10_xx.dbf', pBase+'\'+gcPeriod+'\NSI\mkb10.dbf')

tyu = pcommon+'\mo_vmp_2014.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\mo_vmp_2014.dbf', pBase+'\'+gcPeriod+'\NSI\mo_vmp.dbf')

tyu = pcommon+'\modpacxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\modpacxx.dbf', pBase+'\'+gcPeriod+'\NSI\modpac.dbf')

tyu = pcommon+'\ms_mkbxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\ms_mkbxx.dbf', pBase+'\'+gcPeriod+'\NSI\ms_mkb.dbf')

tyu = pcommon+'\nocodrxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\nocodrxx.dbf', pBase+'\'+gcPeriod+'\NSI\nocodr.dbf')

tyu = pcommon+'\osoerzxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\osoerzxx.dbf', pBase+'\'+gcPeriod+'\NSI\osoerzxx.dbf')

tyu = pcommon+'\osoreexx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\osoreexx.dbf', pBase+'\'+gcPeriod+'\NSI\osoree.dbf')

tyu = pcommon+'\ososchxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\ososchxx.dbf', pBase+'\'+gcPeriod+'\NSI\ososch.dbf')

tyu = pcommon+'\profotxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\profotxx.dbf', pBase+'\'+gcPeriod+'\NSI\profot.dbf')

tyu = pcommon+'\profusxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\profusxx.dbf', pBase+'\'+gcPeriod+'\NSI\profus.dbf')

tyu = pcommon+'\prv002xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)

tyu = pcommon+'\rsv009xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\rsv009xx.dbf', pBase+'\'+gcPeriod+'\NSI\rsv009.dbf')

tyu = pcommon+'\sookodxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\sookodxx.dbf', pBase+'\'+gcPeriod+'\NSI\sookodxx.dbf')

tyu = pcommon+'\sovmnoxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\sovmnoxx.dbf', pBase+'\'+gcPeriod+'\NSI\sovmno.dbf')

tyu = pcommon+'\spi_lpu_dd_2014.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\spi_lpu_dd_2014.dbf', pBase+'\'+gcPeriod+'\NSI\spi_lpu_dd.dbf')

tyu = pcommon+'\spr_ulxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
IF OpenFile(pcommon+'\spr_ulxx', 'sprul', 'shar')<=0
 CREATE TABLE &pBase\&gcPeriod\NSI\street (ul n(5), street c(60), cokr c(2), mo c(4))
 INDEX ON ul TAG ul
 USE 
 =OpenFile(pBase+'\'+gcPeriod+'\NSI\street', 'street', 'excl')<=0
 
 SELECT sprul
 SCAN 
  m.priznak = priznak
  IF m.priznak != 'a'
   LOOP 
  ENDIF 
  m.ul = INT(VAL(kod_fo))
  m.street = ALLTRIM(nmstreet)
    
  INSERT INTO street FROM MEMVAR 
    
 ENDSCAN 
 USE 
 SELECT street
 SORT ON street TO &pBase\&gcPeriod\NSI\qwert
 ZAP
 APPEND FROM &pBase\&gcPeriod\NSI\qwert
 IF fso.FileExists(pBase+'\'+gcPeriod+'\NSI\qwert.dbf')
  fso.DeleteFile(pBase+'\'+gcPeriod+'\NSI\qwert.dbf')
 ENDIF 
 USE IN street 
ENDIF 

tyu = pcommon+'\sprlpuxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
  
IF OpenFile(pcommon+'\sprlpuxx', 'sprlpu', 'shar')<=0
 SELECT sprlpu
 COPY FOR lpu_id=fil_id TO pBase+'\'+gcPeriod+'\NSI\sprlpuxx' ;
  FIELDS lpu_id,fil_id,tpn,vmp,mcod,name,fullname,cokr,adres
 COPY FOR name='œÓÎËÍÎËÌËÍ‡ Ò “œÕ' TO pBase+'\'+gcPeriod+'\NSI\lputpn' ;
  FIELDS lpu_id,fil_id,tpn,vmp,mcod,name,fullname,cokr,adres
 COPY FOR tpn='4' TO pBase+'\'+gcPeriod+'\NSI\horlpu' ;
  FIELDS lpu_id,fil_id,tpn,vmp,mcod,name,fullname,cokr,adres
 USE 
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'excl')<=0
  SET FULLPATH OFF 
  WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
  INDEX ON lpu_id TAG lpu_id
  INDEX ON fil_id TAG fil_id
  INDEX ON mcod TAG mcod
  INDEX ON cokr TAG cokr
  USE 
  SET FULLPATH OFF 
 ENDIF 
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\lputpn', 'lputpn', 'excl')<=0
  SET FULLPATH OFF 
  WAIT "»Õƒ≈ —»–Œ¬¿Õ»≈ ‘¿…À¿ "+ALLTRIM(DBF())+' ...' WINDOW NOWAIT 
  DELETE FOR INLIST(lpu_id,1891,4475)
  INDEX ON lpu_id TAG lpu_id
  INDEX ON fil_id TAG fil_id
  INDEX ON mcod TAG mcod
  USE 
  SET FULLPATH OFF 
 ENDIF 
ENDIF 

tyu = pcommon+'\spraboxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)

IF OpenFile(pcommon+'\spraboxx', 'sprabo', 'shar')<=0
 IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')<=0
  SELECT sprabo
  SET RELATION TO object_id INTO sprlpu
  COPY FOR !EMPTY(sprlpu.lpu_id) AND abn_type='0' fields object_id, abn_name, name;
  TO pBase+'\'+gcPeriod+'\NSI\spraboxx'
  SET RELATION OFF INTO sprlpu
  USE 
  USE IN sprlpu
 ELSE 
  USE IN sprlpu
 ENDIF 
ENDIF 

tyu = pcommon+'\spv015xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\spv015xx.dbf', pBase+'\'+gcPeriod+'\NSI\spv015.dbf')

tyu = pcommon+'\tipgrpxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\tipgrpxx.dbf', pBase+'\'+gcPeriod+'\NSI\tipgrp.dbf')

tyu = pcommon+'\tipno_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\tipno_xx.dbf', pBase+'\'+gcPeriod+'\NSI\tipnomes.dbf')

tyu = pcommon+'\vidvp_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\vidvp_xx.dbf', pBase+'\'+gcPeriod+'\NSI\vidvp.dbf')

tyu = pcommon+'\z_cod_xx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\z_cod_xx.dbf', pBase+'\'+gcPeriod+'\NSI\z_cod.dbf')

tyu = pcommon+'\z_dsnoxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
fso.CopyFile(pcommon+'\z_dsnoxx.dbf', pBase+'\'+gcPeriod+'\NSI\z_dsnoxx.dbf')

fso.CopyFile(pcommon+'\polic_dp.dbf', pBase+'\'+gcPeriod+'\NSI\polic_dp.dbf')
fso.CopyFile(pcommon+'\polic_h.dbf', pBase+'\'+gcPeriod+'\NSI\polic_h.dbf')
*fso.CopyFile(pcommon+'\spi_cz.dbf', pBase+'\'+gcPeriod+'\NSI\spi_cz.dbf')
*fso.CopyFile(pcommon+'\spi_cz_ch.dbf', pBase+'\'+gcPeriod+'\NSI\spi_cz_ch.dbf')
fso.CopyFile(pcommon+'\usrlpu.dbf', pBase+'\'+gcPeriod+'\NSI\usrlpu.dbf')

tyu = pcommon+'\errsmee.dbf'
fso.CopyFile(pcommon+'\errsmee.dbf', pBase+'\'+gcPeriod+'\NSI\errsmee.dbf')
tyu = pcommon+'\errsmee.cdx'
fso.CopyFile(pcommon+'\errsmee.cdx', pBase+'\'+gcPeriod+'\NSI\errsmee.cdx')

IF fso.FileExists(pcommon+'\tarifn.dbf')
 IF OpenFile(pcommon+'\tarifn', 'tarif', 'excl')<=0
  SELECT tarif
  INDEX ON cod TAG cod
  SET ORDER TO cod 

  IF fso.FileExists(pcommon+'\tarimuxx.dbf')
   tyu = pcommon+'\tarimuxx.dbf'
   oSettings.CodePage('&tyu', 866, .t.)

   IF OpenFile(pcommon+'\tarimuxx', 'tarimu', 'excl')<=0
   
*    MESSAGEBOX(CHR(13)+CHR(10)+' ŒÀ»◊≈—“¬Œ «¿œ»—≈… ¬ ‘¿…À≈ TARIFN: '+TRANSFORM(RECCOUNT('tarif'),'99999')+CHR(13)+CHR(10)+;
     'Õ≈ —ŒŒ“¬≈“—¬”≈“  ŒÀ»◊≈—“¬” «¿œ»—≈… ¬ ‘¿…À≈ TARIMU: '+TRANSFORM(RECCOUNT('tarimu'),'99999')+CHR(13)+CHR(10),0+48,'')
    SELECT tarimu 
    INDEX ON cod TAG cod
    SET ORDER TO cod 
    
    m.nIsChngTarif  = 0
    m.nIsChngTarifV = 0
    m.nIsChngStkd   = 0
    m.nIsChngStkdV  = 0
    
    SELECT tarif
    SET RELATION TO cod INTO tarimu
    SCAN 
     IF tarif!=tarimu.tarif
      REPLACE tarif WITH tarimu.tarif
      m.nIsChngTarif  = m.nIsChngTarif + 1
     ENDIF 
*     IF tarif_v!=tarimu.tarif_v
*      REPLACE tarif_v WITH tarimu.tarif_v
*      m.nIsChngTarifV  = m.nIsChngTarifV + 1
*     ENDIF 
     IF tarif_v!=tarimu.tarif
      REPLACE tarif_v WITH tarimu.tarif
      m.nIsChngTarifV  = m.nIsChngTarifV + 1
     ENDIF 
     IF stkd!=tarimu.stkd
      REPLACE stkd WITH tarimu.stkd
      m.nIsChngstkd  = m.nIsChngstkd + 1
     ENDIF 
*     IF stkdv!=tarimu.stkdv
*      REPLACE stkdv WITH tarimu.stkdv
*      m.nIsChngstkdv  = m.nIsChngstkdv + 1
*     ENDIF 
     IF stkdv!=tarimu.stkd
      REPLACE stkdv WITH tarimu.stkd
      m.nIsChngstkdv  = m.nIsChngstkdv + 1
     ENDIF 
    ENDSCAN 
    SET RELATION OFF INTO tarimu
    SET ORDER TO 
    DELETE TAG ALL 
    USE 
    
    SELECT tarimu
    SET ORDER TO 
    DELETE TAG ALL 
    USE 

    IF m.nIsChngTarif > 0
     MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngTarif,'99999')+' «Õ¿◊≈Õ»… TARIF '+CHR(13)+CHR(10)+;
      '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 
    IF m.nIsChngTarifV > 0
     MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngTarifV,'99999')+' «Õ¿◊≈Õ»… TARIF_V '+CHR(13)+CHR(10)+;
      '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 
    IF m.nIsChngStkd > 0
     MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngStkd,'99999')+' «Õ¿◊≈Õ»… STKD '+CHR(13)+CHR(10)+;
     '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 
    IF m.nIsChngStkdV > 0
     MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngStkdV,'99999')+' «Õ¿◊≈Õ»… STKDV '+CHR(13)+CHR(10)+;
      '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
    ENDIF 

   ELSE 
    IF USED('tarimu')
     USE IN tarimu
    ENDIF 
   ENDIF 
  ENDIF 
 ELSE 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
 ENDIF 
ENDIF 

WAIT "Œ¡ÕŒ¬À≈Õ»≈..." WINDOW NOWAIT 
tyu = pcommon+'\tarimuxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
tyu = pcommon+'\reesusxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
tyu = pcommon+'\reesmsxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)
tyu = pcommon+'\usvmpxx.dbf'
oSettings.CodePage('&tyu', 866, .t.)

IF fso.FileExists(pcommon+'\tarifn.dbf')
 IF OpenFile(pcommon+'\tarifn', 'tarif', 'excl')<=0
  SELECT tarif
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\tarimuxx.dbf')
 IF OpenFile(pcommon+'\tarimuxx', 'tarimu', 'excl')<=0
  SELECT tarimu 
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\reesusxx.dbf')
 IF OpenFile(pcommon+'\reesusxx', 'reesus', 'excl')<=0
  SELECT reesus
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\reesmsxx.dbf')
 IF OpenFile(pcommon+'\reesmsxx', 'reesms', 'excl')<=0
  SELECT reesms
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 
IF fso.FileExists(pcommon+'\usvmpxx.dbf')
 IF OpenFile(pcommon+'\usvmpxx', 'usvmp', 'excl')<=0
  SELECT usvmp
  INDEX on cod TAG cod
  SET ORDER TO cod 
 ENDIF 
ENDIF 

*IF USED('tarif')
* MESSAGEBOX('OK',0+64,'tarif')
*ENDIF 
*IF USED('tarimu')
* MESSAGEBOX('OK',0+64,'tarimu')
*ENDIF 
*IF USED('reesus')
* MESSAGEBOX('OK',0+64,'reesus')
*ENDIF 
*IF USED('reesms')
* MESSAGEBOX('OK',0+64,'reesms')
*ENDIF 
*IF USED('usvmp')
* MESSAGEBOX('OK',0+64,'usvmp')
*ENDIF 

IF USED('tarif') AND USED('tarimu') AND USED('reesus') AND USED('reesms') AND USED('usvmp')

SELECT reesus
SET RELATION TO cod INTO tarif
COUNT FOR EMPTY(tarif.cod) TO m.nNewInUs
IF m.nNewInUs > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'¬ –≈≈—“–≈ ”—À”√ (REESUSXX) Œ¡Õ¿–”∆≈ÕŒ '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nNewInUs,'99999')+' ÕŒ¬€’ «¿œ»—≈…!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  IF EMPTY(tarif.cod)
   SCATTER MEMVAR 
   m.comment = m.name 
   INSERT INTO tarif FROM MEMVAR 
  ENDIF 
 ENDSCAN 
ENDIF 
SET RELATION OFF INTO tarif

SELECT tarif
SET RELATION TO cod INTO reesus
COUNT FOR EMPTY(reesus.cod) AND isitus(cod) TO m.nDelInUs

IF m.nDelInUs>0
 MESSAGEBOX(CHR(13)+CHR(10)+'»« –≈≈—“–¿ ”—À”√ (REESUSXX) »— Àﬁ◊≈ÕŒ '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nDelInUs,'99999')+' «¿œ»—≈…!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  m.cod = cod
  IF IsItUs(m.cod) AND EMPTY(reesus.cod)
   DELETE 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO reesus
 PACK 
 SET RELATION TO cod INTO reesus
ENDIF 

m.nIsChngTpn  = 0
m.nIsChngUet1 = 0
m.nIsChngUet2 = 0

SCAN 
 m.cod = cod
 IF !IsItUs(m.cod)
  LOOP 
 ENDIF 
 
 IF tpn != reesus.tpn
  m.nIsChngTpn  = m.nIsChngTpn + 1
  REPLACE tpn WITH reesus.tpn
 ENDIF 
 IF uet1 != reesus.uet1
  m.nIsChngUet1  = m.nIsChngUet1 + 1
  REPLACE uet1 WITH reesus.uet1
 ENDIF 
 IF uet2 != reesus.uet2
  m.nIsChngUet2  = m.nIsChngUet2 + 1
  REPLACE uet2 WITH reesus.uet2
 ENDIF 
ENDSCAN 
SET RELATION OFF INTO reesus

IF m.nIsChngTpn > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngTpn,'99999')+' «Õ¿◊≈Õ»… TPN '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngUet1 > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngUet1,'99999')+' «Õ¿◊≈Õ»… UET1 '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngUet2 > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngUet2,'99999')+' «Õ¿◊≈Õ»… UET2 '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 

SELECT reesms
SET RELATION TO cod INTO tarif
COUNT FOR EMPTY(tarif.cod) TO m.nNewInMs
IF m.nNewInMs > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'¬ –≈≈—“–≈ Ã›—Œ¬ (REESMSXX) Œ¡Õ¿–”∆≈ÕŒ '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nNewInMs,'99999')+' ÕŒ¬€’ «¿œ»—≈…!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  IF EMPTY(tarif.cod)
   SCATTER MEMVAR 
   m.name    = m.namem
   m.comment = m.namem
   INSERT INTO tarif FROM MEMVAR 
  ENDIF 
 ENDSCAN 
ENDIF 
SET RELATION OFF INTO tarif

SELECT tarif
SET RELATION TO cod INTO reesms
COUNT FOR EMPTY(reesms.cod) AND !isitus(cod) TO m.nDelInMs

IF m.nDelInMs > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'»« –≈≈—“–¿ Ã›—Ó‚ (REESMSXX) »— Àﬁ◊≈ÕŒ '+CHR(13)+CHR(10)+;
  TRANSFORM(m.nDelInMs,'99999')+' «¿œ»—≈…!'+CHR(13)+CHR(10),0+64,'')
 SCAN 
  m.cod = cod
  IF !IsItUs(m.cod) AND EMPTY(reesms.cod)
   DELETE 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO reesms
 PACK 
 SET RELATION TO cod INTO reesms
ENDIF 

m.nIsChngNKD = 0
m.nIsChngSTKD = 0

SCAN 
 m.cod  = cod

 IF IsItUs(m.cod)
  LOOP 
 ENDIF 
 
 IF n_kd != reesms.n_kd
  m.nIsChngNKD  = m.nIsChngNKD + 1
  REPLACE n_kd WITH reesms.n_kd
 ENDIF 

ENDSCAN 
SET RELATION OFF INTO reesms

IF m.nIsChngNKD > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngNKD,'99999')+' «Õ¿◊≈Õ»… N_KD '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 

m.nIsChngTarif  = 0
m.nIsChngTarifV = 0
m.nIsChngStkd   = 0
m.nIsChngStkdV  = 0

SET RELATION TO cod INTO tarimu
SCAN 
 m.tarif   = tarif
 m.tarif_v = tarif_v
 m.stkd    = stkd
 m.stkdv   = stkdv
 
 IF tarif!=tarimu.tarif
  REPLACE tarif WITH tarimu.tarif
  m.nIsChngTarif  = m.nIsChngTarif + 1
 ENDIF 
* IF tarif_v!=tarimu.tarif_v
 IF tarif_v!=tarimu.tarif
*  REPLACE tarif_v WITH tarimu.tarif_v
  REPLACE tarif_v WITH tarimu.tarif
  m.nIsChngTarifV  = m.nIsChngTarifV + 1
 ENDIF 
 IF stkd!=tarimu.stkd
  REPLACE stkd WITH tarimu.stkd
  m.nIsChngstkd  = m.nIsChngstkd + 1
 ENDIF 
* IF stkdv!=tarimu.stkdv
 IF stkdv!=tarimu.stkd
*  REPLACE stkdv WITH tarimu.stkdv
  REPLACE stkdv WITH tarimu.stkd
  m.nIsChngstkdv  = m.nIsChngstkdv + 1
 ENDIF 

ENDSCAN 
SET RELATION OFF INTO tarimu

IF m.nIsChngTarif > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngTarif,'99999')+' «Õ¿◊≈Õ»… TARIF '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngTarifV > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngTarifV,'99999')+' «Õ¿◊≈Õ»… TARIF_V '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngStkd > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngStkd,'99999')+' «Õ¿◊≈Õ»… STKD '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 
IF m.nIsChngStkdV > 0
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡ÕŒ¬À≈ÕŒ '+TRANSFORM(m.nIsChngStkdV,'99999')+' «Õ¿◊≈Õ»… STKDV '+CHR(13)+CHR(10)+;
  '¬ ‘¿…À≈ TARIFN!'+CHR(13)+CHR(10),0+64,'')
ENDIF 

ELSE 
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—¬”≈“ Œƒ»Õ »« ◊≈“€–≈’ Õ≈Œ¡’Œƒ»Ã€’'+;
  CHR(13)+CHR(10)+'ƒÀﬂ —Œ«ƒ¿Õ»ﬂ TARIFN ‘¿…ÀŒ¬. TARIFN Õ≈ œ≈–≈—Œ«ƒ¿Õ!',0+64,'')
ENDIF 

IF USED('tarif')
 SELECT tarif 
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN tarif
ENDIF 
IF USED('tarimu')
 SELECT tarimu
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN tarimu
ENDIF 
IF USED('reesus')
 SELECT reesus
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN reesus
ENDIF 
IF USED('reesms')
 SELECT reesms
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN reesms
ENDIF 
IF USED('usvmp')
 SELECT usvmp
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN usvmp
ENDIF 

fso.CopyFile(pcommon+'\tarifn.dbf', pBase+'\'+gcPeriod+'\NSI\tarifn.dbf')
WAIT CLEAR 

DO comreind

*  fso.CopyFolder(pCommon, pBase+'\'+gcPeriod+'\NSI')
  WAIT CLEAR 
 ENDIF 
ELSE 
 IF fso.FileExists(pBase+'\'+gcPeriod+'\NSI\sprlpuxx.dbf')
  IF !fso.FileExists(pBase+'\'+gcPeriod+'\NSI\horlpu.dbf')
   IF OpenFile(pBase+'\'+gcPeriod+'\NSI\sprlpuxx', 'sprlpu', 'shar')>0
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
   ELSE 
    SELECT sprlpu 
    COPY FOR tpn='4' TO pBase+'\'+gcPeriod+'\NSI\horlpu' ;
     FIELDS lpu_id,fil_id,tpn,vmp,mcod,name,fullname,cokr,adres
    IF OpenFile(pBase+'\'+gcPeriod+'\NSI\horlpu', 'horlpu', 'shar')>0
     IF USED('horlpu')
      USE IN horlpu
     ENDIF 
    ELSE 
     SELECT horlpu
     INDEX ON lpu_id TAG lpu_id
     INDEX ON fil_id TAG fil_id 
     INDEX ON mcod TAG mcod
     USE IN horlpu
    ENDIF 
    USE IN sprlpu
   ENDIF 
  ENDIF 
 ENDIF 
ENDIF 

FUNCTION IsItUs(m.usl)
 PRIVATE m.usl
 m.IsItUs = .F.
 IF BETWEEN(m.usl,1001,60999) OR BETWEEN(m.usl,96001,99999) OR ;
  BETWEEN(m.usl,101001,160999) OR BETWEEN(m.usl,196001,199999)
  m.IsItUs = .T.
 ENDIF 
RETURN m.IsItUs