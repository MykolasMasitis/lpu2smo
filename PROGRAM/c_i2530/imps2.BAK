PROCEDURE ImpS2

IF MESSAGEBOX('бш унрхре хлонпрхпнбюрэ дюммше?'+CHR(13)+CHR(10),4+32,'S2')=7
 RETURN 
ENDIF 

pUpDir = fso.GetParentFolderName(pBin)
pUpdDir = pUpDir+'\S2'
IF !fso.FolderExists(pUpdDir)
 fso.CreateFolder(pUpdDir)
ENDIF 

oMailDir = fso.GetFolder(pUpdDir)
MailDirName = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('намюпсфемн '+ALLTRIM(STR(nFilesInMailDir))+' тюикнб!', 0+64,'')

IF nFilesInMailDir<=0
 RETURN 
ENDIF 

CREATE CURSOR curlpu (lpuid n(4), mm c(2))
INDEX on lpuid TAG lpuid
INDEX on STR(lpuid)+PADL(mm,2,'0') TAG unik 
SET ORDER TO unik 

CREATE CURSOR curmm (mm c(2))
INDEX on mm TAG mm
SET ORDER TO mm

FOR EACH oFileInMailDir IN oFilesInMailDir

 m.BFullName = oFileInMailDir.Path
 m.bname     = oFileInMailDir.Name
 
 IF LEN(m.bname)!=12
  LOOP 
 ENDIF 
 IF LOWER(LEFT(m.bname,2))!='pg'
  LOOP 
 ENDIF 
 m.lpuid = INT(VAL(SUBSTR(m.bname,3,4)))
 m.mm    = SUBSTR(m.bname,7,2)
 m.unik  = STR(m.lpuid,4)+m.mm
 
 IF !SEEK(m.unik, 'curlpu')
  INSERT INTO curlpu FROM MEMVAR 
 ENDIF 
 
 IF !SEEK(m.mm, 'curmm')
  INSERT INTO curmm FROM MEMVAR 
 ENDIF 

ENDFOR 

IF RECCOUNT('curmm')<=0
 USE IN curmm
 USE IN curllpu
 MESSAGEBOX('ме сдюкняэ нопедекхрэ леяъж!',0+16,'')
 RETURN 
ENDIF 

IF RECCOUNT('curmm')>1
 MESSAGEBOX('бшаепхре леяъж!',0+64,'')
 SELECT curmm
 m.mm = '12'
 BROWSE 
 RETURN 
ENDIF 

IF RECCOUNT('curmm')=1
 m.mm = curmm.mm
 m.mmname = mes_main(INT(VAL(m.mm)))
 MESSAGEBOX(m.mmname,0+64,'')
ENDIF 
USE IN curmm

m.lYear    = STR(m.tyear,4)
m.lMon     = m.mm
m.lcPeriod = m.lYear+m.mm

IF !fso.FolderExists(pBase+'\'+m.lcPeriod)
 USE IN curlpu
 MESSAGEBOX('нрясрярбсер дхпейрпхъ '+m.lcPeriod+'!',0+16,'')
ENDIF 

IF !fso.FileExists(pBase+'\'+m.lcPeriod+'\aisoms.dbf')
 USE IN curlpu
 MESSAGEBOX('нрясрярбсер тюик '+m.lcPeriod+'\aisoms.dbf!',0+16,'')
 RETURN 
ENDIF 
IF OpenFile(pBase+'\'+m.lcPeriod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
 IF USED('aisoms')
  USE IN aisoms
  USE IN curlpu
 ENDIF 
 RETURN 
ENDIF 

StartOfProc = SECONDS()
nLpu = 0

SELECT curlpu
SCAN 
 IF mm!=m.mm
  LOOP 
 ENDIF 
 m.lpuid = lpuid
 m.mcod = IIF(SEEK(m.lpuid, 'aisoms'), aisoms.mcod, '')
 IF EMPTY(m.mcod)
  LOOP 
 ENDIF 
 IF !fso.FolderExists(pBase+'\'+m.lcPeriod+'\'+m.mcod)
*  MESSAGEBOX('нрясрябсе дхпейрнпхъ '+m.mcod+'!',0+64,m.mcod)
  LOOP 
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\people.dbf')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\people.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\people.cdx')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\people.cdx')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.dbf')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.cdx')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.cdx')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\otdel.dbf')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\otdel.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\otdel.cdx')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\otdel.cdx')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\doctor.dbf')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\doctor.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\doctor.cdx')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\doctor.cdx')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\answer.dbf')
  fso.DeleteFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\answer.dbf')
 ENDIF 
 
 m.pfile  = 'pg'+STR(m.lpuid,4)+m.mm && people
 m.tfile  = 'rg'+STR(m.lpuid,4)+m.mm && talon

 IF !fso.FileExists(pUpdDir+'\'+m.pfile+'.dbf')
  LOOP 
 ENDIF 
 IF !fso.FileExists(pUpdDir+'\'+m.tfile+'.dbf')
  LOOP 
 ENDIF 

 m.People = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\people'
 m.Talon  = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon'
 m.answer = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\Answer'
 m.otdel  = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\Otdel'
 m.doctor = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\Doctor'
 m.error  = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\e'+m.mcod
 m.merror = pBase+'\'+m.lcPeriod+'\'+m.mcod+'\m'+m.mcod

 IF fso.FileExists(m.error+'.dbf')
  fso.DeleteFile(m.error+'.dbf')
 ENDIF 
 IF fso.FileExists(m.error+'.cdx')
  fso.DeleteFile(m.error+'.cdx')
 ENDIF 
 IF fso.FileExists(m.merror+'.dbf')
  fso.DeleteFile(m.merror+'.dbf')
 ENDIF 
 IF fso.FileExists(m.merror+'.cdx')
  fso.DeleteFile(m.merror+'.cdx')
 ENDIF 

 CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1,;
   mcod c(7), prmcod c(7), period c(6), d_beg d, d_end d, s_all n(11,2), ;
   tip_p n(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, ;
   ul n(5), dom c(7), kor c(5), str c(5), kv c(5), d_type c(1), ;
   sv c(3), recid_lpu c(7), IsPr L)
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX on s_all TAG s_all
 USE 
 
 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(25), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), IsPr L, vz l)

 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX on pcod TAG pcod
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
 INDEX ON sn_pol+STR(cod,6)+DTOS(d_u) TAG ExpTag
 INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) TAG unik 
 INDEX ON tip TAG tip
 INDEX ON s_all TAG s_all
 USE 

 CREATE TABLE (Answer) ;
  (RecId c(6),s_pol c(6),n_pol c(16),q c(2),fam c(25),im c(20),ot c(20),dr c(8),w n(1),ans_r c(3),tip_d c(1),lpu_id n(6))
 USE 

 CREATE TABLE (Otdel) ;
	(recid c(6), mcod c(7), iotd c(8), name c(100), pr_name c(100), cnt_bed n(5), fil_id n(6))
 INDEX ON iotd TAG iotd
 USE 

 CREATE TABLE (Doctor) ;
   (pcod c(10),sn_pol c(25),fam c(25),im c(20),ot c(20),dr d, w n(1),;
    prvs n(4), d_ser d, d_prik d, iotd c(8),;
	lgot_r c(1),c_ogrn c(15),lpu_id n(6), fil_id n(6))
 INDEX ON pcod TAG pcod
 USE 

 CREATE TABLE (Error) (f c(1), c_err c(3), rid i)
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 

 CREATE TABLE (mError) ;
  (rid i autoinc, RecId i, cod n(6), k_u n(3), tip c(1), et c(1), ee c(1), usr c(6), d_exp d,;
   e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3), osn230 c(5), e_period c(6),  ;
   koeff n(4,2), straf n(4,2), docexp c(7), ;
   s_all n(11,2), s_1 n(11,2), s_2 n(11,2), impdata d)
 INDEX ON rid TAG rid 
 INDEX ON RecId TAG recid
 INDEX ON PADL(recid,6,'0')+et TAG id_et
 INDEX ON PADL(recid,6,'0')+et+LEFT(err_mee,2) TAG unik
 USE 

 IF OpenFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people
  ENDIF 
  SELECT curlpu
  LOOP 
 ENDIF 
 IF OpenFile(pUpdDir+'\'+m.pfile, 'pfile', 'shar')>0
  IF USED('pfile')
   USE IN pfile
  ENDIF 
  USE IN people 
  SELECT curlpu
  LOOP 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  USE IN people
  USE IN pfile
  SELECT curlpu
  LOOP 
 ENDIF 
 IF OpenFile(pUpdDir+'\'+m.tfile, 'tfile', 'shar')>0
  IF USED('tfile')
   USE IN tfile
  ENDIF 
  USE IN talon 
  USE IN pfile 
  USE IN people 
  SELECT curlpu
  LOOP 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\Answer', 'answer', 'shar')>0
  IF USED('answer')
   USE IN answer
  ENDIF 
  USE IN talon 
  USE IN people
  USE IN pfile
  SELECT curlpu
  LOOP 
 ENDIF 
 IF OpenFile(m.error, 'error', 'shar')>0
  IF USED('error')
   USE IN error
  ENDIF 
  USE IN Answer
  USE IN talon 
  USE IN people
  USE IN pfile
  SELECT curlpu
  LOOP 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\profus', 'profus', 'shar', 'cod')>0
  IF USED('profus')
   USE IN profus
  ENDIF 
  USE IN error
  USE IN Answer
  USE IN talon 
  USE IN people
  USE IN pfile
  SELECT curlpu
  LOOP 
 ENDIF 
 
 WAIT m.mcod WINDOW NOWAIT 
 
 SELECT pfile
 SCAN 
  SCATTER MEMVAR 
  RELEASE recid, tip_p
  m.prlpu = lpu_id_erz
  m.prmcod = IIF(SEEK(m.prlpu, 'aisoms'), aisoms.mcod, '')
  m.period = m.lYear+m.mm
  m.s_all = pr_all
  m.tipp= tip_p
  m.recid_lpu = recid 
  m.sv = ans_r_erz
  
  INSERT INTO people FROM MEMVAR
  m.recid = GETAUTOINCVALUE()
  
  m.c_err = ALLTRIM(koder)
  m.rid = m.recid
  IF !EMPTY(m.c_err)
   m.c_err = ALLTRIM(koder)+'A'
   m.f = 'R'
   INSERT INTO error FROM MEMVAR 
  ENDIF 

  m.recid  = PADL(m.recid,6,'0')
  m.s_pol  = s_pol_erz
  m.n_pol  = n_pol_erz
  m.q      = q_erz
  m.ans_r  = ans_r_erz
  m.tip_d  = tip_d
  m.lpu_id = lpu_id_erz
  m.dr = DTOS(m.dr)
  INSERT INTO answer FROM MEMVAR 
  
 ENDSCAN 
 USE IN pfile
 USE IN people
 USE IN answer
 
 SELECT tfile 
 SCAN 
  SCATTER MEMVAR 
  RELEASE recid 
  m.recid_lpu = recid 
  m.otd       = iotd
  m.profil = IIF(SEEK(m.cod, 'profus'), ALLTRIM(profus.profil), '')
  
  INSERT INTO talon FROM MEMVAR 
  m.recid = GETAUTOINCVALUE()
  
  m.c_err = ALLTRIM(koder)
  m.rid = m.recid
  IF !EMPTY(m.c_err)
   m.c_err = ALLTRIM(koder)+'A'
   m.f = 'S'
   INSERT INTO error FROM MEMVAR 
  ENDIF 

 ENDSCAN 
 USE IN tfile
 USE IN talon 
 USE IN error 
 USE IN profus

 nLpu = nLpu + 1

 WAIT CLEAR 
 SELECT curlpu

ENDSCAN 
USE IN curlpu

EndOfProc  = SECONDS()
LastOfProc = EndOfProc - StartOfProc
MeanTime   = LastOfProc/nLpu

USE IN aisoms

MESSAGEBOX(CHR(13)+CHR(10)+"напюанрйю гюйнмвемю!"+CHR(13)+CHR(10)+;
  "бяецн напюанрюмн кос   : "+TRANSFORM(nLpu, '9999999')+CHR(13)+CHR(10)+;
  "наыее бпелъ напюанрйх  : "+TRANSFORM(LastOfProc,'999.999')+" ЯЕЙ."+CHR(13)+CHR(10)+;
  "япедмее бпелъ напюанрйх: "+TRANSFORM(MeanTime,'999.999')+" ЯЕЙ."+CHR(13)+CHR(10),0+64,"")

RETURN 