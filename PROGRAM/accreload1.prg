FUNCTION  AccReload1(lcmcod, lnlpuid, ismess)
 m.mcod  = lcmcod
 m.lpuid = lnlpuid
 m.IsVed = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

 IF ismess=.t.
  IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ œ≈–≈«¿√–”«»“‹ —◊≈“?'+CHR(13)+CHR(10),4+32,m.mcod)=7
   RETURN 
  ENDIF 
 ENDIF 
 
 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.bpath = pbase+'\'+gcperiod+'\'+m.mcod
 m.bfile = 'b'+m.mcod+'.'+m.mmy
 
 IF !fso.FileExists(m.bpath+'\'+m.bfile)
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À '+m.bfile+' Õ≈ Õ¿…ƒ≈Õ!'+CHR(13)+CHR(10),0+64,'')
  ENDIF 
  RETURN 
 ENDIF 

 ffile = fso.GetFile(m.bpath+'\'+m.bfile)
 lcHead = ''
 IF ffile.size >= 2
  fhandl = ffile.OpenAsTextStream
  lcHead = fhandl.Read(2)
  fhandl.Close
 ENDIF 
 
 IF lchead!='PK'
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À '+UPPER(m.bfile)+' Õ≈ ZIP-¿–’»¬!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN 
 ENDIF 
 
 m.lIsOpen=UnzipOpen(m.bpath+'\'+m.bfile)
 IF !m.lIsOpen
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À '+UPPER(m.bfile)+' Õ≈ ZIP-¿–’»¬!'+CHR(13)+CHR(10),0+16,lcHead)
  ENDIF 
  RETURN 
 ENDIF 
 m.lIsClosed=UnzipClose()
 IF !m.lIsClosed
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À '+UPPER(m.bfile)+' Õ≈ ZIP-¿–’»¬!'+CHR(13)+CHR(10),0+16,lcHead)
  ENDIF 
  RETURN 
 ENDIF 
 
 dItem    = 'D'  + STR(m.lpuid,4) + '.' + m.mmy
 hItem    = 'H'  + STR(m.lpuid,4) + '.' + m.mmy
 nvItem   = 'NV' + STR(m.lpuid,4) + '.' + m.mmy
 nsItem   = 'NS' + STR(m.lpuid,4) + '.' + m.mmy
 rItem    = 'R' + m.qcod + '.' + m.mmy
 sItem    = 'S' + m.qcod + '.' + m.mmy
 dsItem   = 'D79S' + m.qcod + '.' + m.mmy
 sprItem  = 'SPR' + STR(m.lpuid,4) + '.' + m.mmy

 IF !IsIpComplete()
  RETURN 
 ENDIF 
 
 ZipDir  = m.bpath+'\'+SYS(3)
 =UnZipItems()
 
 IF OpenItems(ZipDir)>0
  =CloseItems()
  fso.DeleteFolder(ZipDir)
  RETURN 
 ENDIF 

 IF OpenTemplates()>0
  =CloseTemplates()
  =CloseItems()
  fso.DeleteFolder(ZipDir)
  RETURN 
 ENDIF 
 
 IF !CheckFilesStucture()
  =CloseTemplates()
  =CloseItems()
  fso.DeleteFolder(ZipDir)
  RETURN 
 ENDIF 

 =CloseTemplates()
 =CloseItems()

 oMailDir = fso.GetFolder(pbase+'\'+gcperiod+'\'+m.mcod)
 oFilesInMailDir = oMailDir.Files
 FOR EACH oFileInMailDir IN oFilesInMailDir
  m.BFullName = oFileInMailDir.Path
  m.bname     = oFileInMailDir.Name
  IF LOWER(m.bname)==m.bfile
   LOOP 
  ENDIF 
  fso.DeleteFile(m.BFullName)
 ENDFOR 
 
 fso.CopyFile(ZipDir+'\'+rItem, m.bpath+'\'+rItem)
 fso.CopyFile(ZipDir+'\'+sItem, m.bpath+'\'+sItem)
 fso.CopyFile(ZipDir+'\'+dsItem, m.bpath+'\'+dsItem)
 fso.CopyFile(ZipDir+'\'+dItem, m.bpath+'\'+dItem)
 fso.CopyFile(ZipDir+'\'+hItem, m.bpath+'\'+hItem)
 fso.CopyFile(ZipDir+'\'+nvItem, m.bpath+'\'+nvItem)
 fso.CopyFile(ZipDir+'\'+nsItem, m.bpath+'\'+nsItem)
 fso.CopyFile(ZipDir+'\'+sprItem, m.bpath+'\'+sprItem)
 fso.DeleteFolder(ZipDir)

 People = m.bpath + '\people'
 Talon  = m.bpath + '\talon'
 Otdel  = m.bpath + '\otdel'
 Doctor = m.bpath + '\doctor'
 Error  = m.bpath + '\e' + m.mcod
 mError = m.bpath + '\m' + m.mcod

 =CreateFilesStructure()

 m.s_pred  = 0
 m.paz     = 0
 m.nsch    = 0
 m.krank   = 0
 m.paz_dst = 0
 m.paz_st  = 0
 m.sum_kr  = 0
 m.sum_dst = 0
 m.sum_st  = 0
 m.usl_amb = 0
 m.kd_dst = 0
 m.ms_st  = 0

 SELECT aisoms
 REPLACE s_pred WITH m.s_pred, sum_flk WITH 0, paz WITH m.paz, nsch WITH m.nsch,;
  mee WITH 0, meebad WITH 0, ;
  ambchkdmee WITH 0, dstchkdmee WITH 0, stchkdmee WITH 0, ;
  krank WITH m.krank, paz_dst WITH m.paz_dst, paz_st WITH m.paz_st,;
  sum_kr WITH m.sum_kr, sum_dst WITH m.sum_dst, sum_st WITH m.sum_st,;
  usl_amb WITH m.usl_amb, kd_dst WITH m.kd_dst, ms_st WITH m.ms_st, erz_status WITH 0, erz_id WITH '' ;
  FOR mcod = m.mcod

 IF OpenLocalFiles()>0
  =CloseLocalFiles()
  =CloseTemplates()
  =CloseItems()
  =ClDir()
  RETURN 
 ENDIF 

 =MakePeople()
 =MakeTalon()
 =MakeOtdel() 
 =MakeDoctor()

 =CloseLocalFiles()
 =CloseTemplates()
 =CloseItems()
 =ClDir()

 SELECT aisoms
* REPLACE s_pred WITH m.s_pred, sum_flk WITH 0, paz WITH m.paz, nsch WITH m.nsch,;
  mee WITH 0, meebad WITH 0, ;
  ambchkdmee WITH 0, dstchkdmee WITH 0, stchkdmee WITH 0, ;
  krank WITH m.krank, paz_dst WITH m.paz_dst, paz_st WITH m.paz_st,;
  sum_kr WITH m.sum_kr, sum_dst WITH m.sum_dst, sum_st WITH m.sum_st,;
  usl_amb WITH m.usl_amb, kd_dst WITH m.kd_dst, ms_st WITH m.ms_st, erz_status WITH 0, erz_id WITH ''
 REPLACE s_pred WITH m.s_pred, sum_flk WITH 0, paz WITH m.paz, nsch WITH m.nsch,;
  mee WITH 0, meebad WITH 0, ;
  ambchkdmee WITH 0, dstchkdmee WITH 0, stchkdmee WITH 0, ;
  krank WITH m.krank, paz_dst WITH m.paz_dst, paz_st WITH m.paz_st,;
  sum_kr WITH m.sum_kr, sum_dst WITH m.sum_dst, sum_st WITH m.sum_st,;
  usl_amb WITH m.usl_amb, kd_dst WITH m.kd_dst, ms_st WITH m.ms_st, erz_status WITH 0, erz_id WITH '' ;
  FOR mcod=m.mcod

 IF ismess=.t.
 MESSAGEBOX(CHR(13)+CHR(10)+'«¿√–”« ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
 ENDIF 

RETURN 

FUNCTION OpenTemplates
 tn_result = 0
 tn_result = tn_result + OpenFile("&ptempl\dxxxx.mmy", "d_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\hxxxx.mmy", "h_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\nvxxxx.mmy", "nv_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\nsxxxx.mmy", "ns_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\rqq.mmy", "r_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sqq.mmy", "s_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\d79sqq.mmy", "d79s_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sprxxxx.mmy", "spr_et", "SHARED")
RETURN tn_result

FUNCTION CloseTemplates
 IF USED('d_et')
  USE IN d_et
 ENDIF 
 IF USED('h_et')
  USE IN h_et
 ENDIF 
 IF USED('nv_et')
  USE IN nv_et
 ENDIF 
 IF USED('ns_et')
  USE IN ns_et
 ENDIF 
 IF USED('r_et')
  USE IN r_et
 ENDIF 
 IF USED('s_et')
  USE IN s_et
 ENDIF 
 IF USED('d79s_et')
  USE IN d79s_et
 ENDIF 
 IF USED('spr_et')
  USE IN spr_et
 ENDIF 
RETURN 

FUNCTION CheckFilesStucture
 IF !AreFilesIdent('dfile', 'd_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(ditem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
 IF !AreFilesIdent('nvfile', 'nv_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(nvitem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
 IF !AreFilesIdent('nsfile', 'ns_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(nsitem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
 IF !AreFilesIdent('rfile', 'r_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(ritem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
 IF !AreFilesIdent('sfile', 's_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(sitem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
 IF !AreFilesIdent('dsfile', 'd79s_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(dsitem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
 IF !AreFilesIdent('sprfile', 'spr_et')
  IF ismess=.t.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈¬≈–Õ¿ﬂ —“–” “”–¿ '+UPPER(spritem)+' ‘¿…À¿!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .f.
 ENDIF 
RETURN .T. 

FUNCTION AreFilesIdent(leftfile,rightfile)
 m.AreFilesEqual = .f.
 fld_1 = AFIELDS(tabl_1, leftfile) && œÓ‚ÂÍ‡ d-Ù‡ÈÎ‡ 
 fld_2 = AFIELDS(tabl_2, rightfile)  && 1 ÒÚÓÎ·Âˆ - Ì‡Á‚‡ÌËÂ, 2 - ÚËÔ,  3 - ‡ÁÏÂÌÓÒÚ¸, 4 - ÌÛÎÂÈ ÔÓÒÎÂ Á‡ÔˇÚÓÈ
 IF fld_1 == fld_2 &&  ÓÎ-‚Ó ÔÓÎÂÈ ÒÓ‚Ô‡‰‡ÂÚ!
  FieldsIdent = CompFields() && 0 - ÂÒÚ¸ ÓÚÎË˜Ëˇ, 1 - ÔÓÎÌÓÂ ÒÓ‚Ô‡‰ÂÌËÂ
  IF FieldsIdent==0
   RETURN .F.
  ENDIF 
 ELSE 
  RETURN .F.
 ENDIF 
RETURN 

FUNCTION CompFields()
 FOR nFld = 1 TO fld_1
  IF (tabl_1(nFld,1) == tabl_2(nFld,1)) AND ;
     (tabl_1(nFld,2) == tabl_2(nFld,2)) AND ;
     (tabl_1(nFld,3) == tabl_2(nFld,3))
  ELSE 
   RETURN 0 
  ENDIF 
 ENDFOR 
RETURN 1

FUNCTION CreateFilesStructure
 CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1,;
   mcod c(7), prmcod c(7), period c(6), d_beg d, d_end d, s_all n(11,2), tip_p n(1), sn_pol c(25), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, ;
   ul n(5), dom c(7), kor c(5), str c(5), kv c(5), d_type c(1), ;
   sv c(3), recid_lpu c(6), IsPr L)
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX on s_all TAG s_all
 USE 

 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), tip c(1), d_u d, pcod c(10), ;
	 otd c(4), cod n(6), k_u n(3), d_type c(1), s_all n(11,2), q c(2),	;
	 novor c(7), ds_s c(7), ds_p c(7), profil c(3), rslt n(3), prvs c(9), ishod n(3),;
	 ord n(2), recid_lpu c(7), fil_id n(6), IsPr L, e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3),;
	 e_period c(6), et c(1), e_ds c(6), e_dtype c(1), koeff n(4,2))

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
 
 CREATE TABLE (Otdel) ;
	(recid c(6),mcod c(7),iotd c(4),name c(40),cnt_bed n(5))
 INDEX ON iotd TAG iotd
 USE 

 CREATE TABLE (Doctor) ;
   (pcod c(10),sn_pol c(25),fam c(25),im c(20),ot c(20),dr d, w n(1),;
	prvd_1 n(4),d_prik_1 d(8),prvs_1 c(9),d_ser_1 d(8),stav_1 n(4,2),iotd_1 c(4),;
	prvd_2 n(4),d_prik_2 d(8),prvs_2 c(9),d_ser_2 d(8),stav_2 n(4,2),iotd_2 c(4),;
	prvd_3 n(4),d_prik_3 d(8),prvs_3 c(9),d_ser_3 d(8),stav_3 n(4,2),iotd_3 c(4),;
	prvd_4 n(4),d_prik_4 d(8),prvs_4 c(9),d_ser_4 d(8),stav_4 n(4,2),iotd_4 c(4),;
	lgot_r c(1),c_ogrn c(15),lpu_id n(6),reserv c(20))
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

 CREATE CURSOR pazamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR pazdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazst (c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i
RETURN 

FUNCTION MakeDoctor
 tnvFile = m.bpath+'\'+nvItem
 oSettings.CodePage('&tnvFile', 866, .t.)
 tnsFile = m.bpath+'\'+nsItem
 oSettings.CodePage('&tnsFile', 866, .t.)
 =OpenFile(tnsFile, 'lcDoctor', 'excl')
 =OpenFile(tnvFile, 'lcDoctor2', 'excl')
 SELECT lcDoctor2
 INDEX on pcod TAG pcod 
 SET ORDER TO pcod 
 SELECT lcDoctor
 SET RELATION TO pcod INTO lcDoctor2
 SCAN 
  SCATTER MEMVAR
  m.prvs_1 = lcDoctor2.prvs_1
  m.prvs_2 = lcDoctor2.prvs_2
  m.prvs_3 = lcDoctor2.prvs_3
  m.prvs_4 = lcDoctor2.prvs_4
  m.prvs_5 = lcDoctor2.prvs_5
  m.prvs_6 = lcDoctor2.prvs_6
  m.d_ser_1 = lcDoctor2.d_ser_1
  m.d_ser_2 = lcDoctor2.d_ser_2
  m.d_ser_3 = lcDoctor2.d_ser_3
  m.d_ser_4 = lcDoctor2.d_ser_4
  m.d_ser_5 = lcDoctor2.d_ser_5
  m.d_ser_6 = lcDoctor2.d_ser_6
  m.ps_1 = lcDoctor2.ps_1
  m.ps_2 = lcDoctor2.ps_2
  m.ps_3 = lcDoctor2.ps_3
  m.ps_4 = lcDoctor2.ps_4
  m.ps_5 = lcDoctor2.ps_5
  m.ps_6 = lcDoctor2.ps_6

  m.dr = CTOD(SUBSTR(m.dr,7,2)+'.'+SUBSTR(m.dr,5,2)+'.'+SUBSTR(m.dr,1,4))

  INSERT INTO Doctor FROM MEMVAR 

 ENDSCAN 
 SET RELATION OFF INTO lcDoctor2
 USE 
 SELECT lcDoctor2
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 USE IN Doctor
RETURN 

FUNCTION MakeOtdel
 tFile = m.bpath+'\'+dItem
 oSettings.CodePage('&tFile', 866, .t.)
 =OpenFile(tFile, 'lcOtdel', 'excl')
 SELECT lcOtdel
 SCAN 
  SCATTER FIELDS EXCEPT mcod MEMVAR
  INSERT INTO Otdel FROM MEMVAR 
 ENDSCAN 
 USE 
 USE IN Otdel
RETURN 

FUNCTION MakePeople
 tFile = m.bpath+'\'+rItem
 oSettings.CodePage('&tFile', 866, .t.)
 =OpenFile(tFile, 'lcRFile', 'excl')
 SELECT lcRFile
 m.paz = 0 
 SCAN 
  SCATTER MEMVAR 
  m.qq = ''
  m.sv = ''
  m.recid_lpu = m.recid
  m.period = m.gcPeriod
  RELEASE m.recid, m.d_beg, m.d_end, m.tip_p, m.s_all
  INSERT INTO People FROM MEMVAR
  m.paz = m.paz + 1
 ENDSCAN 
 USE 
 fso.DeleteFile(m.bpath+'\'+rItem)
RETURN 

FUNCTION MakeTalon
 tFile = m.bpath+'\'+sItem
 oSettings.CodePage('&tFile', 866, .t.)
 tFile = m.bpath+'\'+dsItem
 oSettings.CodePage('&tFile', 866, .t.)
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'tarif', 'shar', 'cod')
 =OpenFile(m.bpath+'\'+sItem, 'lcSFile', 'excl')
 =OpenFile(m.bpath+'\'+dsItem, 'lcDSFile', 'excl')
 SELECT lcDSFile
 INDEX ON recid TAG recid 
 SET ORDER TO recid
 SELECT lcSFile
 SET RELATION TO recid INTO lcDSFile

 SELECT lcSFile
 m.nsch = RECCOUNT('lcSFile')
 SCAN
  SCATTER MEMVAR 
  SCATTER FIELDS lcDSFile.novor, lcDSFile.ds_s, lcDSFile.ds_p, lcDSFile.profil, lcDSFile.rslt,;
   lcDSFile.prvs, lcDSFile.ord, lcDSFile.ishod, lcDSFile.fil_id MEMVAR 
  m.recid_lpu = m.recid
  RELEASE m.recid

  IF IsUsl(m.cod)
   IF !SEEK(m.sn_pol, 'pazamb')
    INSERT INTO pazamb (sn_pol) VALUES (m.sn_pol)
    m.krank = m.krank + 1
   ENDIF 
  ENDIF 
   
  IF IsMes(m.cod) OR IsVMP(m.cod)
   IF !SEEK(m.sn_pol, 'pazst')
    INSERT INTO pazst (c_i) VALUES (m.c_i)
    m.paz_st = m.paz_st + 1
   ENDIF 
  ENDIF 

  IF IsKd(m.cod)
   IF !SEEK(m.sn_pol, 'pazdst')
    INSERT INTO pazdst (sn_pol) VALUES (m.sn_pol)
    m.paz_dst = m.paz_dst + 1
   ENDIF 
  ENDIF 

  m.otd = m.iotd
  m.s_all = fsumm(m.cod, m.tip, m.k_u, m.IsVed)
  m.q = ''
  m.period = m.gcPeriod
   m.s_pred = m.s_pred + s_all
  
  IF OCCURS(' ',ALLTRIM(m.pcod)) > 0 && —ÓÒÚ‡‚ÌÓÈ ÍÓ‰ ‚‡˜‡
   m.pcod  = SUBSTR(ALLTRIM(m.pcod),1,AT(' ',ALLTRIM(m.pcod))-1)
*   m.docvs = SUBSTR(ALLTRIM(m.pcod),AT(' ',ALLTRIM(m.pcod))+1)
  ELSE 
   m.pcod  = ALLTRIM(LEFT(ALLTRIM(m.pcod),10))
*   m.docvs = ''
  ENDIF 
  
  INSERT INTO Talon FROM MEMVAR 
*  INSERT INTO Talon_sv FROM MEMVAR 
 ENDSCAN 
 SET RELATION OFF INTO lcDSFile
 USE 
 SELECT lcDSFile
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 fso.DeleteFile(m.bpath+'\'+sItem)
 fso.DeleteFile(m.bpath+'\'+dsItem)
 USE IN Tarif

 SELECT sn_pol, 1 AS tip_p, MIN(d_u) as min_p, MAX(d_u) as max_p, SUM(s_all) as s_all FROM talon WHERE EMPTY(tip) ;
   GROUP BY sn_pol INTO CURSOR intp
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 SELECT sn_pol, 2 AS tip_p, MIN(d_u-k_u) as min_s, MAX(d_u) as max_s, SUM(s_all) as s_all FROM talon GROUP BY sn_pol ;
  WHERE !EMPTY(tip) INTO CURSOR ints
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 USE IN talon
 SELECT people
 SET RELATION TO sn_pol INTO intp
 SET RELATION TO sn_pol INTO ints ADDITIVE 
* SET ORDER TO unkey IN people_sv
* SET RELATION TO mcod+sn_pol INTO people_sv ADDITIVE 
* SET ORDER TO sn_pol IN people_sv
* SET RELATION TO sn_pol INTO people_sv ADDITIVE 
 SCAN 
  m.d_beg = MIN(IIF(!EMPTY(intp.min_p), intp.min_p, m.tdat2), IIF(!EMPTY(ints.min_s), ints.min_s, m.tdat2))
  m.d_end = MAX(intp.max_p, ints.max_s)
  DO CASE 
   CASE intp.tip_p == 1 AND ints.tip_p != 2
    m.tip_p = 1
   CASE intp.tip_p != 1 AND ints.tip_p == 2
    m.tip_p = 2
   CASE intp.tip_p == 1 AND ints.tip_p == 2
    m.tip_p = 3
   OTHERWISE 
    m.tip_p = 0
  ENDCASE 
  REPLACE people.d_beg WITH m.d_beg, people.d_end WITH m.d_end, tip_p WITH m.tip_p
  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
  REPLACE people.s_all WITH m.s_all

*  REPLACE people_sv.d_beg WITH m.d_beg, people_sv.d_end WITH m.d_end, people_sv.tip_p WITH m.tip_p
  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
*  REPLACE people_sv.s_all WITH m.s_all

 ENDSCAN 
* SET RELATION OFF INTO people_sv
 SET RELATION OFF INTO ints
 SET RELATION OFF INTO intp
 USE IN people
 USE IN ints
 USE IN intp 
RETURN 

FUNCTION OpenLocalFiles
 tn_result = 0
 tn_result = tn_result + OpenFile(People, 'People','shar')
 tn_result = tn_result + OpenFile(Talon, 'Talon','shar')
 tn_result = tn_result + OpenFile(Otdel, 'Otdel','shar')
 tn_result = tn_result + OpenFile(Doctor, 'Doctor','shar')
RETURN tn_result

FUNCTION CloseLocalFiles
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('otdel')
  USE IN otdel
 ENDIF 
 IF USED('doctor')
  USE IN doctor
 ENDIF 
RETURN 

FUNCTION ClDir
 IF fso.FileExists(m.bpath+'\'+dItem)
  fso.DeleteFile(m.bpath+'\'+dItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+hItem)
  fso.DeleteFile(m.bpath+'\'+hItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+nvItem)
  fso.DeleteFile(m.bpath+'\'+nvItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+nsItem)
  fso.DeleteFile(m.bpath+'\'+nsItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+rItem)
  fso.DeleteFile(m.bpath+'\'+rItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+sItem)
  fso.DeleteFile(m.bpath+'\'+sItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+dsItem)
  fso.DeleteFile(m.bpath+'\'+dsItem)
 ENDIF 
 IF fso.FileExists(m.bpath+'\'+sprItem)
  fso.DeleteFile(m.bpath+'\'+sprItem)
 ENDIF 
RETURN 

FUNCTION OpenItems(lcdir)
 PRIVATE  lcdir
 m.tnresult=0
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+dItem,  "dfile",  "SHARED")
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+nvItem, "nvfile", "SHARED")
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+nsItem, "nsfile", "SHARED")
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+rItem,  "rfile",  "SHARED")
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+sItem,  "sfile",  "SHARED")
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+dsItem, "dsfile", "SHARED")
 m.tnresult = m.tnresult + OpenFile(lcdir+'\'+sprItem, "sprfile", "SHARED")
RETURN m.tnresult

FUNCTION CloseItems
 IF USED('dfile')
  USE IN dfile
 ENDIF 
 IF USED('hfile')
  USE IN hfile
 ENDIF 
 IF USED('nvfile')
  USE IN nvfile
 ENDIF 
 IF USED('nsfile')
  USE IN nsfile
 ENDIF 
 IF USED('rfile')
  USE IN rfile
 ENDIF 
 IF USED('sfile')
  USE IN sfile
 ENDIF 
 IF USED('dsfile')
  USE IN dsfile
 ENDIF 
 IF USED('sprfile')
  USE IN sprfile
 ENDIF 
RETURN 

FUNCTION IsIpComplete
 m.lIsIpComplete = .t.
 UnzipOpen(m.bpath+'\'+m.bfile)
 DO CASE 
  CASE !UnzipGotoFileByName(dItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+dItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(hItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+hItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(nvItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+nvItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(nsItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+nsItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(rItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+rItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(sItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+sItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(dsItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+dsItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
  CASE !UnzipGotoFileByName(sprItem)
   m.lIsIpComplete = .f.
   IF ismess=.t.
   MESSAGEBOX(CHR(13)+CHR(10)+'¬ ¿–’»¬ÕŒÃ ‘¿…À≈ '+UPPER(m.bfile)+CHR(13)+CHR(10)+' Œ“—”“—“¬”≈“ '+sprItem+' ‘¿…À!'+CHR(13)+CHR(10),0+16,'ÕÂÍÓÏÔÎÂÍÚ')
   ENDIF 
 ENDCASE 
 UnzipClose()
RETURN m.lIsIpComplete

FUNCTION UnZipItems
 ZipName = m.bpath+'\'+m.bfile
 IF !fso.FolderExists(ZipDir)
  fso.CreateFolder(ZipDir)
 ENDIF 

 UnzipOpen(ZipName)
 UnzipGotoFileByName(rItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(sItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(dsItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(dItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(hItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(nvItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(nsItem)
 UnzipFile(ZipDir+'\')
 UnzipGotoFileByName(sprItem)
 UnzipFile(ZipDir+'\')
 UnzipClose()
RETURN 