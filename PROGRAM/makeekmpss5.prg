FUNCTION MakeEKMPSS5(lcpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp) && Акт ЭКМП
 DotName = 'ACT_EKMP_cel.xls'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+;
   m.dotname,0+32,'')
  RETURN 
 ENDIF 

 IF IsUsrDir=.T.
  m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
  IF !fso.FolderExists(m.usrdir)
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SSACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SVACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ELSE 
  IF !fso.FolderExists(pmee)
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ENDIF 

 m.TipOfExp = TipOfExp
 DO CASE 
  CASE m.TipOfExp = '4'
   m.ctipofexp = 'плановая ЭКМП'
  CASE m.TipOfExp = '5'
   m.ctipofexp = 'целевая ЭКМП'
  CASE m.TipOfExp = '6'
   m.ctipofexp = 'тематическая ЭКМП'
  OTHERWISE 
   m.ctipofexp = ''
 ENDCASE 
 
 oal = ALIAS()
 CREATE CURSOR curpaz (sn_pol c(25)) 
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 SELECT &oal

 orecc = RECNO()
 m.nlocked = 0 
 SCAN 
  IF !ISRLOCKED()
   LOOP 
  ENDIF 
  m.sn_pol = sn_pol
  IF !SEEK(m.sn_pol, 'curpaz')
   INSERT INTO curpaz FROM MEMVAR 
   m.nlocked = m.nlocked + 1
  ENDIF 
 ENDSCAN 
 GO (orecc) 
 
 m.IsMulti=0
 IF m.nlocked>0
  IF MESSAGEBOX('ФОРМИРОВАТЬ АКТЫ НА ВСЕХ ОТОБРАННЫХ?'+CHR(13)+CHR(10),4+32,'')=6
   m.IsMulti=1
  ELSE
   m.IsMulti=0
  ENDIF 
 ENDIF 

 IF m.IsMulti=1
  MESSAGEBOX('ВЫ ВЫБРАЛИ ПЕЧАТЬ ВСЕХ ОТОБРАННЫХ!'+CHR(13)+CHR(10),0+64,'')
  SELECT curpaz 
  SCAN 
   m.lpolis = sn_pol
   =MakeEKMPSS1(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)
  ENDSCAN 
 ELSE 
  m.lpolis = sn_pol
  =MakeEKMPSS1(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)
 ENDIF 
 
 SELECT &oal
 GO (orecc) 

RETURN 
 
FUNCTION MakeEKMPSS1(lpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)
 
 m.expname = 'Акт '
 DO CASE 
  CASE m.tipofexp = '4'
   m.expname = m.expname + 'плановой'
  CASE m.tipofexp = '5'
   m.expname = m.expname + 'целевой'
  CASE m.tipofexp = '6'
   m.expname = m.expname + 'тематической'
 ENDCASE 
 m.expname = m.expname + ' экспертизы качества медицинской помощи'

 oldalias = ALIAS()
 m.sn_pol = lpolis

 SELECT merror
 COUNT FOR sn_pol = m.sn_pol AND merror.et = m.TipOfExp TO m.nIsExps

 SELECT talon
* COUNT FOR sn_pol = m.sn_pol AND SEEK(recid, 'merror', 'recid') AND merror.et = m.TipOfExp TO m.nIsExps

 IF m.nIsExps == 0
  IF m.IsMulti=0
   MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ СЧЕТУ ЭКМП НЕ ПРОВОДИЛОСЬ!'+CHR(13)+CHR(10),0+64,'')
  ENDIF 
  SELECT (oldalias)
  RETURN
 ENDIF 

 m.mcod       = people.mcod 
 m.lpuid      = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.lpuaddress = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod, '')
 m.IsVed      = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.d_beg      = DTOC(people.d_beg)
 m.d_end      = DTOC(people.d_end)

 m.exp_dat1 = '01.'+PADL(m.tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 m.sex     = IIF(people.w==1,'мужской','женский')
 m.dr      = DTOC(people.dr)
 m.address = 'Москва'
 m.ppolis = STRTRAN(ALLTRIM(m.sn_pol),' ','') && Для названия Акта
 

 CREATE CURSOR ttalon (sn_pol c(25),c_i c(30),ds c(6),tip c(1),d_u d,pcod c(10),otd c(4),cod n(6),k_u n(3),;
  d_type c(1),s_all n(11,2),e_cod n(6),e_ku n(3),e_tip c(1),err_mee c(3),e_period c(6),et c(1),docexp c(7), ;
  s_1 n(11,2), s_2 n(11,2))

 CREATE CURSOR diags (ds c(6), dsname c(160))
 INDEX on ds TAG ds
 SET ORDER TO ds

 SELECT talon 

 m.TipOfMee = 0
 
 m.s_lech = 0
 m.dat1 = {31.12.2099}
 m.dat2 = {01.01.2000}
 m.dslast = 0
 SCAN FOR sn_pol == m.sn_pol

  IF SEEK(recid, 'serror')
   LOOP 
  ENDIF 
  IF !SEEK(PADL(recid,6,'0')+m.TipOfExp, 'merror', 'id_et')
   LOOP 
  ENDIF 
 
  SCATTER FIELDS EXCEPT recid,mcod,period,sn_pol,q,novor,ds_s,ds_p,profil,rslt,prvs,ord,ishod,recid_lpu,ispr MEMVAR 

  m.e_sall = 0
  m.s_lech = m.s_lech + m.s_all
  m.ds     = ds
  m.dsname = IIF(SEEK(m.ds, 'mkb10'), m.ds+', '+mkb10.name_ds, '')

  m.et       = merror.et
  m.ee       = merror.ee
  m.e_cod    = merror.e_cod
  m.e_ku     = merror.e_ku
  m.e_tip    = merror.e_tip
  m.err_mee  = merror.err_mee 
  m.osn230   = merror.osn230
  m.koeff    = merror.koeff
  m.e_period = merror.e_period
  m.straf    = merror.straf
  m.s_1      = merror.s_1
  m.s_2      = merror.s_2
  
  m.docexp   = merror.docexp

  DO CASE 
   CASE m.TipAcc == 0 && Сводный счет
   CASE m.TipAcc == 1 && Амбулаторный счет
    m.dat1 = MIN(m.d_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
   CASE m.TipAcc == 2 && Дневной стационар
    m.dat1 = MIN(m.d_u-m.k_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
   CASE m.TipAcc == 3 && Стационар
    m.dat1 = MIN(m.d_u-m.k_u+1, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
  ENDCASE 

  DO CASE 
   CASE EMPTY(err_mee)
   CASE LEFT(err_mee,2)=='W0'
    m.TipOfMee = IIF(m.TipOfMee<=1, 1, m.TipOfMee)
   OTHERWISE 
    m.TipOfMee = 2 && Есть ошибки
*    IF (!EMPTY(m.e_cod) AND m.cod != m.e_cod) OR (!EMPTY(m.e_ku) AND m.k_u != m.e_ku) OR (!EMPTY(m.e_tip) AND m.e_tip != m.tip)
*     m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
     m.e_sall = s_1
*    ENDIF 
  ENDCASE 

  INSERT INTO ttalon FROM MEMVAR 

  IF !SEEK(m.ds, 'diags')
   INSERT INTO diags FROM MEMVAR 
  ENDIF 

 ENDSCAN 

 m.fioexp   = IIF(SEEK(m.docexp, 'explist'), ;
  ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot), '')
 m.fioexp2   = IIF(SEEK(m.docexp, 'explist'), ;
  ALLTRIM(explist.fam)+' '+LEFT(ALLTRIM(explist.im),1)+'.'+LEFT(ALLTRIM(explist.ot),1)+'.', '')

 m.dat1 = IIF(m.TipAcc==3, m.dat1, m.dat1)
 m.TipOfMee = IIF(m.TipOfMee==0, 1 ,m.TipOfMee)
 
 SELECT people

 DO CASE 
  CASE m.TipAcc == 0
   m.AddToName = 'св'
   m.tipakt = 'свод'
  CASE m.TipAcc == 1
   m.AddToName = 'амб'
   m.tipakt = 'амб'
  CASE m.TipAcc == 2
   m.AddToName = 'дст'
   m.tipakt = 'дн/стац'
  CASE m.TipAcc == 3
   m.AddToName = 'ст'
   m.tipakt = 'стац'
 ENDCASE 

 ooal = ALIAS()
 SELECT recid FROM ssacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) AND ;
  tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND IsOk=IIF(m.TipOfMee=1, .t., .f.) AND doctyp='XLS';
  AND docexp=m.docexp INTO CURSOR rqwest NOCONSOLE 
 m.nfileid = recid
 USE 
 SELECT (ooal)
 
 IF m.nfileid>0
* DocName = pmee+IIF(m.TipOfMee == 1,'\ActEKMPPlanOK_','\ActEKMPPlan_')+m.mcod+'_'+m.ppolis+'_'+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)+m.AddToName
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO ssacts (period,doctyp,mcod,codexp,tipacc,isok,sn_pol,fam,im,ot,docexp) ;
   VALUES ;
  (m.gcperiod,'XLS',m.mcod,INT(VAL(m.tipofexp)),m.tipacc,IIF(m.TipOfMee=1, .t., .f.),;
   PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,m.docexp)
  m.nfileid = GETAUTOINCVALUE()
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
  UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 IF fso.FileExists(DocName+'.xls')
  oFile = fso.GetFile(DocName+'.xls')
  m.DateCreated      = TTOC(oFile.DateCreated)
  m.DateLastAccessed = TTOC(oFile.DateLastAccessed)
  m.DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF  m.IsMulti=0

  IF MESSAGEBOX('ПО ВЫБРАННОМУ СЧЕТУ АКТ УЖЕ ФОРМИРОВАЛСЯ!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА СОЗДАНИЯ АКТА            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ОТКРЫТИЯ АКТА : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ИЗМЕНЕНИЯ АКТА: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ ПЕРЕФОРМИРОВАТЬ АКТ?',4+32,'') == 7 

   USE IN ttalon
   USE IN diags
   RETURN
  ELSE 

   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid
   
  ENDIF 

  ELSE 

   UPDATE ssacts SET actdate=DATETIME() WHERE recid = m.nfileid

  ENDIF 

 ENDIF 

 SELECT ttalon
* BROWSE 
 m.d_akt = DTOC(DATE())
 m.n_akt = mcod + m.qcod + PADL(m.tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid)) + ' от ' + m.d_akt
 m.n_schet = STR(tYear,4)+PADL(m.tMonth,2,'0')
 IF m.TipAcc == 0
  m.dat1 = IIF(SEEK(m.sn_pol, 'people'), people.d_beg, {})
  m.dat2 = IIF(SEEK(m.sn_pol, 'people'), people.d_end, {})
 ENDIF 
 m.dslast = IIF(!INLIST(m.TipAcc,2,3), IIF(m.dat2-m.dat1>0, m.dat2-m.dat1, 1), m.dslast)
 m.ds = ds
 m.dsnam = IIF(SEEK(m.ds, 'mkb10'), ALLTRIM(mkb10.name_ds), '')
 m.pcod = pcod
 m.docfam = IIF(SEEK(m.pcod, 'doctor'), ALLTRIM(doctor.Fam)+' '+ALLTRIM(doctor.Im)+' '+ALLTRIM(doctor.Ot), '')
 m.c_i = ALLTRIM(c_i)
 
 SELECT ttalon 
 m.totdefs = 0
 m.defs = ''
 IF m.TipOfMee == 1 && Все хорошо!
  nRow = 2
  m.tot_badsum = 0
  m.tot_goodsum = 0
  m.tot_straf   = 0
  SCAN 
   m.cod = cod 
   m.otd = otd
   m.otdname = IIF(SEEK(m.otd,'otdel'), ALLTRIM(otdel.name), '')
   m.pcod = pcod 
   m.doctorname = IIF(SEEK(m.pcod, 'doctor'), PROPER(ALLTRIM(doctor.fam))+' '+;
    PROPER(ALLTRIM(doctor.im))+' '+PROPER(ALLTRIM(doctor.ot))+', '+pcod, '') 
   m.tot_goodsum = m.tot_goodsum + s_all
   nRow = nRow + 1
  ENDSCAN 
*  USE 
  m.koplate   = m.tot_goodsum
  m.nekoplate = m.tot_badsum
  m.resume = 'ДЕФЕКТОВ ОФОРМЛЕНИЯ ПЕРВИЧНОЙ МЕДИЦИНСКОЙ ДОКУМЕНТАЦИИ, НАРУШЕНИЙ ПРИ ОКАЗАНИИ МЕДИЦИНСКОЙ ПОМОЩИ НЕ ВЫЯВЛЕНО.'    

 ELSE && Есть ошибки!
  m.tot_badsum  = 0
  m.tot_goodsum = 0
  m.tot_straf   = 0
  SCAN 
   m.cod = cod 
   m.er_c = err_mee
   m.osn230 = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), sookod.osn230, '')	
   m.osn230name = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), ALLTRIM(sookod.f_komment), '')	
   m.defs = m.defs+m.osn230+'('+m.osn230name+')'+';'+CHR(13)+CHR(10)
   m.totdefs = m.totdefs + IIF(LEFT(m.er_c,2)!='W0', 1, 0)
   m.otd = otd
   m.otdname = IIF(SEEK(m.otd,'otdel'), ALLTRIM(otdel.name), '')
   m.pcod = pcod 
   m.doctorname = IIF(SEEK(m.pcod, 'doctor'), PROPER(ALLTRIM(doctor.fam))+' '+;
    PROPER(ALLTRIM(doctor.im))+' '+PROPER(ALLTRIM(doctor.ot))+', '+pcod, '') 
   m.s_1 = s_1
   m.s_2 = s_2
   m.tot_badsum = m.tot_badsum + m.s_1
   m.tot_straf = m.tot_straf + m.s_2
  ENDSCAN 
*  USE 
  m.koplate   = m.s_lech - m.tot_badsum
  m.nekoplate = m.tot_badsum

  m.resume = REPLICATE('_',255)
 ENDIF 
 
* m.tot_straf = 0
 m.docfio  = m.usrfam+' '+m.usrim+' '+m.usrot
 
 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'

 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 
 USE IN ttalon 

 SELECT (oldalias)
* GO (orec)

 WAIT CLEAR 

RETURN 
