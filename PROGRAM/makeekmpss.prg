FUNCTION MakeEKMPSS(lcpolis, lcPath, IsVisible, IsQuit, TipAcc, TipOfExp) && Экспертное заключение ЭКМП

 DotName = 'Акт_ЭКМП_СС.dot'
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

 WAIT "ЗАПУСК WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 

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

 oWord.Visible = .t.
  
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
 
* SELECT talon
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
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name)+', '+m.mcod, '')
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
  d_type c(1),s_all n(11,2),e_cod n(6),e_ku n(3),e_tip c(1),err_mee c(3),e_period c(6),et c(1),docexp c(7))

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
  m.dsname = IIF(SEEK(m.ds, 'mkb10'), mkb10.name_ds, '')

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
  tipacc=m.tipacc AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND IsOk=IIF(m.TipOfMee=1, .t., .f.) AND doctyp='DOC' ;
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
  (m.gcperiod,'DOC',m.mcod,INT(VAL(m.tipofexp)),m.tipacc,IIF(m.TipOfMee=1, .t., .f.),;
   PADR(STRTRAN(m.sn_pol,' ',''),25),people.fam,people.im,people.ot,m.docexp)
  m.nfileid = GETAUTOINCVALUE()
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\ssacts\'+PADL(m.nfileid,6,'0')
  UPDATE ssacts SET actname=PADL(m.nfileid,6,'0')+'.doc', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 IF fso.FileExists(DocName+'.doc')
  oFile = fso.GetFile(DocName+'.doc')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
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
 
 m.n_akt = mcod + m.qcod + PADL(m.tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
 m.d_akt = IIF(m.qcod!='P2', DTOC(DATE()), '')
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

 oDoc = oWord.Documents.Add(pTempl+'\'+DotName)

 oDoc.Bookmarks('n_akt').Select  
 oWord.Selection.TypeText(m.n_akt)
 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('tipofexp').Select  
 oWord.Selection.TypeText(m.ctipofexp)
* oDoc.Bookmarks('d_exp').Select  
* oWord.Selection.TypeText(DATE())
 oDoc.Bookmarks('c_i').Select  
 oWord.Selection.TypeText(ALLTRIM(c_i))
 oDoc.Bookmarks('doc_name').Select  
 oWord.Selection.TypeText(ALLTRIM(m.docfam))
 oDoc.Bookmarks('sn_pol').Select  
 oWord.Selection.TypeText(ALLTRIM(m.sn_pol))
 oDoc.Bookmarks('sex').Select  
 oWord.Selection.TypeText(ALLTRIM(m.sex))
 oDoc.Bookmarks('dr').Select  
 oWord.Selection.TypeText(ALLTRIM(m.dr))
 oDoc.Bookmarks('address').Select  
 oWord.Selection.TypeText(ALLTRIM(m.address))
 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('n_account').Select  
 oWord.Selection.TypeText(m.n_schet)
 oDoc.Bookmarks('ds_last').Select  
 oWord.Selection.TypeText(TRANSFORM(m.dslast, '999'))
 oDoc.Bookmarks('s_lech').Select  
 oWord.Selection.TypeText(TRANSFORM(m.s_lech, '9999999.99'))
 oDoc.Bookmarks('fioexp').Select  
 oWord.Selection.TypeText(m.fioexp+' (код: '+m.docexp+')')

 oDoc.Bookmarks('d_exp').Select  
 oWord.Selection.TypeText(m.d_akt)

 oDoc.Bookmarks('ds').Select  
 oWord.Selection.TypeText(m.ds+', '+m.dsnam)

 nRow = 2
 m.err01 = ''
 SCAN 
  m.cod = cod 
  m.otd = otd
  m.otdname = IIF(SEEK(m.otd,'otdel'), ALLTRIM(otdel.name), '')
  m.pcod = pcod 
  m.doctorname = IIF(SEEK(m.pcod, 'doctor'), PROPER(ALLTRIM(doctor.fam))+' '+;
   PROPER(ALLTRIM(doctor.im))+' '+PROPER(ALLTRIM(doctor.ot))+', '+pcod, '') 
 
  m.err_mee = LEFT(err_mee,2)
  m.eee01 = IIF(SEEK(m.err_mee,'errsmee'), errsmee.osn230+' ('+ALLTRIM(errsmee.f_komment)+')'+CHR(13)+CHR(10),'')
  m.err01 = m.err01 + m.eee01

  oDoc.Tables(1).Cell(nRow,1).Select
  oWord.Selection.InsertRows
  IF IsUsl(m.cod)
   oWord.Selection.TypeText(m.doctorname)
  ELSE 
   oWord.Selection.TypeText(m.otdname)
  ENDIF 
  oDoc.Tables(1).Cell(nRow,2).Select
  oWord.Selection.TypeText(PADL(cod,6,'0'))
  oDoc.Tables(1).Cell(nRow,3).Select
  DO CASE 
   CASE IsMes(m.cod) OR IsVMP(m.cod)
    oWord.Selection.TypeText(IIF(k_u>1, DTOC(d_u-k_u-1), DTOC(d_u-1)))
   CASE IsKD(m.cod)
   CASE IsUsl(m.cod)
    oWord.Selection.TypeText('')
   OTHERWISE 
  ENDCASE 
  oDoc.Tables(1).Cell(nRow,4).Select
  oWord.Selection.TypeText(DTOC(d_u))
  oDoc.Tables(1).Cell(nRow,5).Select
  oWord.Selection.TypeText(STR(k_u,3))

  nRow = nRow + 1
 ENDSCAN 
 oDoc.Bookmarks('err01').Select  
 oWord.Selection.TypeText(m.err01)
 USE 

 IF m.TipOfMee == 1 && Все хорошо!
  m.resume = 'ДЕФЕКТОВ ОФОРМЛЕНИЯ ПЕРВИЧНОЙ МЕДИЦИНСКОЙ ДОКУМЕНТАЦИИ, НАРУШЕНИЙ ПРИ ОКАЗАНИИ МЕДИЦИНСКОЙ ПОМОЩИ НЕ ВЫЯВЛЕНО.'    
 ELSE && Есть ошибки!
  m.resume = REPLICATE('_',255)
 ENDIF 
 
 SELECT diags 
 IF RECCOUNT()<=1
 ELSE 
  SET ORDER TO 
  SCAN 
   IF ds==m.ds
    LOOP 
   ENDIF 
   oDoc.Tables(2).Cell(nRow,1).Select
   oWord.Selection.InsertRows
   oWord.Selection.TypeText(ds+', '+ALLTRIM(dsname))
  ENDSCAN 
 ENDIF 
 USE IN diags 
 
 oDoc.Bookmarks('diagname').Select  
 oWord.Selection.TypeText(m.dsname)

 IF m.TipOfMee == 1 AND m.qcod!='P2' && Все хорошо!
  oDoc.Bookmarks('punkt2').Select  
  oWord.Selection.TypeText('НЕТ')
  oDoc.Bookmarks('punkt3').Select  
  oWord.Selection.TypeText('БЕЗ ЗАМЕЧАНИЙ')
  oDoc.Bookmarks('punkt4').Select  
  oWord.Selection.TypeText('нет')
  oDoc.Bookmarks('punkt5').Select  
  oWord.Selection.TypeText('БЕЗ ЗАМЕЧАНИЙ')
  oDoc.Bookmarks('punkt6').Select  
  oWord.Selection.TypeText('НЕТ')
 ENDIF 

 IF m.qcod!='P2'
 oDoc.Bookmarks('resume').Select  
 oWord.Selection.TypeText(m.resume)
 ENDIF 

 oDoc.Bookmarks('fioexp2').Select  
 oWord.Selection.TypeText(m.fioexp2)

 oDoc.SaveAs(DocName,0)

 SELECT (oldalias)
* GO (orec)

 WAIT CLEAR 

RETURN 
