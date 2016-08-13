FUNCTION MakeEKMPSS0(lcPath, IsVisible, IsQuit, TipAcc, TipOfExp)

 WAIT "«¿œ”—  WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 
 
 m.lWasUsedSprLpu = IIF(USED('sprlpu'), .T., .F.)
 m.lWasUsedTarif  = IIF(USED('tarif'), .T., .F.)
 m.lWasUsedMkb    = IIF(USED('mkb10'), .T., .F.)
 m.lWasUsedSooKod = IIF(USED('sookod'), .T., .F.)
 m.lWasUsedTalon  = IIF(USED('talon'), .T., .F.)
 m.lWasUsedPeople = IIF(USED('people'), .T., .F.)
 m.lWasUsedDoctor = IIF(USED('doctor'), .T., .F.)
 m.lWasUsedStreet = IIF(USED('street'), .T., .F.)
 m.lWasUsedOtdel  = IIF(USED('otdel'), .T., .F.)
 m.lWasUsedSSActs = IIF(USED('ssacts'), .T., .F.)

* IF !OpenFiles(lcpath)
*  RETURN 
* ENDIF 

 orec = RECNO()
 m.nLocked = 0 
 COUNT FOR ISRLOCKED() TO m.nLocked
 GO (orec)
 
 IF m.nLocked<=0
  polis = sn_pol
  =MakeEKMPSS0One(polis, IsVisible, IsQuit, TipAcc, TipOfExp)
 ELSE 
  orec = RECNO()
  SCAN FOR ISRLOCKED()
   polis = sn_pol
   =MakeEKMPSS0One(polis, IsVisible, IsQuit, TipAcc, TipOfExp)
  ENDSCAN 
  GO (orec)
 ENDIF 

* =CloseFiles()

 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

RETURN 

FUNCTION MakeEKMPSS0One(polis, IsVisible, IsQuit, TipAcc, TipOfExp)

 PRIVATE oal, orec

 m.TipOfExp = TipOfExp
 DO CASE 
  CASE m.TipOfExp = '0'
   m.ctipofexp = 'ÔÛÒÚÓÈ ·Î‡ÌÍ › Ãœ'
  CASE m.TipOfExp = '4'
   m.ctipofexp = 'ÔÎ‡ÌÓ‚‡ˇ › Ãœ'
  CASE m.TipOfExp = '5'
   m.ctipofexp = 'ˆÂÎÂ‚‡ˇ › Ãœ'
  CASE m.TipOfExp = '6'
   m.ctipofexp = 'ÚÂÏ‡ÚË˜ÂÒÍ‡ˇ › Ãœ'
  OTHERWISE 
   m.ctipofexp = ''
 ENDCASE 

 oal = ALIAS()
 orec = RECNO()

 pmee0 = pmee+'\BLANK'
 IF !fso.FolderExists(pmee0)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(pmee0))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 DotName = '¿ÍÚ_› Ãœ_——.dot'

 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ Œ“◊≈“¿'+CHR(13)+CHR(10)+;
   '¿ÍÚ_› Ãœ_ÔÎ‡Ì_——.dot',0+32,'')
  RETURN 
 ENDIF 
 
 oldalias = ALIAS()
* m.sn_pol = sn_pol
 m.sn_pol = polis

 m.mcod       = people.mcod 
 m.lpuid      = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.lpuaddress = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name)+', '+m.mcod, '')
 m.IsVed      = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

 m.exp_dat1   = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2   = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 m.sex = IIF(people.w==1,'ÏÛÊÒÍÓÈ','ÊÂÌÒÍËÈ')
 m.dr  = DTOC(people.dr)
 m.address = 'ÃÓÒÍ‚‡ '
* m.ppolis = STRTRAN(ALLTRIM(m.sn_pol),' ','') && ƒÎˇ Ì‡Á‚‡ÌËˇ ¿ÍÚ‡

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

  SCATTER FIELDS EXCEPT recid,mcod,period,sn_pol,q,novor,ds_s,ds_p,profil,rslt,prvs,ord,ishod,recid_lpu,ispr MEMVAR 

  m.ds     = ds
  m.dsname = IIF(SEEK(m.ds, 'mkb10'), mkb10.name_ds, '')

  DO CASE 
   CASE m.TipAcc == 0 && —‚Ó‰Ì˚È Ò˜ÂÚ
   CASE m.TipAcc == 1 && ¿Ï·ÛÎ‡ÚÓÌ˚È Ò˜ÂÚ
    m.dat1 = MIN(m.d_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
   CASE m.TipAcc == 2 && ƒÌÂ‚ÌÓÈ ÒÚ‡ˆËÓÌ‡
    m.dat1 = MIN(m.d_u-m.k_u, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
   CASE m.TipAcc == 3 && —Ú‡ˆËÓÌ‡
    m.dat1 = MIN(m.d_u-m.k_u+1, m.dat1)
    m.dat2 = MAX(m.d_u, m.dat2)
    m.dslast = m.dslast + k_u
  ENDCASE 

  IF !SEEK(m.ds, 'diags')
   INSERT INTO diags FROM MEMVAR 
  ENDIF 

 ENDSCAN 
 m.dat1 = IIF(m.TipAcc==3, m.dat1, m.dat1)
 m.TipOfMee = IIF(m.TipOfMee==0, 1 ,m.TipOfMee)
 
 SELECT people

 DO CASE 
  CASE m.TipAcc == 0
   m.AddToName = 'Ò‚'
   m.tipakt = 'Ò‚Ó‰'
  CASE m.TipAcc == 1
   m.AddToName = '‡Ï·'
   m.tipakt = '‡Ï·'
  CASE m.TipAcc == 2
   m.AddToName = '‰ÒÚ'
   m.tipakt = '‰Ì/ÒÚ‡ˆ'
  CASE m.TipAcc == 3
   m.AddToName = 'ÒÚ'
   m.tipakt = 'ÒÚ‡ˆ'
 ENDCASE 

 DocName = pmee0+'\'+m.sn_pol
 oDoc = oWord.Documents.Add(pTempl+'\'+DotName)

 m.d_akt = DTOC(DATE())
 m.n_schet = STR(tYear,4)+PADL(tMonth,2,'0')
 IF m.TipAcc == 0
  m.dat1 = IIF(SEEK(m.sn_pol, 'people'), people.d_beg, {})
  m.dat2 = IIF(SEEK(m.sn_pol, 'people'), people.d_end, {})
 ENDIF 

 m.dslast = IIF(!INLIST(m.TipAcc,2,3), IIF(m.dat2-m.dat1>0, m.dat2-m.dat1, 1), m.dslast)
 m.ds     = ds
 m.dsnam  = IIF(SEEK(m.ds, 'mkb10'), ALLTRIM(mkb10.name_ds), '')
 m.pcod   = pcod
 m.docfam = IIF(USED('doctor'), IIF(SEEK(m.pcod, 'doctor'), ALLTRIM(doctor.Fam)+' '+ALLTRIM(doctor.Im)+' '+ALLTRIM(doctor.Ot), ''), '')
 m.fioexp = m.usrfam+' '+m.usrim+' '+m.usrot

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

 oDoc.Bookmarks('ds').Select  
 oWord.Selection.TypeText(m.ds+', '+m.dsnam)

 oDoc.Bookmarks('diagname').Select  
 oWord.Selection.TypeText(m.dsname)

 oDoc.SaveAs(DocName,0)

 SELECT (oal)
 GO (orec)

 WAIT CLEAR 

RETURN 

FUNCTION OpenFiles(ppath)
 IF m.lWasUsedSprLpu = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpu', 'sprlpu', 'shared', 'mcod')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedTarif = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shared', 'cod')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedStreet = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\street', 'street', 'shared', 'ul')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedMkb = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mkb10', 'mkb10', 'shared', 'ds')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedSooKod = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF
 ENDIF 

* pPath = pBase+'\'+gcPeriod+'\'+m.mcod
 
 IF m.lWasUsedTalon = .F.
  IF OpenFile(pPath+'\Talon', 'Talon', 'SHARED', 'sn_pol')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF
 ENDIF  
 
 IF m.lWasUsedPeople = .F.
  IF OpenFile(pPath+'\people', 'people', 'SHARED', 'sn_pol')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 
 
 IF m.lWasUsedDoctor = .F.
  IF OpenFile(pPath+'\doctor', 'doctor', 'SHARED', 'pcod')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 
  
 IF m.lWasUsedOtdel = .F.
  IF OpenFile(pPath+'\Otdel', 'Otdel', 'SHARED', 'iotd')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 

 IF m.lWasUsedSSActs = .F.
  IF OpenFile(pmee+'\ssacts\ssacts', 'ssacts', 'SHARED')>0
   =CloseFiles()
   RETURN .F. 
  ENDIF 
 ENDIF 
RETURN .T.

FUNCTION CloseFiles
 IF m.lWasUsedSprLpu = .F.
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
 ENDIF 
 IF m.lWasUsedTarif = .F.
  IF USED('tarif')
   USE IN tarif 
  ENDIF 
 ENDIF 
 IF m.lWasUsedStreet = .F.
  IF USED('street')
   USE IN street 
  ENDIF 
 ENDIF 
 IF m.lWasUsedMkb = .F.
  IF USED('mkb10')
   USE IN mkb10
  ENDIF 
 ENDIF 
 IF m.lWasUsedSooKod = .F.
  IF USED('sookod')
   USE IN sookod
  ENDIF 
 ENDIF 
 IF m.lWasUsedTalon = .F.
  IF USED('talon')
   USE IN talon 
  ENDIF 
 ENDIF 
 IF m.lWasUsedPeople = .F.
  IF USED('people')
   USE IN people
  ENDIF 
 ENDIF 
 IF m.lWasUsedDoctor = .F.
  IF USED('doctor')
   USE IN doctor
  ENDIF 
 ENDIF 
 IF m.lWasUsedOtdel = .F.
  IF USED('otdel')
   USE IN otdel
  ENDIF 
 ENDIF 
 IF m.lWasUsedSSActs = .F.
  IF USED('ssacts')
   USE IN ssacts
  ENDIF 
 ENDIF 
RETURN 
