FUNCTION MakeMEESv(lcPath, IsVisible, IsQuit, TipOfExp, TipOfPeriod)
 
 m.tipofext = tipofexp
 m.TipOfPeriod = TipOfPeriod && 0-локальный период, 1 - сводный!
 
 DotName = 'Акт_МЭЭ_свод.dot'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+;
   'Акт_МЭЭ_свод.dot',0+32,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.mcod  = mcod 
 IF m.TipOfPeriod=0
  pPath = pBase+'\'+gcPeriod+'\'+m.mcod
  TFile = 'talon'
  mFile = 'm'+m.mcod
 ELSE 
  pPath = pBase+'\'+gcPeriod+'\0000000\'+m.mcod
  TFile = 't'+flcod
  mFile = 'm'+flcod
 ENDIF 

 IF OpenFile(pPath+'\'+TFile, 'Talon', 'SHARED', 'recid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pPath+'\'+mFile, 'merror', 'SHARED')>0
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 
 oal = ALIAS()
 SELECT merror
 m.nexps=0
 COUNT FOR !EMPTY(err_mee) AND et=m.tipofexp TO m.nexps
 IF m.nexps=0
  MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ ЛПУ ЭКСПЕРТИЗА'+CHR(13)+CHR(10)+;
   'ЗАДАННОГО ТИПА НЕ ПРОВОДИЛАСЬ!',0+64,'')
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT &oal
  RETURN 
 ENDIF 
 SELECT &oal

 m.lWasUsedTarif=.t.
 IF !USED('tarif')
  m.lWasUsedTarif=.f.
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shared', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   RETURN 
  ENDIF 
 ENDIF 

 WAIT "ЗАПУСК WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 
 
 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)
 m.edat1    = DTOC(DATE())
 m.edat2    = m.edat1  

 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name)+', '+m.mcod, '')
 m.fioexp  = m.usrfam+' '+m.usrim+' '+m.usrot

 ooal = ALIAS()
 IF m.TipOfPeriod=0

 SELECT recid FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
  INTO CURSOR rqwest NOCONSOLE  
 m.nfileid = recid
 USE 
 SELECT (ooal)
 
 IF m.nfileid>0
  DocName = pmee+'\svacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO svacts (period,mcod,codexp) ;
   VALUES ;
  (m.gcperiod,m.mcod,INT(VAL(m.tipofexp)))
  m.nfileid = GETAUTOINCVALUE()
  DocName = pmee+'\svacts\'+PADL(m.nfileid,6,'0')
  UPDATE svacts SET actname=PADL(m.nfileid,6,'0')+'.doc', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 ELSE 
 
 DocName = pBase+'\'+gcPeriod+'\0000000\'+m.mcod+'\mee'+TipOfExp+'sv'+flcod

 ENDIF 
  
 IF fso.FileExists(DocName+'.doc')
  oFile = fso.GetFile(DocName+'.doc')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF MESSAGEBOX('ПО ВЫБРАННОМУ ЛПУ АКТ УЖЕ ФОРМИРОВАЛСЯ!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА СОЗДАНИЯ АКТА            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ОТКРЫТИЯ АКТА : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ИЗМЕНЕНИЯ АКТА: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ ПЕРЕФОРМИРОВАТЬ АКТ?',4+32,'') == 7 
   
   oWord.Quit
   USE IN tarif
   USE IN talon 
   USE IN merror
   SELECT aisoms
   RETURN
  ELSE

   IF m.TipOfPeriod=0
   UPDATE svacts SET actdate=DATETIME() WHERE recid = m.nfileid
   ENDIF 
   
  ENDIF 

 ENDIF 

 m.IsExpMee = .f.

 m.checked_tot = 0
 m.checked_amb = 0
 m.checked_dst = 0
 m.checked_st  = 0

 m.bad_kol   = 0
 m.bad_sum   = 0
 m.opl_tot   = 0
 m.vzaim_tot = 0
 m.tot_straf = 0
 nRow        = 4
* IF FIELD('SENT')='SENT'
*  m.dschet    = TTOC(sent)+', номер счета '+STR(tYear,4)+PADL(tMonth,2,'0')
* ELSE 
*  m.dschet    = TTOC(DATETIME())+', номер счета '+STR(tYear,4)+PADL(tMonth,2,'0')
* ENDIF 

 m.dschet    = TTOC(aisoms.sent)+', номер счета '+STR(tYear,4)+PADL(tMonth,2,'0')
  
 CREATE CURSOR qwert (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR qwertamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertst (c_i c(30))
 INDEX on c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR qwertdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertbad (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT merror
 SET RELATION TO recid INTO talon 
 
 SCAN 

  m.sn_pol = talon.sn_pol
  m.cod    = cod
  m.c_i    = talon.c_i

  IF et!=m.TipOfExp
   LOOP 
  ENDIF 

  IF m.IsExpMee = .f.

   m.IsExpMee = .t.

   oDoc = oWord.Documents.Add(pTempl+'\'+DotName)

   IF m.TipOfPeriod=0
   m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
   ELSE 
   m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'
   ENDIF 
   oDoc.Bookmarks('n_akt').Select  
   oWord.Selection.TypeText(m.n_akt)
   m.d_akt = DTOC(DATE())
   oDoc.Bookmarks('d_akt').Select  
   oWord.Selection.TypeText(m.d_akt)
   oDoc.Bookmarks('fioexp').Select  
   oWord.Selection.TypeText(m.fioexp)
   oDoc.Bookmarks('smo_name').Select  
   oWord.Selection.TypeText(m.qname)
   oDoc.Bookmarks('lpu_name').Select  
   oWord.Selection.TypeText(m.lpuname)
   oDoc.Bookmarks('exp_dat1').Select  
   oWord.Selection.TypeText(m.exp_dat1)
   oDoc.Bookmarks('exp_dat2').Select  
   oWord.Selection.TypeText(m.exp_dat2)
   oDoc.Bookmarks('edat1').Select  
   oWord.Selection.TypeText(m.edat1)
   oDoc.Bookmarks('edat2').Select  
   oWord.Selection.TypeText(m.edat2)
   oDoc.Bookmarks('dschet').Select  
   oWord.Selection.TypeText(m.dschet)
  ENDIF 
   
  IF !SEEK(m.sn_pol, 'qwert')
   INSERT INTO qwert (sn_pol) VALUES (m.sn_pol)
   m.checked_tot = m.checked_tot + 1
  ENDIF 
   
  IF IsUsl(m.cod) AND !SEEK(m.sn_pol, 'qwertamb')
   INSERT INTO qwertamb (sn_pol) VALUES (m.sn_pol)
   m.checked_amb = m.checked_amb + 1
  ENDIF 
   
  IF IsKD(m.cod) AND !SEEK(m.sn_pol, 'qwertdst')
   INSERT INTO qwertdst (sn_pol) VALUES (m.sn_pol)
   m.checked_dst = m.checked_dst + 1
  ENDIF 
   
  IF (IsMes(m.cod) OR IsVMP(m.cod)) AND !SEEK(m.c_i, 'qwertst')
   INSERT INTO qwertst (c_i) VALUES (m.c_i)
   m.checked_st = m.checked_st + 1
  ENDIF 

  m.er_c   = err_mee
  m.osn230 = osn230

  oDoc.Tables(1).Cell(nRow,1).Select 
  oWord.Selection.InsertRows
  oWord.Selection.TypeText(STR(m.checked_tot,3))
  oDoc.Tables(1).Cell(nRow,2).Select && Полис
  oWord.Selection.TypeText(ALLTRIM(m.sn_pol))
  oDoc.Tables(1).Cell(nRow,3).Select && Карта
  oWord.Selection.TypeText(ALLTRIM(m.c_i))
  oDoc.Tables(1).Cell(nRow,4).Select && Начало обращения
  oWord.Selection.TypeText(IIF(!IsMes(m.cod) and !IsVMP(m.cod), DTOC(talon.d_u),DTOC(talon.d_u-talon.k_u+1)))
  oDoc.Tables(1).Cell(nRow,5).Select && Конец обращения
  oWord.Selection.TypeText(DTOC(talon.d_u))
  oDoc.Tables(1).Cell(nRow,6).Select && Код МКБ
  oWord.Selection.TypeText(ALLTRIM(talon.ds))
  oDoc.Tables(1).Cell(nRow,7).Select && Код услуги
  oWord.Selection.TypeText(PADL(m.cod,6,'0'))
  oDoc.Tables(1).Cell(nRow,9).Select && Код дефекта по Пр.230
  oWord.Selection.TypeText(m.osn230)
  oDoc.Tables(1).Cell(nRow,10).Select && Код ошибки
  oWord.Selection.TypeText(IIF(LEFT(UPPER(m.er_c),2)!='W0', m.er_c, ''))
     
  m.ns_all = 0
  m.delta  = 0

  m.opl_tot   = m.opl_tot + s_all && Изменено 12.09.12 по замечанию СОГАЗ - неверная сумма в сводном акте "к оплате"

  IF LEFT(UPPER(err_mee),2) != 'W0'
   IF !SEEK(m.sn_pol, 'qwertbad')
    INSERT INTO qwertbad (sn_pol) VALUES (m.sn_pol)
    m.bad_kol = m.bad_kol+1
   ENDIF 

   IF koeff<=0 && Старый механизм

    IF EMPTY(e_cod) AND EMPTY(e_tip) AND EMPTY(e_ku)
     m.bad_sum = m.bad_sum + s_all
     m.delta = s_all
     m.vzaim_tot = m.vzaim_tot + m.delta
    ENDIF  && Полное снятие!

    IF (!EMPTY(e_cod) AND cod != e_cod) OR (!EMPTY(e_ku) AND k_u != e_ku) ;
     OR (!EMPTY(e_tip) AND e_tip != tip)
     m.ns_all = fsumm(e_cod, e_tip, e_ku, m.IsVed)
     m.delta = s_all - m.ns_all
     m.bad_sum = m.bad_sum + m.delta
     m.vzaim_tot = m.vzaim_tot + m.delta
    ENDIF && Частичное снятие

   ELSE && Новый механизм

    m.bad_sum = m.bad_sum + s_all*koeff
    m.delta = s_all*koeff
    m.vzaim_tot = m.vzaim_tot + m.delta
    m.tot_straf = m.tot_straf + straf*m.ynorm

   ENDIF 

  ENDIF 

  m.s_all = s_all
  oDoc.Tables(1).Cell(nRow,8).Select && Оплачено за услуги
  oWord.Selection.TypeText(TRANSFORM(m.s_all, '9999999.99'))
  oDoc.Tables(1).Cell(nRow,11).Select && Размер взаимозачета
  oWord.Selection.TypeText(TRANSFORM(m.delta, '9999999.99'))
  oDoc.Tables(1).Cell(nRow,12).Select && Штраф
  oWord.Selection.TypeText(TRANSFORM(straf*m.ynorm, '9999.99'))
  oDoc.Tables(1).Cell(nRow,13).Select && Примечание
  oWord.Selection.TypeText(IIF(LEFT(UPPER(err_mee),2) == 'W0', 'Без замечаний', ''))
  nRow = nRow + 1

 ENDSCAN 
 SET RELATION OFF INTO talon 

 USE IN talon 
 USE IN merror
 USE IN qwert
 USE IN qwertamb
 USE IN qwertst
 USE IN qwertdst
  
 SELECT (ooal)

 m.checked_tot = m.checked_amb + m.checked_st + m.checked_dst

 IF m.IsExpMee = .t.
  oDoc.Bookmarks('checked_tot').Select  
  oWord.Selection.TypeText(TRANSFORM(m.checked_tot,'99999'))
  oDoc.Bookmarks('checked_amb').Select  
  oWord.Selection.TypeText(TRANSFORM(m.checked_amb,'99999'))
  oDoc.Bookmarks('checked_st').Select  
  oWord.Selection.TypeText(TRANSFORM(m.checked_st,'99999'))
  oDoc.Bookmarks('checked_dst').Select  
  oWord.Selection.TypeText(TRANSFORM(m.checked_dst,'99999'))

  oDoc.Bookmarks('bad_kol').Select  
  oWord.Selection.TypeText(TRANSFORM(m.bad_kol, '9999999'))
  oDoc.Bookmarks('bad_sum').Select  
  oWord.Selection.TypeText(TRANSFORM(m.bad_sum, '9999999.99'))

  oDoc.Bookmarks('opl_tot').Select  
  oWord.Selection.TypeText(TRANSFORM(m.opl_tot,'9999999.99'))
  oDoc.Bookmarks('vzaim_tot').Select  
  oWord.Selection.TypeText(TRANSFORM(m.vzaim_tot,'9999999.99'))
  oDoc.Bookmarks('straf_tot').Select  
  oWord.Selection.TypeText(TRANSFORM(m.tot_straf,'9999999.99'))
  
  oDoc.Bookmarks('fioexp2').Select  
  oWord.Selection.TypeText(m.fioexp)

  oDoc.SaveAs(DocName,0)
  IF IsVisible == .F.
   oDoc.Close(0)
  ENDIF 
 ENDIF 
 
 IF m.lWasUsedTarif = .f.
  USE IN tarif 
 ENDIF 
 
 SELECT aisoms

 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

RETURN 