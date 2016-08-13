FUNCTION MtPrn(lcPath, IsVisible, IsQuit)

 m.mcod  = mcod
 m.lpuid = lpuid
 m.lpuname  = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 
 m.period = NameOfMonth(tMonth)+ ' ' + STR(tYear,4)
 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa = s_pred

 m.usl_all=0
 m.usl_all_ok = 0
 m.pr_sum = 0
 m.usl_all_bad = 0
 m.sum_all_bad = 0
 m.sum_all_bad = 0
 m.sum_all_ok = 0

 DotName = pTempl + "\MtxxxxQQmmy.dot"
 DocName = lcPath + "\Mt" + STR(m.lpuid,4) + m.qcod + m.mmy

 eeFile = 'e'+m.mcod
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)

 m.n_aktmek = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 oDoc.Bookmarks('n_aktmek').Select  
 oWord.Selection.TypeText(m.n_aktmek)
 
 m.d_aktmek = DTOC(DATE())
 oDoc.Bookmarks('d_aktmek').Select  
 oWord.Selection.TypeText(m.d_aktmek)
 
 m.akt_month = NameOfMonth(tMonth)
 oDoc.Bookmarks('akt_month').Select  
 oWord.Selection.TypeText(m.akt_month)

 m.akt_year = STR(tYear,4)
 oDoc.Bookmarks('akt_year').Select  
 oWord.Selection.TypeText(m.akt_year)
 
 m.nr_akt = m.mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 oDoc.Bookmarks('nr_akt').Select  
 oWord.Selection.TypeText(m.nr_akt)
 
 oDoc.Bookmarks('lpu_namemek').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('mcod').Select  
 oWord.Selection.TypeText(m.mcod)

 m.smo_name = m.qname
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.smo_name)
 oDoc.Bookmarks('qq').Select  
 oWord.Selection.TypeText(m.qcod)


 m.n_akt = mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 oDoc.Bookmarks('n_akt').Select  
 oWord.Selection.TypeText(m.n_akt)
 
 m.d_akt = DTOC(DATE())
 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)

 m.n_mek = PADL(tMonth,2,'0') + '/' + STR(tYear,4)
 oDoc.Bookmarks('n_mek').Select  
 oWord.Selection.TypeText(m.n_mek)

 m.d_mek = DTOC(TTOD(sent))
 oDoc.Bookmarks('d_mek').Select  
 oWord.Selection.TypeText(m.d_mek)

 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)

 USE &lcPath\Talon     IN 0 ALIAS Talon  SHARED 
 USE &lcPath\People    IN 0 ALIAS People SHARED ORDER sn_pol 
 USE &lcPath\&eeFile   IN 0 ALIAS sError SHARED ORDER rid 
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx' IN 0 ALIAS sookod SHARED ORDER er_c
 USE pcommon+'\prv002xx' IN 0 ALIAS prv002 SHARED ORDER profil 
 
 SELECT sError
 SET RELATION TO LEFT(c_err,2) INTO sookod
 SELECT Talon 
 SET RELATION TO sn_pol INTO people 
 SET RELATION TO RecId INTO sError ADDITIVE 

 nRow      = 2
 nError    = 1
 m.sum_def = 0
 
 CREATE CURSOR prt2 (osn230 c(9), ssum n(11,2))
 SELECT prt2
 INDEX on osn230 TAG osn230
 SET ORDER TO osn230
 
 SELECT Talon 
 
 m.usl_all = 0

 SCAN 
  m.cod = cod
  m.d_type = d_type

  m.usl_all = m.usl_all + 1

  m.pr_sum = m.pr_sum + s_all

  IF EMPTY(sError.rid)
   m.usl_all_ok = m.usl_all_ok + 1
   m.sum_all_ok = m.sum_all_ok + s_all
  ELSE 
   m.usl_all_bad = m.usl_all_bad + 1
   m.sum_all_bad = m.sum_all_bad + s_all
  ENDIF 
  
  IF !EMPTY(sError.rid)
   m.ssum = s_all
   m.osn230 = sookod.osn230

   IF !SEEK(sookod.osn230, 'prt2')
    INSERT INTO prt2 (osn230, ssum) VALUES (m.osn230, m.ssum)
   ELSE 
    UPDATE prt2 SET ssum = ssum + m.ssum WHERE osn230=m.osn230
   ENDIF 

   oDoc.Tables(4).Rows(nRow).Select 
   oWord.Selection.InsertRows

   oDoc.Tables(4).Rows(nRow).Cells(1).Select && Номер по порядку
   oWord.Selection.TypeText(ALLTRIM(STR(nError)))
   oDoc.Tables(4).Rows(nRow).Cells(2).Select && Полис
   oWord.Selection.TypeText(ALLTRIM(sn_pol))
   oDoc.Tables(4).Rows(nRow).Cells(3).Select && Диагноз
   oWord.Selection.TypeText(ALLTRIM(ds))
   oDoc.Tables(4).Rows(nRow).Cells(4).Select && Дата начала лечения
   oWord.Selection.TypeText(DTOC(people.d_beg))
   oDoc.Tables(4).Rows(nRow).Cells(5).Select && Дата окончания лечения
   oWord.Selection.TypeText(DTOC(people.d_end))
   oDoc.Tables(4).Rows(nRow).Cells(6).Select && Код дефекта по 230
   oWord.Selection.TypeText(ALLTRIM(sookod.osn230))
   oDoc.Tables(4).Rows(nRow).Cells(7).Select && Что за дефект по-русски :-)
   oWord.Selection.TypeText(ALLTRIM(sookod.comment))
   oDoc.Tables(4).Rows(nRow).Cells(8).Select && Скока денег :-)
   oWord.Selection.TypeText(TRANSFORM(s_all, '99 999 999.99'))
    
   m.sum_def = m.sum_def + s_all

   nRow = nRow + 1
   nError = nError+1
  ENDIF 
 ENDSCAN 

 oDoc.BookMarks('sum_def').Select && Итого по акту
 oWord.Selection.TypeText(TRANSFORM(m.sum_def, '99 999 999.99'))

 SET RELATION OFF INTO people 
 USE IN people 
 
 SELECT prt2
 SCAN
  oDoc.Tables(4).Rows(nRow+3).Cells(1).Select && Код ошибки
  oWord.Selection.InsertRows
  oWord.Selection.TypeText(ALLTRIM(osn230))
  oDoc.Tables(4).Rows(nRow+3).Cells(2).Select && Сумма по ошибке
  oWord.Selection.TypeText(TRANSFORM(ssum,'99 999 999.99'))

 ENDSCAN 
 USE 
 
 CREATE CURSOR PrSvod (profil c(100), k_totl n(6), s_totl n(11,2),;
  k_good n(6), s_good n(11,2), k_bad n(6), s_bad n(11,2) )
  
 SELECT PrSvod
 INDEX ON profil TAG profil 
 SET ORDER TO profil 
 
 m.k_totl = 0
 m.s_totl = 0
 m.k_good = 0
 m.s_good = 0
 m.k_bad  = 0
 m.s_bad  = 0
 
 SELECT talon
 SCAN 
  m.d_type = d_type
   m.profil = profil 
   m.k_u    = k_u
   m.s_all  = s_all

   IF EMPTY(sError.rid)
    m.k_good = k_u
    m.s_good = s_all
    m.k_bad  = 0
    m.s_bad  = 0
   ELSE 
    m.k_good = 0
    m.s_good = 0
    m.k_bad  = k_u
    m.s_bad  = s_all
   ENDIF 

   IF SEEK(m.profil, 'PrSvod')
    UPDATE PrSvod SET k_totl = k_totl+m.k_u, s_totl = s_totl + m.s_all,;
     k_good = k_good+m.k_good, s_good = s_good + m.s_good, ;
     k_bad = k_bad+m.k_bad, s_bad = s_bad + m.s_bad ;
     WHERE profil=m.profil
   ELSE 
    INSERT INTO PrSvod (profil, k_totl, s_totl, k_good, s_good, k_bad, s_bad) ;
     VALUES ;
     (m.profil, m.k_u, m.s_all, m.k_good, m.s_good, m.k_bad, m.s_bad)
   ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO sError
 USE 
 SELECT sError
 SET RELATION OFF INTO sookod
 USE IN sookod
 USE
 
 SELECT PrSvod
 SET RELATION TO LEFT(profil,3) INTO prv002

 nRow = 3
 m.k_totl2 = 0
 m.s_totl2 = 0
 m.k_good2 = 0
 m.s_good2 = 0
 m.k_bad2  = 0
 m.s_bad2  = 0
 SCAN 
  oDoc.Tables(5).Cell(nRow,1).Select 
  oWord.Selection.TypeText(ALLTRIM(prv002.pr_name))
  oWord.Selection.InsertRowsBelow
  oDoc.Tables(5).Cell(nRow,2).Select 
  oWord.Selection.TypeText(TRANSFORM(k_totl,'999 999'))
  oDoc.Tables(5).Cell(nRow,3).Select 
  oWord.Selection.TypeText(TRANSFORM(s_totl,'99 999 999.99'))
  oDoc.Tables(5).Cell(nRow,4).Select 
  oWord.Selection.TypeText(TRANSFORM(k_bad,'999 999'))
  oDoc.Tables(5).Cell(nRow,5).Select 
  oWord.Selection.TypeText(TRANSFORM(s_bad,'99 999 999.99'))
  oDoc.Tables(5).Cell(nRow,6).Select 
  oWord.Selection.TypeText(TRANSFORM(k_good,'999 999'))
  oDoc.Tables(5).Cell(nRow,7).Select 
  oWord.Selection.TypeText(TRANSFORM(s_good,'99 999 999.99'))

  nRow = nRow + 1

  m.k_totl2 = m.k_totl2 + k_totl
  m.s_totl2 = m.s_totl2 + s_totl
  m.k_good2 = m.k_good2 + k_good
  m.s_good2 = m.s_good2 + s_good
  m.k_bad2  = m.k_bad2  + k_bad 
  m.s_bad2  = m.s_bad2  + s_bad
 ENDSCAN 

 oDoc.Tables(5).Cell(nRow,1).Select 
 oWord.Selection.TypeText('Итого')
 oDoc.Tables(5).Cell(nRow,2).Select 
 oWord.Selection.TypeText(TRANSFORM(m.k_totl2,'999 999'))
 oDoc.Tables(5).Cell(nRow,3).Select 
 oWord.Selection.TypeText(TRANSFORM(m.s_totl2,'99 999 999.99'))
 oDoc.Tables(5).Cell(nRow,4).Select 
 oWord.Selection.TypeText(TRANSFORM(m.k_bad2,'999 999'))
 oDoc.Tables(5).Cell(nRow,5).Select 
 oWord.Selection.TypeText(TRANSFORM(m.s_bad2,'99 999 999.99'))
 oDoc.Tables(5).Cell(nRow,6).Select 
 oWord.Selection.TypeText(TRANSFORM(m.k_good2,'999 999'))
 oDoc.Tables(5).Cell(nRow,7).Select 
 oWord.Selection.TypeText(TRANSFORM(m.s_good2,'99 999 999.99'))
 
 oDoc.BookMarks('say_summa').Select
 oWord.Selection.TypeText(cpr(INT(m.s_good2))+' '+;
  PADL(INT(((m.s_good2)-INT(m.s_good2))*100),2,'0')+' КОПЕЕК')
 
 SET RELATION OFF INTO prv002
 USE IN prv002
 USE 
 
 SELECT AisOms

 oDoc.Bookmarks('kol_usl_predst').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.usl_all)))

 oDoc.Bookmarks('sum_predst').Select  
 oWord.Selection.TypeText(TRANSFORM(m.pr_sum, '99999999.99'))

 oDoc.Bookmarks('kol_usl_bad').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.usl_all_bad)))

 oDoc.Bookmarks('sum_bad').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sum_all_bad,'99999999.99'))

 oDoc.Bookmarks('sum_iskl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sum_all_bad,'99999999.99'))

 oDoc.Bookmarks('sum_ok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sum_all_ok,'99999999.99'))

 oDoc.SaveAs(DocName, 0)
 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
 ENDTRY 
 
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  oDoc.Close
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 
 
RETURN  

