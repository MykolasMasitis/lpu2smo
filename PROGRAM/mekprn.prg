FUNCTION MekPrn(lcPath, IsVisible, IsQuit)

 m.mcod  = mcod
 m.lpuid = lpuid
 m.lpuname  = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.lpuadr   = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 
 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 m.mmy = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa = s_pred

 DotName = pTempl + "\¿ÍÚ_Ã› .dot"
 DocName = lcPath + "\Akt" + STR(m.lpuid,4) + m.qcod + m.mmy

 eeFile = 'e'+m.mcod

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)

 m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 oDoc.Bookmarks('n_akt').Select  
 oWord.Selection.TypeText(m.n_akt)
 
 m.d_akt = DTOC(DATE())
 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)
 
 m.akt_month = NameOfMonth(tMonth)
 oDoc.Bookmarks('akt_month').Select  
 oWord.Selection.TypeText(m.akt_month)

 m.akt_year = STR(tYear,4)
 oDoc.Bookmarks('akt_year').Select  
 oWord.Selection.TypeText(m.akt_year)
 
 m.nr_akt = m.mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 oDoc.Bookmarks('nr_akt').Select  
 oWord.Selection.TypeText(m.nr_akt)
 
 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('mcod').Select  
 oWord.Selection.TypeText(m.mcod)

 m.smo_name = m.qname
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.smo_name)
 oDoc.Bookmarks('qq').Select  
 oWord.Selection.TypeText(m.qcod)

 m.dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.dat2 = DTOC(GOMONTH(CTOD(m.dat1),1)-1)

 m.usl_amb   = 0
 m.sum_amb   = 0

 m.usl_gosp  = 0 
 m.sum_gosp  = 0

 m.usl_dstac = 0
 m.sum_dstac = 0

 m.pr_sum = 0
 m.sum_zh = 0

 m.usl_all_ok  = 0
 m.usl_all_bad = 0
 m.sum_all_ok  = 0
 m.sum_all_bad = 0

 m.usl_amb_ok  = 0
 m.usl_amb_bad = 0
 m.sum_amb_ok  = 0
 m.sum_amb_bad = 0

 m.usl_gosp_ok  = 0
 m.usl_gosp_bad = 0
 m.sum_gosp_ok  = 0
 m.sum_gosp_bad = 0

 m.usl_dstac_ok  = 0
 m.usl_dstac_bad = 0
 m.sum_dstac_ok  = 0
 m.sum_dstac_bad = 0

 USE &lcPath\Talon IN 0 ALIAS Talon SHARED 
 USE &lcPath\&eeFile IN 0 ALIAS sError SHARED ORDER rid 

 nRowGosp  = 0 
 nRowDStac = 0
 nRowAmb   = 0

 SELECT Talon 
 SET RELATION TO RecId INTO sError

 m.usl_all = 0
 
 SCAN 
  m.cod = cod
  m.d_type = d_type

  m.usl_all = m.usl_all + 1

  m.pr_sum = m.pr_sum + IIF(!INLIST(m.d_type,'z','h'), s_all , 0)
  m.sum_zh = m.sum_zh + IIF(INLIST(m.d_type,'z','h'), s_all , 0)


  IF EMPTY(sError.rid)
   m.usl_all_ok = m.usl_all_ok + 1
   m.sum_all_ok = m.sum_all_ok + IIF(!INLIST(m.d_type,'z','h'), s_all , 0)
  ELSE 
   m.usl_all_bad = m.usl_all_bad + 1
   m.sum_all_bad = m.sum_all_bad + s_all
  ENDIF 
 ENDSCAN 

 SET RELATION OFF INTO sError
 USE 
 USE IN sError
 
 m.sum_all_ok = m.sum_all_ok - m.sum_zh
 
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

