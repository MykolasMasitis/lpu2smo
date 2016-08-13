FUNCTION oms6u(lcPath)

 otalk = SET("Talk")
 SET TALK OFF 

 USE pcommon+'\smo'      ALIAS smo     IN 0 SHARED ORDER code 
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx' ALIAS sprcokr IN 0 SHARED ORDER cokr
 
 m.mmy        = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.mcod       = mcod
 m.lpuid      = lpuid
 m.lpuname    = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.cokr       = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
 m.cokr_name  = IIF(SEEK(m.cokr, 'sprcokr'), ALLTRIM(sprcokr.name_okr), '')
 m.smoname    = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.fullname), '')
 m.smonames   = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.name), '')
 m.arcfname   = 'b'+m.mcod+'.'+m.mmy
 m.message_id = ALLTRIM(cmessage)
 
 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 
 m.n_schet = STR(tYear,4)+PADL(tMonth,2,'0')
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa = s_pred

 poi_file = fso.GetFile(lcPath + '\' + arcfname)
 m.arcfdate = poi_file.DateLastModified
 
 DotName = pTempl + "\S_qqmmy.dot"
 DocName = lcPath + "\S_"+UPPER(m.qcod)+m.mmy

 USE &lcPath\Talon IN 0 ALIAS Talon SHARED 
 SELECT Talon 
 m.kol_sch = RECCOUNT()
 SUM s_all FOR d_type='z' TO sum_z
 SUM s_all FOR d_type='h' TO sum_h
 
 USE IN smo 
 USE IN talon
 USE IN sprcokr
 
 SELECT AisOms

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)
 
 oDoc.Bookmarks('n_schet').Select  
 oWord.Selection.TypeText(m.n_schet)

 oDoc.Bookmarks('lpuname').Select  
 oWord.Selection.TypeText(m.lpuname+', '+m.cokr_name+', '+m.mcod)

 oDoc.Bookmarks('smoname').Select  
 oWord.Selection.TypeText(m.smoname)

 oDoc.Bookmarks('periodname').Select  
 oWord.Selection.TypeText(m.period)

 oDoc.Bookmarks('smonames').Select  
 oWord.Selection.TypeText(m.smonames)

 oDoc.Bookmarks('kol_paz').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.kol_paz)))

 oDoc.Bookmarks('kol_sch').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.kol_sch)))
 
 oDoc.Bookmarks('summa').Select  
 oWord.Selection.TypeText(TRANSFORM(m.summa,'99 999 999.99'))

 oDoc.Bookmarks('sum_z').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sum_z,'99 999 999.99'))

 oDoc.Bookmarks('sum_h').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sum_h,'99 999 999.99'))

 oDoc.Bookmarks('arcfname').Select  
 oWord.Selection.TypeText(m.arcfname)

 oDoc.Bookmarks('arcfdate').Select  
 oWord.Selection.TypeText(DTOC(m.arcfdate))

 oDoc.Bookmarks('message_id').Select  
 oWord.Selection.TypeText(m.message_id)

 oDoc.SaveAs(DocName, 0)
 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
 ENDTRY 
 
 oWord.Visible = .t.
 
 SET TALK &otalk
 
RETURN  

