FUNCTION AktMEEss(tmcod, IsVisible, IsQuit)

 lcPath = pbase+'\'+gcPeriod+'\'+tmcod
 DotName = pTempl + "\Акт_МЭЭ_СС_план.dot"
 DocName = lcPath + "\AktMeeSSPl_" + ALLTRIM(sn_pol)

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)
 
* oDoc.Bookmarks('DatPriemki').Select  
* oWord.Selection.TypeText(m.datpriemki)

 oDoc.SaveAs(DocName, 0)
 
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

RETURN 