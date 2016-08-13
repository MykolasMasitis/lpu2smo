FUNCTION oms6cword(lcPath, IsVisible, IsQuit)
 
 USE pcommon+'\smo' ALIAS smo IN 0 SHARED ORDER code 
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx' ALIAS sprcokr IN 0 SHARED ORDER cokr
 IF !USED('sprlpu')
  m.WasUsedSprLpu = .f.
  =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "lpu_id")
 ELSE 
  m.WasUsedSprLpu = .t.
 ENDIF 

 SELECT AisOms
 
 m.mmy        = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.mcod       = mcod
 m.lpuid      = lpuid
 m.lpuname    = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.cokr     = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
 m.cokr_name = IIF(SEEK(m.cokr, 'sprcokr'), ALLTRIM(sprcokr.name_okr), '')
 m.smoname    = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.fullname), '')
 m.smonames   = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.name), '')
 m.arcfname   = 'b'+m.mcod+'.'+m.mmy
 m.message_id = ALLTRIM(cmessage)
 m.datpriemki = TTOC(Recieved)
 
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa   = s_pred

 poi_file   = fso.GetFile(lcPath + '\' + arcfname)
 m.arcfdate = poi_file.DateLastModified
 
 ZipItemCount = 5

 DotName = pTempl + "\Prqqmmy.dot"
 DocName = lcPath + "\Pr" + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 
 eeFile = 'e'+m.mcod
 USE &lcPath\Talon  IN 0 ALIAS Talon SHARED 
 USE &lcPath\People IN 0 ALIAS people SHARED 
 USE &lcPath\&eeFile IN 0 ALIAS sError SHARED ORDER rid 
 USE &lcPath\&eeFile IN 0 ALIAS rError SHARED ORDER rrid AGAIN 

 SELECT people
 SET RELATION TO RecId INTO rError
 m.pazPrd = RECCOUNT('people')
 COUNT FOR EMPTY(rError.rid) TO m.PazPrin
 SET RELATION OFF INTO rError
 USE
 USE IN rError

 SELECT Talon 
 SET RELATION TO RecId INTO sError

 m.summa   = 0
 m.SchPrin = 0
 m.SchIskl = 0
 m.SumPrin = 0
 m.SumIskl = 0
 
 SCAN 
* IF INLIST(d_type, 'z', 'h')
*  m.SumPrin = m.SumPrin - s_all
* ELSE 
  m.summa = m.summa + s_all
  m.SchPrin = m.SchPrin + IIF(EMPTY(sError.rid), 1, 0)
  m.SumPrin = m.SumPrin + IIF(EMPTY(sError.rid), s_all, 0)
  m.SchIskl = m.SchIskl + IIF(!EMPTY(sError.rid), 1, 0)
  m.SumIskl = m.SumIskl + IIF(!EMPTY(sError.rid), s_all, 0)
* ENDIF
 ENDSCAN  

 m.kol_sch = RECCOUNT('Talon')

 SET RELATION OFF INTO sError
 USE 
 USE IN sError
 
 USE IN smo 
 USE IN sprcokr
 IF m.WasUsedSprLpu = .f.
  USE IN SprLpu
 ENDIF 

 SELECT AisOms

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)
 
 oDoc.Bookmarks('DatPriemki').Select  
 oWord.Selection.TypeText(m.datpriemki)
 oDoc.Bookmarks('SmoName').Select  
 oWord.Selection.TypeText(m.smoname)
 oDoc.Bookmarks('LpuName').Select  
 oWord.Selection.TypeText(m.lpuname+', '+m.cokr_name+', '+m.mcod)
 oDoc.Bookmarks('Period').Select  
 oWord.Selection.TypeText(m.period)

 oDoc.Bookmarks('PazPrd').Select  
 oWord.Selection.TypeText(TRANSFORM(m.kol_paz,'999999'))
 oDoc.Bookmarks('SchPrd').Select  
 oWord.Selection.TypeText(TRANSFORM(m.kol_sch,'999999'))
 oDoc.Bookmarks('SumPrd').Select  
 oWord.Selection.TypeText(TRANSFORM(m.summa,'99999999.99'))

 oDoc.Bookmarks('PazPrin').Select  
 oWord.Selection.TypeText(TRANSFORM(m.PazPrin,'999999'))
 oDoc.Bookmarks('SchPrin').Select  
 oWord.Selection.TypeText(TRANSFORM(m.SchPrin,'999999'))
 oDoc.Bookmarks('SumPrin').Select  
 oWord.Selection.TypeText(TRANSFORM(m.SumPrin,'99999999.99'))

 oDoc.Bookmarks('SchIskl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.SchIskl,'999999'))
 oDoc.Bookmarks('SumIskl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.SumIskl,'99999999.99'))

 oDoc.Bookmarks('arcfname').Select  
 oWord.Selection.TypeText(m.arcfname)
 oDoc.Bookmarks('arcfdate').Select  
 oWord.Selection.TypeText(DTOC(m.arcfdate))
 oDoc.Bookmarks('KolVloz').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.ZipItemCount)))
 oDoc.Bookmarks('sumz').Select  
 oWord.Selection.TypeText(TRANSFORM(0,'99 999 999.99'))
 oDoc.Bookmarks('sumh').Select  
 oWord.Selection.TypeText(TRANSFORM(0,'99 999 999.99'))

 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
*  MESSAGEBOX(CHR(13)+CHR(10)+;
   '—Œ’–¿Õ≈Õ»≈ ¬ PDF-‘Œ–Ã¿“≈'+CHR(13)+CHR(10)+;
   'Õ≈ œŒƒƒ≈–∆»¬¿≈“—ﬂ ”—“¿ÕŒ¬À≈ÕÕŒ…'+CHR(13)+CHR(10)+;
   '¬≈–—»≈… WORD!',0+64,'')
 ENDTRY 

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

