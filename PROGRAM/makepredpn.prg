FUNCTION MakePredpN(lcPath, IsVisible, IsQuit)
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÑÔÎÐÌÈÐÎÂÀÒÜ ÏÐÅÄÏÈÑÀÍÈÅ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 DotName = '\Ïðåäïèñàíèå'
 IF !fso.FileExists(ptempl+'\'+DotName+'.xlt')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ÏÐÅÄÏÈÑÀÍÈß!'+CHR(13)+CHR(10),0+16,;
   'Ïðåäïèñàíèå.xlt')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(lcPath)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ËÏÓ!'+CHR(13)+CHR(10)+;
   lcPath+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.mcod    = RIGHT(ALLTRIM(lcpath),7)
 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
 m.lpuname = IIF(SEEK(m.mcod, 'sprlpu', 'mcod'), ALLTRIM(sprlpu.fullname), '')
 m.lpudog  = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')

 mfile = 'm'+ m.mcod
 IF !fso.FileExists(lcPath+'\'+mfile+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ M-ÔÀÉË!'+CHR(13)+CHR(10),0+16, mfile+'.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(lcPath+'\'+mfile, 'mfile', 'shar')>0
  IF USED('mfile')
   USE IN mfile
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT mfile 
 m.nRecsStrafs = 0
 SCAN 
  IF straf<=0
   LOOP 
  ENDIF 
 
  m.nRecsStrafs = m.nRecsStrafs + 1
  
 ENDSCAN 
 
 IF m.nRecsStrafs = 0
  USE IN mfile
  MESSAGEBOX(CHR(13)+CHR(10)+'Ê ÂÛÁÐÀÍÍÎÌÓ ËÏÓ ØÒÐÀÔÍÛÕ ÑÀÍÊÖÈÉ ÍÅ ÏÐÈÌÅÍßËÎÑÜ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 WAIT "ÇÀÏÓÑÊ EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR
 
 m.mmy = PADL(m.tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\prp'+LOWER(m.qcod)+m.mmy
 
 IF fso.FileExists(DocName+'.xls')
  fso.DeleteFile(DocName+'.xls')
 ENDIF 
 IF fso.FileExists(DocName+'.pdf')
  fso.DeleteFile(DocName+'.pdf')
 ENDIF 
 
 oDoc = oExcel.WorkBooks.Add(pTempl+'\'+DotName)

 WITH oExcel
  .Cells(10,1).Value = m.qname
  .Cells(16,2).Value = m.lpuname
  .Cells(18,3).Value = DTOC(m.tdat1)+' ã.'
  .Cells(18,6).Value = DTOC(m.tdat2)+' ã.'
 ENDWITH 

 SELECT mfile 
 nRow = 24
 nNn  = 1

 m.s_straf = 0 
 m.ss_2    = 0

 SCAN 
  IF straf<=0
   LOOP 
  ENDIF 
  
  SCATTER MEMVAR 

  m.ss_2    = m.ss_2 + m.s_2
  
  WITH oExcel
   .Cells(nRow,1).Value = m.nNn
   .Cells(nRow,7).Value = m.osn230
   .Cells(nRow,9).Value = TRANSFORM(m.straf,'9.99')
   .Cells(nRow,12).Value = TRANSFORM(m.s_2,'99999.99')
  ENDWITH 

  nRow = nRow + 1
  nNn  = nNn + 1
  oExcel.Rows(nRow).Insert

  oRange = oExcel.Range(oExcel.Cells(nRow,2), oExcel.Cells(nRow,4))
  oRange.Merge
  oRange = oExcel.Range(oExcel.Cells(nRow,5), oExcel.Cells(nRow,6))
  oRange.Merge
  oRange = oExcel.Range(oExcel.Cells(nRow,7), oExcel.Cells(nRow,8))
  oRange.Merge
  oRange = oExcel.Range(oExcel.Cells(nRow,9), oExcel.Cells(nRow,11))
  oRange.Merge
  oRange = oExcel.Range(oExcel.Cells(nRow,12), oExcel.Cells(nRow,14))
  oRange.Merge

 ENDSCAN 

 WITH oExcel
  .Cells(nRow+2,12).Value = TRANSFORM(m.ss_2,'999999.99')
  .Cells(nRow+5,1).Value = 'Â ñîîòâåòñòâèè ñ Äîãîâîðîì íà îêàçàíèå è îïëàòó ìåäèöèíñêîé ïîìîùè ïî îáÿçàòåëüíîìó ìåäèöèíñêîìó ñòðàõîâàíèþ ¹ '+;
    m.lpudog + ' îò ____________20___ ã.'
  .Cells(nRow+7,1).Value = '1. Ïåðå÷èñëèòü øòðàô â ðàçìåðå '+TRANSFORM(FLOOR(m.ss_2),'999999')+' ðóá., '+;
    TRANSFORM((m.ss_2-FLOOR(m.ss_2))*100,'99')+' êîï. ïî ñëåäóþùèì ðåêâèçèòàì:'
 ENDWITH 

* oDoc.Tables(3).Rows(nRow).Cells(1).Select
* oWord.Selection.TypeText('Èòîãî:')
* oDoc.Tables(3).Rows(nRow).Cells(4).Select
* oWord.Selection.TypeText(TRANSFORM(m.ss_2,'99999.99'))
 
 m.sumstraf = cpr(FLOOR(m.ss_2))
 m.kopstraf = m.ss_2 - FLOOR(m.ss_2)

* oDoc.Bookmarks('ndog').Select  
* oWord.Selection.TypeText(m.lpudog)
* oDoc.Bookmarks('qname2').Select  
* oWord.Selection.TypeText(m.qname)
* oDoc.Bookmarks('qname3').Select  
* oWord.Selection.TypeText(m.qname)
* oDoc.Bookmarks('qname4').Select  
* oWord.Selection.TypeText(m.qname)

* oDoc.Bookmarks('sumstraf').Select  
* oWord.Selection.TypeText(m.sumstraf)
* oDoc.Bookmarks('kopstraf').Select  
* oWord.Selection.TypeText(PADL(m.kopstraf,2,'0'))

 oDoc.SaveAs(DocName)

 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 

 USE IN mfile
RETURN 
