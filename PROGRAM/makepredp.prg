FUNCTION MakePredp(lcPath, IsVisible, IsQuit)
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÑÔÎÐÌÈÐÎÂÀÒÜ ÏÐÅÄÏÈÑÀÍÈÅ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 DotName = '\Ïðåäïèñàíèå'
 IF !fso.FileExists(ptempl+'\'+DotName+'.dot')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ÏÐÅÄÏÈÑÀÍÈß!'+CHR(13)+CHR(10),0+16,;
   'Ïðåäïèñàíèå.dot')
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
* IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
*  IF USED('sprlpu')
*   USE IN sprlpu
*  ENDIF 
*  IF USED('mfile')
*   USE IN mfile
*  ENDIF 
*  RETURN 
* ENDIF 
 
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
*  USE IN sprlpu
  MESSAGEBOX(CHR(13)+CHR(10)+'Ê ÂÛÁÐÀÍÍÎÌÓ ËÏÓ ØÒÐÀÔÍÛÕ ÑÀÍÊÖÈÉ ÍÅ ÏÐÈÌÅÍßËÎÑÜ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 WAIT "ÇÀÏÓÑÊ WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR
 
 m.mmy = PADL(m.tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\prp'+LOWER(m.qcod)+m.mmy
 
 IF fso.FileExists(DocName+'.doc')
  fso.DeleteFile(DocName+'.doc')
 ENDIF 
 IF fso.FileExists(DocName+'.pdf')
  fso.DeleteFile(DocName+'.pdf')
 ENDIF 
 
 oDoc = oWord.Documents.Add(pTempl+'\'+DotName)

 oDoc.Bookmarks('lpuname').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('qname').Select  
 oWord.Selection.TypeText(m.qname+', êîä ÑÌÎ '+m.qcod)
 oDoc.Bookmarks('lpuname2').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('d_beg').Select  
 oWord.Selection.TypeText(DTOC(m.tdat1)+' ã.')
 oDoc.Bookmarks('d_end').Select  
 oWord.Selection.TypeText(DTOC(m.tdat2)+' ã.')

 SELECT mfile 
 nRow = 2

 m.s_straf = 0 
 m.ss_2    = 0

 SCAN 
  IF straf<=0
   LOOP 
  ENDIF 
  
  SCATTER MEMVAR 

  m.ss_2    = m.ss_2 + m.s_2

  oDoc.Tables(3).Rows(nRow).Cells(2).Select
  oWord.Selection.TypeText(m.osn230)
  oDoc.Tables(3).Rows(nRow).Cells(3).Select
  oWord.Selection.TypeText(TRANSFORM(m.straf,'9.99'))
  oDoc.Tables(3).Rows(nRow).Cells(4).Select
  oWord.Selection.TypeText(TRANSFORM(m.s_2,'99999.99'))

  nRow = nRow + 1

  oDoc.Tables(3).Rows(nRow).Select 
  oWord.Selection.InsertRows

 ENDSCAN 

 oDoc.Tables(3).Rows(nRow).Cells(1).Select
 oWord.Selection.TypeText('Èòîãî:')
 oDoc.Tables(3).Rows(nRow).Cells(4).Select
 oWord.Selection.TypeText(TRANSFORM(m.ss_2,'99999.99'))
 
 m.sumstraf = cpr(FLOOR(m.ss_2))
 m.kopstraf = m.ss_2 - FLOOR(m.ss_2)

 oDoc.Bookmarks('ndog').Select  
 oWord.Selection.TypeText(m.lpudog)
 oDoc.Bookmarks('qname2').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('qname3').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('qname4').Select  
 oWord.Selection.TypeText(m.qname)

 oDoc.Bookmarks('sumstraf').Select  
 oWord.Selection.TypeText(m.sumstraf)
 oDoc.Bookmarks('kopstraf').Select  
 oWord.Selection.TypeText(PADL(m.kopstraf,2,'0'))

 oDoc.SaveAs(DocName, 0)
* TRY 
*  oDoc.SaveAs(DocName, 17)
* CATCH 
* ENDTRY 

 IF fso.FileExists(DocName+'.pdf')
  IF fso.FileExists(DocName+'.doc')
   fso.DeleteFile(DocName+'.doc')
  ENDIF 
 ENDIF 

 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

 USE IN mfile
* USE IN sprlpu

RETURN 
