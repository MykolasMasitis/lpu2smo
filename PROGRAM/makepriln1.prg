PROCEDURE MakePrilN1
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹'+CHR(13)+CHR(10)+;
  'œ–»ÀŒ∆≈Õ»ﬂ π1   ƒŒ√Œ¬Œ–”?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\PrilN1.dot')
  MESSAGEBOX('Œ—“”“—¬”≈“ ÿ¿¡ÀŒÕ ƒŒ ”Ã≈Õ“¿!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\volumes', 'volumes', 'shar', 'mcod')>0
  USE IN sprlpu
  RETURN 
 ENDIF 
 
 PrilDir = fso.GetParentFolderName(pbin)+'\PRILN1'
 IF !fso.FolderExists(PrilDir)
  fso.CreateFolder(PrilDir)
 ENDIF 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 SELECT sprlpu 
 
 SCAN 
  m.mcod = mcod
  m.lpuname = ALLTRIM(superfull)
  WAIT m.mcod WINDOW NOWAIT 
  
  m.sumyear = 0
  m.sum01 = 0
  m.sum02 = 0
  m.sum03 = 0
  m.sum04 = 0
  IF SEEK(m.mcod, 'volumes')
   m.sumyear = volumes.sumyear
   m.sum01 = volumes.sum01
   m.sum02 = volumes.sum02
   m.sum03 = volumes.sum03
   m.sum04 = volumes.sum04
  ENDIF 

  DocName = PrilDir+'\'+m.mcod+'_pril1'
  oDoc = oWord.Documents.Add(pTempl+'\PrilN1')

  IF fso.FileExists(DocName+'.doc')
   LOOP 
  ENDIF 

  oDoc.Bookmarks('mcod').Select  
  oWord.Selection.TypeText(m.mcod+'.23')
  oDoc.Bookmarks('lpuname').Select  
  oWord.Selection.TypeText(m.lpuname)
  oDoc.Bookmarks('sumyear').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sumyear,'99999999.99'))
  oDoc.Bookmarks('sum01').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sum01,'99999999.99'))
  oDoc.Bookmarks('sum02').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sum02,'99999999.99'))
  oDoc.Bookmarks('sum03').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sum03,'99999999.99'))
  oDoc.Bookmarks('sum04').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sum04,'99999999.99'))

  oDoc.SaveAs(DocName,0)
  oDoc.Close

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 
  WAIT CLEAR 
 
 ENDSCAN 
 USE 
 USE IN volumes

 SET ESCAPE &OldEscStatus

 WAIT "Œ—“¿ÕŒ¬ WORD..." WINDOW NOWAIT  
 oWord.Quit
 WAIT CLEAR 
 
  
RETURN 