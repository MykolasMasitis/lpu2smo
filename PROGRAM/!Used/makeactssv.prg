PROCEDURE MakeActsSV
 IF MESSAGEBOX('¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹'+CHR(13)+CHR(10)+;
  '¿ “€ —¬≈– » — Àœ”'+CHR(13)+CHR(10)+;
  '«¿ '+NameOfMonth(tMonth)+' Ã≈—ﬂ÷ '+STR(tYear,4)+' √Œƒ¿?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 DotName = '¿ÍÚ Ò‚ÂÍË ‡Ò˜ÂÚÓ‚.xlt'

 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ Œ“◊≈“¿'+CHR(13)+CHR(10)+;
   '¿ÍÚ Ò‚ÂÍË ‡Ò˜ÂÚÓ‚.xlt',0+32,'')
 ENDIF 
 
 ActsDir = LEFT(pout, RAT('\',pout))+'ACTS'
 IF !fso.FolderExists(ActsDir)
  fso.CreateFolder(ActsDir)
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\SprLpuxx', "SprLpu", "shar", "mcod") > 0 
  USE IN aisoms
  RETURN 
 ENDIF 
* IF !fso.FileExists(pOut+'\'+gcPeriod+'\mdr'+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)+'.dbf')
*  USE IN sprlpu
*  USE IN aisoms
*  MESSAGEBOX('Õ≈ –¿—◊»“¿Õ¿ ÃŒƒ≈–Õ»«¿÷»ﬂ!'+CHR(13)+CHR(10)+;
   'Œ“—”“—“¬”≈“ ‘¿…À '+gcPeriod+'\mdr'+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)+'.dbf',;
   0+16,'')
*  RETURN 
* ENDIF 
* IF OpenFile(pOut+'\'+gcPeriod+'\mdr'+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2), "mdrn", "excl") > 0 
*  USE IN sprlpu
*  USE IN aisoms
*  RETURN 
* ENDIF 

* SELECT mdrn
* INDEX ON mcod TAG mcod 
* SET ORDER TO mcod

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 WAIT "«¿œ”—  EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 

 WITH oExcel
  .SheetsInNewWorkbook = 1
  .Interactive= .F. 
  .DisplayAlerts = .F.
  .ReferenceStyle= -4150  && xlR1C1
 ENDWITH 

 WAIT CLEAR 
 
 m.NameOfPeriod = NameOfMonth(tMonth)+' '+STR(tYear)
 ClmnName = 'Sum'+PADL(tMonth,2,'0')
 
 SELECT aisoms
* SET RELATION TO mcod INTO mdrn
 SCAN 
  m.mcod  = mcod 
  m.lpuid = lpuid
  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+' ,'+m.mcod, '')
  WAIT 'Œ¡–¿¡Œ“ ¿ '+m.mcod WINDOW NOWAIT 
  DocName = 'ActSS_'+m.mcod
  IF !fso.FileExists(ActsDir + '\' + DocName + '.xls')
   fso.CopyFile(pTempl+'\¿ÍÚ Ò‚ÂÍË ‡Ò˜ÂÚÓ‚.xlt', ActsDir+'\ActSS_'+m.mcod+'.xls')
  ENDIF

*  sum_modern = mdrn.&ClmnName
  
  oBook = oExcel.WorkBooks.Open(ActsDir+'\'+DocName)
  nSheets = oBook.Sheets.Count - 1
  IF nSheets >= tMonth

   IF nSheets = tMonth
    oSheet = oBook.Sheets(tMonth-1)
    oSheet.Cells.Select
    oExcel.Selection.Copy
    oSheet = oBook.Sheets(tMonth)
    oSheet.Name='¿ÔÂÎ¸'
    oSheet.Paste
    oBook.Sheets.Add(,oBook.Sheets.Item(tMonth))
   ENDIF 

   m.sdolgtotlpu = IIF(tMonth>1, oBook.Sheets(tMonth-1).Cells(34,3).Value, 0)
   m.sdolgtotsmo = IIF(tMonth>1, oBook.Sheets(tMonth-1).Cells(34,4).Value, 0)
   m.sdolgsmo2lpulpu = IIF(tMonth>1, oBook.Sheets(tMonth-1).Cells(35,3).Value, 0)
   m.sdolgsmo2lpusmo = IIF(tMonth>1, oBook.Sheets(tMonth-1).Cells(35,4).Value, 0)
   m.sdolglpu2smolpu = IIF(tMonth>1, oBook.Sheets(tMonth-1).Cells(36,3).Value, 0)
   m.sdolglpu2smosmo = IIF(tMonth>1, oBook.Sheets(tMonth-1).Cells(36,4).Value, 0)

   oSheet=oBook.Sheets(tMonth)
   oSheet.Select

   WITH oSheet
    .Cells(08,03) = m.NameOfPeriod
    .Cells(09,03) = m.lpuname
    .Cells(10,03) = m.lpuid
    
    .Cells(15,3) = m.sdolgtotlpu
    .Cells(15,4) = m.sdolgtotsmo
    .Cells(16,3) = m.sdolgsmo2lpulpu
    .Cells(16,4) = m.sdolgsmo2lpusmo
    .Cells(17,3) = m.sdolglpu2smolpu
    .Cells(17,4) = m.sdolglpu2smosmo
    
    .Cells(18,3) = s_pred
    .Cells(18,4) = s_pred
    .Cells(19,3) = s_pred
    .Cells(19,4) = s_pred
    
    .Cells(21,3) = sum_flk + sum_mee
    .Cells(21,4) = sum_flk + sum_mee 
    .Cells(22,3) = sum_flk 
    .Cells(22,4) = sum_flk 
    .Cells(23,3) = sum_mee
    .Cells(23,4) = sum_mee 
    
    .Cells(28,3) = s_pred - sum_flk - sum_mee 
    .Cells(28,4) = s_pred - sum_flk - sum_mee
    .Cells(29,3) = s_pred - sum_flk - sum_mee - sum_modern
    .Cells(29,4) = s_pred - sum_flk - sum_mee - sum_modern
    .Cells(30,3) = sum_modern
    .Cells(30,4) = sum_modern

   ENDWITH 
   
  ENDIF 
  
  oBook.Save
  oBook.Close

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 
  
 ENDSCAN 
* SET RELATION OFF INTO mdrn
 WAIT CLEAR 

 WAIT "Œ—“¿ÕŒ¬ ¿ EXCEL..." WINDOW NOWAIT 
 oExcel.Quit
 WAIT CLEAR 

 USE IN aisoms
 USE IN sprlpu
* SELECT mdrn 
* SET ORDER TO
* DELETE TAG ALL 
* USE

 SET ESCAPE &OldEscStatus

RETURN 