PROCEDURE FormSh3Bis
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ œŒ œ–Œ¬≈ƒ≈ÕÕ€Ã › —œ≈–“»«¿Ã?'+CHR(13)+CHR(10),4+32,'‘Œ–Ã¿ ÿ-3¡ËÒ')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À DSPCODES.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 m.et = '2'
 DO FORM SlExpTip TO m.resp
 IF m.resp=.f.
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 m.docname = ''
 DO CASE 
  CASE m.et='2'
   m.docname = m.docname + 'œÎ‡ÌÓ‚‡ˇ Ã››'
  CASE m.et='3'
   m.docname = m.docname + '÷ÂÎÂ‚‡ˇ Ã››'
  CASE m.et='4'
   m.docname = m.docname + 'œÎ‡ÌÓ‚‡ˇ › Ãœ'
  CASE m.et='5'
   m.docname = m.docname + '÷ÂÎÂ‚‡ˇ › Ãœ'
  CASE m.et='6'
   m.docname = m.docname + '“ÂÏ‡ÚË˜ÂÒÍ‡ˇ › Ãœ'
  CASE m.et='7'
   m.docname = m.docname + '“ÂÏ‡ÚË˜ÂÒÍ‡ˇ Ã››'
 ENDCASE  

 m.docname = m.docname + ' Á‡ ' + NameOfMonth(m.tmonth)+' '+STR(tYear,4)
 
 PUBLIC oExcel AS Excel.Application
 WAIT "«‡ÔÛÒÍ MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 m.BookName = 'sh3bis_'+m.et
 m.nOpBooks = oExcel.Workbooks.Count 
 IF m.nOpBooks>0
  FOR m.nBook=1 TO m.nOpBooks
   m.cBookName = LOWER(ALLTRIM(oExcel.Workbooks.Item(m.nBook).Name))
   IF m.cBookName=m.BookName+'.xls'
    oExcel.Workbooks.Item(m.nBook).Close 
   ENDIF 
  NEXT 
 ENDIF 

 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 
 oSheet = oExcel.ActiveSheet 

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,14))
 oRange.Merge
 oExcel.Cells(1,1) = m.docname
 oExcel.Cells(1,1).HorizontalAlignment = -4108

 oRange = oExcel.Range(oExcel.Cells(2,3), oExcel.Cells(2,5))
 oRange.Merge
 oExcel.Cells(2,3) = 'œÓÙËÎ‡ÍÚË˜ÂÒÍËÂ ÏÂÓÔËˇÚËˇ'
 oExcel.Cells(2,3).HorizontalAlignment = -4108
 oRange = oExcel.Range(oExcel.Cells(2,6), oExcel.Cells(2,8))
 oRange.Merge
 oExcel.Cells(2,6) = '¿Ï·ÛÎ‡ÚÓÌÓ-ÔÓÎËÍÎËÌË˜ÂÒÍ‡ˇ ÔÓÏÓ˘¸'
 oExcel.Cells(2,6).HorizontalAlignment = -4108
 oRange = oExcel.Range(oExcel.Cells(2,9), oExcel.Cells(2,11))
 oRange.Merge
 oExcel.Cells(2,9) = '—Ú‡ˆËÓÌ‡Á‡ÏÂ˘‡˛˘‡ˇ ÔÓÏÓ˘¸'
 oExcel.Cells(2,9).HorizontalAlignment = -4108
 oRange = oExcel.Range(oExcel.Cells(2,12), oExcel.Cells(2,14))
 oRange.Merge
 oExcel.Cells(2,12) = '—Ú‡ˆËÓÌ‡Ì‡ˇ ÔÓÏÓ˘¸'
 oExcel.Cells(2,12).HorizontalAlignment = -4108

 oExcel.Cells(3,3)  = 'œÓ‚ÂÂÌÓ'
 oExcel.Cells(3,4)  = 'ƒÂÙÂÍÚÌ˚ı'
 oExcel.Cells(3,5)  = 'ƒÂÙÂÍÚÓ‚'
 oExcel.Cells(3,6)  = 'œÓ‚ÂÂÌÓ'
 oExcel.Cells(3,7)  = 'ƒÂÙÂÍÚÌ˚ı'
 oExcel.Cells(3,8)  = 'ƒÂÙÂÍÚÓ‚'
 oExcel.Cells(3,9)  = 'œÓ‚ÂÂÌÓ'
 oExcel.Cells(3,10) = 'ƒÂÙÂÍÚÌ˚ı'
 oExcel.Cells(3,11) = 'ƒÂÙÂÍÚÓ‚'
 oExcel.Cells(3,12) = 'œÓ‚ÂÂÌÓ'
 oExcel.Cells(3,13) = 'ƒÂÙÂÍÚÌ˚ı'
 oExcel.Cells(3,14) = 'ƒÂÙÂÍÚÓ‚'

 oExcel.Columns(1).NumberFormat  = '@'
 oExcel.Columns(1).ColumnWidth   = 7
 oExcel.Columns(1).NumberFormat  = '@'
 oExcel.Columns(2).ColumnWidth   = 25
 oExcel.Columns(3).NumberFormat  = '0'
 oExcel.Columns(3).ColumnWidth   = 10
 oExcel.Columns(4).NumberFormat  = '0'
 oExcel.Columns(4).ColumnWidth   = 10
 oExcel.Columns(5).NumberFormat  = '0'
 oExcel.Columns(5).ColumnWidth   = 10
 oExcel.Columns(6).NumberFormat  = '0'
 oExcel.Columns(6).ColumnWidth   = 10
 oExcel.Columns(7).NumberFormat  = '0'
 oExcel.Columns(7).ColumnWidth   = 10
 oExcel.Columns(8).NumberFormat  = '0'
 oExcel.Columns(8).ColumnWidth   = 10
 oExcel.Columns(9).NumberFormat  = '0'
 oExcel.Columns(9).ColumnWidth   = 10
 oExcel.Columns(10).NumberFormat = '0'
 oExcel.Columns(10).ColumnWidth  = 10
 oExcel.Columns(11).NumberFormat = '0'
 oExcel.Columns(11).ColumnWidth  = 10
 oExcel.Columns(12).NumberFormat = '0'
 oExcel.Columns(12).ColumnWidth  = 10
 oExcel.Columns(13).NumberFormat = '0'
 oExcel.Columns(13).ColumnWidth  = 10
 oExcel.Columns(14).NumberFormat = '0'
 oExcel.Columns(14).ColumnWidth  = 10
 
 m.n = 4
 
 WAIT "–¿—◊≈“..." WINDOW NOWAIT 
 SELECT aisoms
 SET RELATION TO mcod INTO sprlpu
 SCAN 
  m.mcod     = mcod
  m.lpuname  = sprlpu.name
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('merror')<=0
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  SELECT merror 
  COUNT FOR et=m.et TO m.nerrs
  IF m.nerrs<=0
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF
  
  CREATE CURSOR curdsp (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol 
  CREATE CURSOR curamb (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol 
  CREATE CURSOR curdst (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol 
  CREATE CURSOR curst (c_i c(30))
  INDEX ON c_i TAG c_i
  
  CREATE CURSOR curdspdef (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol 
  CREATE CURSOR curambdef (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol 
  CREATE CURSOR curdstdef (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol 
  CREATE CURSOR curstdef (c_i c(30))
  INDEX ON c_i TAG c_i
  
  m.dspdefs = 0
  m.ambdefs = 0
  m.dstdefs = 0
  m.stdefs  = 0

  SELECT merror
  SET RELATION TO recid INTO talon 
  SCAN 
   IF et != m.et
    LOOP 
   ENDIF 
   m.c_i    = talon.c_i
   m.sn_pol = talon.sn_pol
   m.cod    = cod
   m.profil = talon.profil
   m.lIsDsp = IIF(SEEK(m.cod, 'dspcodes'), .T., .F.)
   m.err    = UPPER(LEFT(err_mee,2))
   m.lIsDef = IIF(m.err!='W0', .T., .F.)

   DO CASE 
    CASE IsUsl(m.cod)
     IF !SEEK(m.sn_pol, 'curamb')
      INSERT INTO curamb FROM MEMVAR 
     ENDIF 
     IF m.lIsDef
      IF !SEEK(m.sn_pol, 'curambdef')
       INSERT INTO curambdef FROM MEMVAR 
      ENDIF 
     ENDIF 
     m.ambdefs = m.ambdefs + IIF(m.lIsDef, 1, 0)
     IF m.lIsDsp
      IF !SEEK(m.sn_pol, 'curdsp')
       INSERT INTO curdsp FROM MEMVAR 
      ENDIF 
      IF m.lIsDef
       IF !SEEK(m.sn_pol, 'curdspdef')
        INSERT INTO curdspdef FROM MEMVAR 
       ENDIF 
      ENDIF 
      m.dspdefs = m.dspdefs + IIF(m.lIsDef, 1, 0)
     ENDIF 

    CASE IsKD(m.cod)
     IF !SEEK(m.sn_pol, 'curdst')
      INSERT INTO curdst FROM MEMVAR 
     ENDIF 
     IF m.lIsDef
      IF !SEEK(m.sn_pol, 'curdstdef')
       INSERT INTO curdstdef FROM MEMVAR 
      ENDIF 
     ENDIF 
     m.dstdefs = m.dstdefs + IIF(m.lIsDef, 1, 0)

    CASE IsMes(m.cod) OR IsVmp(m.cod)
     IF !SEEK(m.c_i, 'curst')
      INSERT INTO curst FROM MEMVAR 
     ENDIF 
     IF m.lIsDef
      IF !SEEK(m.c_i, 'curstdef')
       INSERT INTO curstdef FROM MEMVAR 
      ENDIF 
     ENDIF 
     m.stdefs = m.stdefs + IIF(m.lIsDef, 1, 0)

    OTHERWISE 

   ENDCASE 

  ENDSCAN 
  SET RELATION OFF INTO talon 
  USE IN talon 
  USE IN merror 
  
  WITH oSheet
  .Cells(n,1) = m.mcod
  .Cells(n,2) = m.lpuname

  .Cells(n,3) = RECCOUNT('curdsp')
  .Cells(n,4) = RECCOUNT('curdspdef')
  .Cells(n,5) = m.dspdefs

  .Cells(n,6) = RECCOUNT('curamb')
  .Cells(n,7) = RECCOUNT('curambdef')
  .Cells(n,8) =  m.ambdefs

  .Cells(n,9)  = RECCOUNT('curdst')
  .Cells(n,10) = RECCOUNT('curdstdef')
  .Cells(n,11) =  m.dstdefs
 
  .Cells(n,12) = RECCOUNT('curst')
  .Cells(n,13) = RECCOUNT('curstdef')
  .Cells(n,14) =  m.stdefs
  ENDWITH 
  m.n = m.n + 1

  USE IN curdsp
  USE IN curamb
  USE IN curdst
  USE IN curst

  USE IN curdspdef
  USE IN curambdef
  USE IN curdstdef
  USE IN curstdef

  SELECT aisoms 
  
 ENDSCAN
 USE IN aisoms
 USE IN sprlpu
 USE IN dspcodes

 WAIT CLEAR 

 IF fso.FileExists(pmee+'\'+m.gcperiod+'\'+m.BookName+'.xls')
  TRY 
   fso.DeleteFile(pmee+'\'+m.gcperiod+'\'+m.BookName+'.xls')
   oBook.SaveAs(pmee+'\'+m.gcperiod+'\'+m.BookName,18)
  CATCH  
   MESSAGEBOX('‘¿…À '+m.BookName+'.XLS Œ “–€“!',0+64,'')
  ENDTRY 
 ELSE 
  oBook.SaveAs(pmee+'\'+m.gcperiod+'\'+m.BookName,18)
 ENDIF 
 oExcel.Visible = .t.

RETURN 