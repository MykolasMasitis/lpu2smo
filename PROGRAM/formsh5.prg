PROCEDURE FormSh5
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ —“¿“»—“» ” —Õﬂ“»… œŒ Ã› ?'+CHR(13)+CHR(10),4+32,'‘Œ–Ã¿ ÿ-5')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee+'\'+m.gcperiod)
  fso.CreateFolder(pmee+'\'+m.gcperiod)
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\sookodxx.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À SOOKODXX.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sookodxx', 'sookod', 'shar', 'er_c')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 
 PUBLIC oExcel AS Excel.Application
 WAIT "«‡ÔÛÒÍ MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 m.BookName = 'sh5'
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
 oBook  = oExcel.WorkBooks.Add
 oSheet = oexcel.ActiveSheet
 oSheet.PageSetup.Orientation=2
 oSheet.name='ÙÓÏ‡ ÿ-5'
 
 WITH oExcel
  .Rows(1).RowHeight=25
  .Rows(2).RowHeight=25
  .Rows(2).WrapText = .t.
  .Cells(1,1) = '—Ú‡ÚËÒÚËÍ‡ ÒÌˇÚËÈ ÔÓ Ã›  Á‡ '+NameOfMonth(m.tmonth)+' '+STR(tyear,4)+' „Ó‰‡'
  .Cells(1,1).HorizontalAlignment = -4108
  .Rows(1).VerticalAlignment = -4108
  .Rows(2).VerticalAlignment = -4108
  .Range(.Cells(1,1),.Cells(1,5)).Merge
  .Cells(2,1)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ Ã√‘ŒÃ—'
  .Cells(2,2)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ ‘‘ŒÃ—'
  .Cells(2,3)='Õ‡ËÏÂÌÓ‚‡ÌËÂ ‰ÂÙÂÍÚ‡'
  .Cells(2,4)=' ÓÎ-‚Ó ‰ÂÙÂÍÚÓ‚'
  .Cells(2,5)='—ÛÏÏ‡ ÒÌˇÚËÈ ÔÓ ‰ÂÙÂÍÚÛ'
 ENDWITH 

 WAIT "–¿—◊≈“..." WINDOW NOWAIT 

 CREATE CURSOR curdata (c_err c(3), k_u n(6), s_all n(11,2))
 INDEX ON c_err TAG c_err
 SET ORDER TO c_err 

 SELECT aisoms
 SCAN 
  m.mcod     = mcod
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'merror', 'shar', 'rid')>0
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
  
  CREATE CURSOR curid (recid i)
  INDEX ON recid TAG recid
  SET ORDER TO recid 

  SELECT merror
  SET RELATION TO rid INTO talon 
  SCAN 
   IF DELETED()
    LOOP 
   ENDIF 
   m.fll = UPPER(f)
   IF m.fll!='S'
    LOOP 
   ENDIF 

   m.recid  = talon.recid
   m.c_err  = c_err
   m.s_all  = talon.s_all
   m.ok_u   = 0
   m.os_all = 0

   IF SEEK(m.c_err, 'curdata')
    m.ok_u   = curdata.k_u
    m.os_all = curdata.s_all
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    m.nk_u   = m.ok_u + 1
    m.ns_all = m.os_all + m.s_all
    UPDATE curdata SET k_u=m.nk_u, s_all=m.ns_all WHERE c_err=m.c_err
   ELSE 
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    INSERT INTO curdata (c_err,k_u,s_all) VALUES (m.c_err,1,m.s_all)
   ENDIF 


  ENDSCAN 
  SET RELATION OFF INTO talon 
  USE IN talon 
  USE IN merror 
  USE IN curid

  SELECT aisoms 
  
 ENDSCAN
 USE IN aisoms
 WAIT CLEAR 

 WITH oExcel
  .Columns(1).NumberFormat  = '@'
  .Columns(1).ColumnWidth   = 5
  .Columns(2).NumberFormat  = '@'
  .Columns(2).ColumnWidth   = 5
  .Columns(3).NumberFormat  = '@'
  .Columns(3).ColumnWidth   = 75
  .Columns(4).NumberFormat  = '0'
  .Columns(4).ColumnWidth   = 10
  .Columns(5).NumberFormat  = '0.00'
  .Columns(5).ColumnWidth   = 20
 ENDWITH 

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Œ“◊≈“¿..." WINDOW NOWAIT 
 SELECT curdata 
 SET RELATION TO LEFT(c_err,2) INTO sookod
 m.n = 3
 m.k_u = 0
 m.s_all = 0 
 SCAN 
  WITH oExcel
   .Cells(m.n,1) = c_err
   .Cells(m.n,2) = sookod.osn230
   .Cells(m.n,3) = sookod.comment
   .Cells(m.n,4) = k_u
   .Cells(m.n,5) = s_all
   m.k_u = m.k_u + k_u
   m.s_all = m.s_all + s_all
  ENDWITH 
  m.n = m.n + 1
 ENDSCAN 
 SET RELATION OFF INTO sookod
 USE IN curdata 
 USE IN sookod
 oExcel.Range(oExcel.Cells(m.n,1),oExcel.Cells(m.n,3)).Merge
 oExcel.Cells(m.n,1) = '»ÚÓ„Ó:'
 oExcel.Rows(m.n).RowHeight=25
* oExcel.Cells(m.n,4).FormulaR1C1 = "=SUM(cells(3,3);cells(&n,3))"
 oExcel.Cells(m.n,4) = m.k_u
 oExcel.Cells(m.n,5) = m.s_all
 oExcel.Rows(m.n).VerticalAlignment = -4108
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