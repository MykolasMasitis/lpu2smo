PROCEDURE FormPGMEE(paratip)
 
 PUBLIC m.TipOfMee
 m.TipOfMee = STR(paratip,1)
 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
  'ФОРМУ ПГ?'+CHR(13)+CHR(10),4+32,'МЭЭ')==7
  RETURN 
 ENDIF 
 
 m.pgdat1 = m.tdat1
 m.pgdat2 = m.tdat2
 m.ischecked = .f.
 DO FORM SelPeriod
 IF m.ischecked = .f.
  RETURN 
 ENDIF 

 PUBLIC oExcel AS Excel.Application
 WAIT "Запуск MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 
 m.nmonthes = (MONTH(m.pgdat2) - MONTH(m.pgdat1)) + 12

 DIMENSION PgMee(21,10)
 PgMee = 0
 DIMENSION TipOfErrsSv(3,3)
 TipOfErrsSv = 0
 
 
 FOR m.nmonth = 0 TO m.nmonthes
 
  m.pgdat = GOMONTH(m.pgdat2, -m.nmonthes+m.nmonth)
  m.pgperiod = STR(YEAR(m.pgdat),4)+PADL(MONTH(m.pgdat),2,'0')

  =FormPGMeeOne(m.pgperiod)

 NEXT 
 
 IF fso.FileExists(pOut+'\Pg_MEE.xls')
  fso.DeleteFile(pOut + '\Pg_MEE.xls')
 ENDIF 

 oslast = oexcel.ActiveWorkbook.Worksheets(oexcel.ActiveWorkbook.Worksheets.Count)
 os = oexcel.ActiveWorkbook.Worksheets('Сводка')
 os.Move(,oslast)

 BookName = pOut+'\Pg_MEE'
 oBook.SaveAs(BookName,18)
 oExcel.Visible = .t.
 
 RELEASE m.TipOfMee

RETURN 
 
FUNCTION  FormPGMeeOne(_pgperiod)
 
 m.lcPgPeriod = _pgperiod

 IF !fso.FileExists(pBase+'\'+lcPgPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&lcPgPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+lcPgPeriod+'\'+'nsi'+'\TarifN', 'Tarif', 'SHARED', 'cod ') > 0
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 
 SCAN 
  m.mcod = mcod
  m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

  WAIT m.mcod WINDOW NOWAIT 
  
  IF !fso.FolderExists(pbase+'\'+lcPgPeriod+'\'+m.mcod)
*   MESSAGEBOX(CHR(13)+CHR(10)+'ДИРЕКТОРИЯ '+m.mcod+' ОТСУТСТВУЕТ!'+CHR(13)+CHR(10),0+48,'')
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\talon.dbf')
*   MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ TALON.DBF ОТСУТСТВУЕТ!'+CHR(13)+CHR(10),0+48, m.mcod)
   LOOP 
  ENDIF 
  
  IF OpenFile(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 

  IF OpenFile(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   LOOP 
  ENDIF 
  

  DIMENSION TipOfErrs(3,3)
  TipOfErrs = 0
  m.WasExp = .f.

  SELECT merror
  SCAN 
   m.cod = cod 
   m.err_mee = LEFT(UPPER(ALLTRIM(err_mee)),2)
   m.e_period = e_period
   
   IF EMPTY(m.err_mee)
    LOOP 
   ELSE 
    m.WasExp = .t.
   ENDIF 
   IF et != m.TipOfMee
    m.WasExp = .f.
    LOOP 
   ENDIF 
   IF EMPTY(e_period)
    m.WasExp = .f.
    LOOP 
   ENDIF 

   m.eperiod = '01.'+SUBSTR(m.e_period,5,2)+'.'+LEFT(m.e_period,4)
   TRY 
    m.eperiod = CTOD(m.eperiod)
   CATCH 
    m.WasExp = .f.
    LOOP 
   ENDTRY 
   
   IF !BETWEEN(m.eperiod,m.pgdat1,m.pgdat2)
    m.WasExp = .f.
    LOOP 
   ENDIF 
   
   DO CASE 
    CASE IsUsl(m.cod) && Аммб-пол.
     DO CASE 
      CASE INLIST(m.err_mee,'EP','NP','EM') && В строку 4.3
       TipOfErrs(1,1) = TipOfErrs(1,1) + 1
       TipOfErrsSv(1,1) = TipOfErrsSv(1,1) + 1
      CASE INLIST(m.err_mee,'W2','WO','EW','PP','UU','GG') && В строку 9
       TipOfErrs(2,1) = TipOfErrs(2,1) + 1
       TipOfErrsSv(2,1) = TipOfErrsSv(2,1) + 1
      CASE m.err_mee = 'W0'
      CASE EMPTY(m.err_mee)
      OTHERWISE && В строку 10
       TipOfErrs(3,1) = TipOfErrs(3,1) + 1
       TipOfErrsSv(3,1) = TipOfErrsSv(3,1) + 1
     ENDCASE 
    CASE IsMes(m.cod) OR IsVMP(m.cod) && Ст.
     DO CASE 
      CASE INLIST(m.err_mee,'EP','NP','EM') && В строку 4.3
       TipOfErrs(1,2) = TipOfErrs(1,2) + 1
       TipOfErrsSv(1,2) = TipOfErrsSv(1,2) + 1
      CASE INLIST(m.err_mee,'W2','WO','EW','PP','UU','GG') && В строку 9
       TipOfErrs(2,2) = TipOfErrs(2,2) + 1
       TipOfErrsSv(2,2) = TipOfErrsSv(2,2) + 1
      CASE m.err_mee = 'W0'
      CASE EMPTY(m.err_mee)
      OTHERWISE && В строку 10
       TipOfErrs(3,2) = TipOfErrs(3,2) + 1
       TipOfErrsSv(3,2) = TipOfErrsSv(3,2) + 1
     ENDCASE 
    CASE IsKD(m.cod) && Дневн. ст.
     DO CASE 
      CASE INLIST(m.err_mee,'EP','NP','EM') && В строку 4.3
       TipOfErrs(1,3) = TipOfErrs(1,3) + 1
       TipOfErrsSv(1,3) = TipOfErrsSv(1,3) + 1
      CASE INLIST(m.err_mee,'W2','WO','EW','PP','UU','GG') && В строку 9
       TipOfErrs(2,3) = TipOfErrs(2,3) + 1
       TipOfErrsSv(2,3) = TipOfErrsSv(2,3) + 1
      CASE m.err_mee = 'W0'
      CASE EMPTY(m.err_mee)
      OTHERWISE && В строку 10
       TipOfErrs(3,3) = TipOfErrs(3,3) + 1
       TipOfErrsSv(3,3) = TipOfErrsSv(3,3) + 1
     ENDCASE 
   ENDCASE 
   WAIT CLEAR 
  ENDSCAN 
  USE IN merror
  USE IN talon 
 
  IF m.WasExp = .f.
   LOOP 
  ENDIF 
  
  TRY 
   oSheet = oBook.WorkSheets(m.mcod)
  CATCH 
   IF oexcel.ActiveSheet.name!='Лист1'
    oSheet = oBook.WorkSheets.Add(,oexcel.ActiveSheet)
   ELSE 
    oSheet = oexcel.ActiveSheet
   ENDIF 
   oSheet.Name = m.mcod
  ENDTRY 

*  oSheet.Select
  FOR iii=1 TO 1
   oexcel.Columns(iii).NumberFormat='@'
  ENDFOR 
*  oSheet.Select

  =PutTextOnThePage()


*  WITH oExcel.ActiveSheet 
  WITH oSheet

   PgMee(03,03) = PgMee(03,03)+(aisoms.ambchkdmee)
   PgMee(03,04) = PgMee(03,04)+(aisoms.stchkdmee)
   PgMee(03,05) = PgMee(03,05)+(aisoms.dstchkdmee)
   PgMee(03,06) = PgMee(03,06)+(aisoms.ambchkdmee+aisoms.dstchkdmee+aisoms.stchkdmee)
   .Cells(03,3) = IIF(!ISNULL(.Cells(03,3).Value),.Cells(03,3).Value,0) + aisoms.ambchkdmee
   .Cells(03,4) = IIF(!ISNULL(.Cells(03,4).Value),.Cells(03,4).Value,0) + aisoms.stchkdmee
   .Cells(03,5) = IIF(!ISNULL(.Cells(03,5).Value),.Cells(03,5).Value,0) + aisoms.dstchkdmee
   .Cells(03,6) = IIF(!ISNULL(.Cells(03,6).Value),.Cells(03,6).Value,0) + aisoms.ambchkdmee+aisoms.dstchkdmee+aisoms.stchkdmee

   .Cells(04,03) = 0
   .Cells(04,04) = 0
   .Cells(04,05) = 0
   .Cells(04,06) = 0

   .Cells(05,3) = 0
   .Cells(05,4) = 0
   .Cells(05,5) = 0
   .Cells(05,6) = 0
   .Cells(05,3) = IIF(!ISNULL(.Cells(05,3).Value),.Cells(05,3).Value,0) + aisoms.ambchkdmee
   .Cells(05,4) = IIF(!ISNULL(.Cells(05,4).Value),.Cells(05,4).Value,0) + aisoms.stchkdmee
   .Cells(05,5) = IIF(!ISNULL(.Cells(05,5).Value),.Cells(05,5).Value,0) + aisoms.dstchkdmee
   .Cells(05,6) = IIF(!ISNULL(.Cells(05,6).Value),.Cells(05,6).Value,0) + aisoms.ambchkdmee+aisoms.dstchkdmee+aisoms.stchkdmee

   PgMee(06,03) = PgMee(06,03) + (aisoms.ambbadmee)
   PgMee(06,04) = PgMee(06,04) + (aisoms.stbadmee)
   PgMee(06,05) = PgMee(06,05) + (aisoms.dstbadmee)
   PgMee(06,06) = PgMee(06,06) + (aisoms.ambbadmee + aisoms.stbadmee + aisoms.dstbadmee)
   .Cells(06,03) = IIF(!ISNULL(.Cells(06,3).Value),.Cells(06,3).Value,0) + aisoms.ambbadmee
   .Cells(06,04) = IIF(!ISNULL(.Cells(06,4).Value),.Cells(06,4).Value,0) + aisoms.stbadmee
   .Cells(06,05) = IIF(!ISNULL(.Cells(06,5).Value),.Cells(06,5).Value,0) + aisoms.dstbadmee
   .Cells(06,06) = IIF(!ISNULL(.Cells(06,6).Value),.Cells(06,6).Value,0) + aisoms.ambbadmee + aisoms.stbadmee + aisoms.dstbadmee

   .Cells(11,3) = IIF(!ISNULL(.Cells(11,3).Value),.Cells(11,3).Value,0) + TipOfErrs(1,1)
   .Cells(11,4) = IIF(!ISNULL(.Cells(11,4).Value),.Cells(11,4).Value,0) + TipOfErrs(1,2)
   .Cells(11,5) = IIF(!ISNULL(.Cells(11,5).Value),.Cells(11,5).Value,0) + TipOfErrs(1,3)
   .Cells(11,6) = IIF(!ISNULL(.Cells(11,6).Value),.Cells(11,6).Value,0) + TipOfErrs(1,1) + TipOfErrs(1,2) + TipOfErrs(1,3)

   .Cells(20,3) = IIF(!ISNULL(.Cells(20,3).Value),.Cells(20,3).Value,0) + TipOfErrs(2,1)
   .Cells(20,4) = IIF(!ISNULL(.Cells(20,4).Value),.Cells(20,4).Value,0) + TipOfErrs(2,2)
   .Cells(20,5) = IIF(!ISNULL(.Cells(20,5).Value),.Cells(20,5).Value,0) + TipOfErrs(2,3)
   .Cells(20,6) = IIF(!ISNULL(.Cells(20,6).Value),.Cells(20,6).Value,0) + TipOfErrs(2,1) + TipOfErrs(2,2) + TipOfErrs(2,3)

   .Cells(21,3) = IIF(!ISNULL(.Cells(21,3).Value),.Cells(21,3).Value,0) +  TipOfErrs(3,1)
   .Cells(21,4) = IIF(!ISNULL(.Cells(21,4).Value),.Cells(21,4).Value,0) +  TipOfErrs(3,2)
   .Cells(21,5) = IIF(!ISNULL(.Cells(21,5).Value),.Cells(21,5).Value,0) +  TipOfErrs(3,3)
   .Cells(21,6) = IIF(!ISNULL(.Cells(21,6).Value),.Cells(21,6).Value,0) +  TipOfErrs(3,1) + TipOfErrs(3,2) + TipOfErrs(3,3)

  ENDWITH 

  RELEASE TipOfErrs
  
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE 
 USE IN tarif

 TRY 
  oSheet = oBook.WorkSheets('Сводка')
 CATCH 
  oSheet = oBook.WorkSheets.Add(,oexcel.ActiveSheet)
  oSheet.name = 'Сводка'
 ENDTRY 

 =PutTextOnThePage() && Формирование сводного листа

* WITH oExcel.ActiveSheet 
 WITH oSheet
  
  .Cells(03,03) = IIF(!ISNULL(.Cells(03,3).Value),.Cells(03,3).Value,0) + PgMee(03,03)
  .Cells(03,04) = IIF(!ISNULL(.Cells(03,4).Value),.Cells(03,4).Value,0) + PgMee(03,04)
  .Cells(03,05) = IIF(!ISNULL(.Cells(03,5).Value),.Cells(03,5).Value,0) + PgMee(03,05)
  .Cells(03,06) = IIF(!ISNULL(.Cells(03,6).Value),.Cells(03,6).Value,0) + PgMee(03,06)

  .Cells(04,03) = 0
  .Cells(04,04) = 0
  .Cells(04,05) = 0
  .Cells(04,06) = 0

  .Cells(05,03) = IIF(!ISNULL(.Cells(05,3).Value),.Cells(05,3).Value,0) + PgMee(03,03)
  .Cells(05,04) = IIF(!ISNULL(.Cells(05,4).Value),.Cells(05,4).Value,0) + PgMee(03,04)
  .Cells(05,05) = IIF(!ISNULL(.Cells(05,5).Value),.Cells(05,5).Value,0) + PgMee(03,05)
  .Cells(05,06) = IIF(!ISNULL(.Cells(05,6).Value),.Cells(05,6).Value,0) + PgMee(03,06)

  .Cells(06,03) = IIF(!ISNULL(.Cells(06,3).Value),.Cells(06,3).Value,0) + PgMee(06,03)
  .Cells(06,04) = IIF(!ISNULL(.Cells(06,4).Value),.Cells(06,4).Value,0) + PgMee(06,04)
  .Cells(06,05) = IIF(!ISNULL(.Cells(06,5).Value),.Cells(06,5).Value,0) + PgMee(06,05)
  .Cells(06,06) = IIF(!ISNULL(.Cells(06,6).Value),.Cells(06,6).Value,0) + PgMee(06,06)

  .Cells(11,3) = IIF(!ISNULL(.Cells(11,3).Value),.Cells(11,3).Value,0) + TipOfErrsSv(1,1)
  .Cells(11,4) = IIF(!ISNULL(.Cells(11,4).Value),.Cells(11,4).Value,0) + TipOfErrsSv(1,2)
  .Cells(11,5) = IIF(!ISNULL(.Cells(11,5).Value),.Cells(11,5).Value,0) + TipOfErrsSv(1,3)
  .Cells(11,6) = IIF(!ISNULL(.Cells(11,6).Value),.Cells(11,6).Value,0) + TipOfErrsSv(1,1) + TipOfErrsSv(1,2) + TipOfErrsSv(1,3)

  .Cells(20,3) = IIF(!ISNULL(.Cells(20,3).Value),.Cells(20,3).Value,0) + TipOfErrsSv(2,1)
  .Cells(20,4) = IIF(!ISNULL(.Cells(20,4).Value),.Cells(20,4).Value,0) + TipOfErrsSv(2,2)
  .Cells(20,5) = IIF(!ISNULL(.Cells(20,5).Value),.Cells(20,5).Value,0) + TipOfErrsSv(2,3)
  .Cells(20,6) = IIF(!ISNULL(.Cells(20,6).Value),.Cells(20,6).Value,0) + TipOfErrsSv(2,1) + TipOfErrsSv(2,2) + TipOfErrsSv(2,3)

  .Cells(21,3) = IIF(!ISNULL(.Cells(21,3).Value),.Cells(21,3).Value,0) + TipOfErrsSv(3,1)
  .Cells(21,4) = IIF(!ISNULL(.Cells(21,4).Value),.Cells(21,4).Value,0) + TipOfErrsSv(3,2)
  .Cells(21,5) = IIF(!ISNULL(.Cells(21,5).Value),.Cells(21,5).Value,0) + TipOfErrsSv(3,3)
  .Cells(21,6) = IIF(!ISNULL(.Cells(21,6).Value),.Cells(21,6).Value,0) + TipOfErrsSv(3,1) + TipOfErrsSv(3,2) + TipOfErrsSv(3,3)

  FOR iii=2 TO 10
   oexcel.Columns(iii).AutoFit
  ENDFOR 
  
 ENDWITH 

RETURN 

FUNCTION PutTextOnThePage
  oExcel.Range(oexcel.Cells(01,01),oexcel.Cells(01,06)).Merge
  WITH oExcel.ActiveSheet 
   .Cells(01,01) = 'Таблица 3.3 формы ПГ'
   .Cells(01,01).HorizontalAlignment = -4152
   .Cells(03,01) = 'Количество проведенных плановых медико-экономических экспертиз'
   .Cells(03,2) = '1'
   .Cells(04,01) = 'в т.ч. тематических'
   .Cells(04,2) = '1.1'
   .Cells(05,01) = 'Всего рассмотрено страховых случаев при проведении ;
    плановых медико-экономических экспертиз'
   .Cells(05,2) = '2'
   .Cells(06,01) = 'выявлено страховых случаев, содержащих нарушения'
   .Cells(06,2) = '2.1'
   .Cells(07,01) = 'Выявлено нарушений, всего, в т.ч.:'
   .Cells(07,2) = '3'
   .Cells(08,01) = 'дефекты оформления первичной медицинской документации, всего, в т.ч.:' 
   .Cells(08,2) = '4'
   .Cells(09,01) = 'непредставление первичной медицинской документации без уважительных причин'
   .Cells(09,2) = '4.1'
   .Cells(10,01) = 'дефекты оформления и ведения первичной документации'
   .Cells(10,2) = '4.2'
   .Cells(11,01) = 'несоответствие данных первичной документации данным счетов (реестра счетов)'
   .Cells(11,2) = '4.3'
   .Cells(12,01) = 'нарушения при оказании медицинской помощи, всего, в т.ч.'
   .Cells(12,2) = '5.'
   .Cells(13,01) = 'нарушения в выполнении необходимых мероприятий в соответствии ;
    с порядком и (или) стандартами медицинской помощи'
   .Cells(13,2) = '5.1'
   .Cells(14,01) = 'необоснованное несоблюдение сроков оказания медицинской помощи'
   .Cells(14,2) = '5.2'
   .Cells(15,01) = 'нарушения, связанные с госпитализацией застрахованного лица'
   .Cells(15,2) = '5.3'
   .Cells(16,01) = 'нарушения информированности застрахованных лиц'
   .Cells(16,2) = '6'
   .Cells(17,01) = 'нарушения, ограничивающие доступность медицинской помощи ;
    для застрахованных лиц, всего, в т.ч.'
   .Cells(17,2) = '7'
   .Cells(18,01) = 'нарушения условий оказания медицинской помощи, в том числе, ;
    сроков ожидания медицинской помощи, предоставляемой в плановом порядке'
   .Cells(18,2) = '7.1'
   .Cells(19,01) = 'взимание платы с застрахованных лиц за медицинскую помощь'
   .Cells(19,2) = '8'
   .Cells(20,01) = 'нарушения, связанные с предъявлением на оплату счетов и реестров счетов'
   .Cells(20,2) = '9'
   .Cells(21,01) = 'прочие нарушения в соответствии с Перечнем'
   .Cells(21,2) = '10'

   .Cells(02,03) = 'Амбулаторно-поликлиническая помощь'
   .Cells(02,03).Orientation = 90
   .Cells(02,04) = 'Стационарная помощь'
   .Cells(02,04).Orientation = 90
   .Cells(02,05) = 'Стационар-замещающая помощь'
   .Cells(02,05).Orientation = 90
   .Cells(02,06) = 'Всего'
   .Cells(02,06).Orientation = 90

   .Rows("1:1").WrapText = .t.
   .Columns("A:A").ColumnWidth = 45
   .Columns("A:A").WrapText = .t.

  ENDWITH 
  
  WITH oExcel.Range(oexcel.Cells(1,1),oexcel.Cells(21,6))
    .Font.Name='Calibri'
    .Font.Size = 10
    .Borders(07).LineStyle = 1 
    .Borders(07).Weight    = 4
    .Borders(08).LineStyle = 1
    .Borders(08).Weight    = 4
    .Borders(09).LineStyle = 1
    .Borders(09).Weight    = 4
    .Borders(10).LineStyle = 1
    .Borders(10).Weight    = 4
    .Borders(11).LineStyle = 1
    .Borders(11).Weight    = 2
    .Borders(12).LineStyle = 1
    .Borders(12).Weight    = 2
  ENDWITH 

 oExcel.ActiveSheet.Cells(01,01).Font.Size = 11

RETURN 