PROCEDURE FindPaz
 m.pgdat1 = m.tdat1
 m.pgdat2 = m.tdat2

 m.ischecked=.f.
 DO FORM SelPeriod
 IF m.ischecked=.f.
  RETURN 
 ENDIF 

 m.pgdat1 = CTOD('01.'+PADL(MONTH(m.pgdat1),2,'0')+'.'+STR(YEAR(m.pgdat1),4))
 m.pgdat2 = CTOD('01.'+PADL(MONTH(m.pgdat2),2,'0')+'.'+STR(YEAR(m.pgdat2),4))
 
 m.sn_pol = ''
 m.ischecked=.f.
 DO FORM SelPolis
 IF m.ischecked=.f.
  RETURN 
 ENDIF 
 
 IF EMPTY(m.sn_pol)
  RETURN 
 ENDIF 
 
 m.sn_pol = PADR(ALLTRIM(m.sn_pol),25)
 
 m.lcdat = m.pgdat2

 IF !fso.FileExists(pBase+'\'+m.gcperiod+'\nsi\tarifn.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcperiod+'\nsi\sprlpuxx.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  USE IN tarif 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 

 CREATE CURSOR curpaz (mcod c(7), lpuname c(40), period c(6), sn_pol c(25), c_i c(30), ds c(6), pcod c(10), cod n(6), uslname c(40), k_u n(3), tip c(1), ;
  d_u d, s_all n(11,2), fil_id n(6), d_type c(1))
 INDEX on mcod TAG mcod 
 INDEX on cod TAG cod
 INDEX on d_u TAG d_u
 INDEX on ds TAG ds
 
 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 DO WHILE m.lcdat>=m.pgdat1
  m.lcperiod = LEFT(DTOS(m.lcdat),6)
  
  IF !fso.FolderExists(pBase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0 
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 

  SELECT aisoms 
  WAIT m.lcperiod WINDOW NOWAIT 
  SCAN 
   m.mcod = mcod 
   m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.name, '')
   IF !fso.FolderExists(pBase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'sn_pol')>0
    IF USED('talon')
     USE IN talon 
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+m.lcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
    IF USED('people')
     USE IN people
    ENDIF 
    USE IN talon 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT talon 
   
   IF !SEEK(m.sn_pol)
    USE IN talon
    USE IN people  
    SELECT aisoms 
    LOOP 
   ENDIF 
   
   m.fio = ''
   IF SEEK(m.sn_pol, 'people')
    m.fio = ALLTRIM(people.fam)+' '+ALLTRIM(people.im)+' '+ALLTRIM(people.ot)+', '+DTOC(people.dr)
   ENDIF 
   
   DO WHILE sn_pol=m.sn_pol
    m.mcod   = mcod
    m.period = period
    m.c_i    = c_i
    m.ds     = ds
    m.pcod   = pcod
    m.cod    = cod
    m.k_u    = k_u
    m.tip    = tip
    m.d_u    = d_u
    m.k_u    = k_u
    m.s_all  = s_all
    m.fil_id = fil_id
    m.d_type = d_type
    
    m.uslname = IIF(SEEK(m.cod, 'tarif'), tarif.comment, '')
    
    INSERT INTO curpaz FROM MEMVAR 

    SKIP 
   ENDDO 

   USE IN talon 
   USE IN people
   SELECT aisoms
   
  ENDSCAN 
  WAIT CLEAR 
  USE IN aisoms
  
  m.lcdat = GOMONTH(m.lcdat,-1)

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDDO 
 
 SET ESCAPE &OldEscStatus
 
 SELECT curpaz
 GO TOP 
 
 oms5n(m.sn_pol, .t., .f.)
 
 USE IN curpaz 
 USE IN tarif 
 USE IN sprlpu
 
RETURN 

FUNCTION  oms5n(polis, IsVisible, IsQuit)

# DEFINE xlDiagonalDown	     5	&& Диагональная от верхнего левого угла в нижний правый каждой ячейки в диапазоне
# DEFINE xlDiagonalUp	     6	&& Диагональная из нижнего левого угла в правый верхний каждой ячейки в диапазоне.
# DEFINE xlEdgeBottom	     9	&& Нижнаяя для всего диапазона ячеек
# DEFINE xlEdgeLeft	         7  && Левая для всего диапазона ячеек.
# DEFINE xlEdgeRight	    10	&& Правая для всего диапазона ячеек.
# DEFINE xlEdgeTop	         8	&& Верхняя для всего диапазона ячеек.
# DEFINE xlInsideHorizontal	12	&& Горизонтальные границы всех внутренних ячеек диапазона
# DEFINE xlInsideVertical	11	&& Вертикальные границы всех внутренних ячеек диапазона

# DEFINE xlContinuous	     1	&& Непрерывная линия
# DEFINE xlDash	         -4115	&& Пунктирная линия
# DEFINE xlDashDot	         4	&& Пунктир с точкой
# DEFINE xlDashDotDot	     5	&& Пунктир с двумя идущими подряд точками
# DEFINE xlDot	         -4118	&& Линия из точек
# DEFINE xlDouble	     -4119	&& Двойная линия
# DEFINE xlLineStyleNone -4142	&& Без линий
# DEFINE xlSlantDashDot	    13	&& Наклонная пунктирная

# DEFINE xlHairline	    1	&& Самая тонкая граница
# DEFINE xlMedium	-4138	&& Средняя толщина
# DEFINE xlThick	    4	&& Толстая граница
# DEFINE xlThin	        2   && Тонкая граница

m.sn_pol = polis

pMail = fso.GetParentFolderName(pbin)+'\MEE'

IF !fso.FolderExists(pMail)
 fso.CreateFolder(pMail)
ENDIF 

m.SortTip = '0'
oal = ALIAS()
IF fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\outs.dbf')
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\outs', 'outs', 'shar')>0
  IF USED('outs')
   USE IN outs
  ENDIF 
 ENDIF 
ENDIF 
CREATE CURSOR TipSort (name c(20), cod c(1))
INSERT INTO TipSort (name,cod) VALUES ('Не сортировать','0')
INSERT INTO TipSort (name,cod) VALUES ('Дата услуги/МЭС','1')
INSERT INTO TipSort (name,cod) VALUES ('Код услуги/МЭС','2')
INSERT INTO TipSort (name,cod) VALUES ('Диагноз','3')
SELECT (oal)

DO FORM SortOms5

PUBLIC oExcel AS Excel.Application
WAIT "Запуск MS Excel..." WINDOW NOWAIT 
TRY 
 oExcel=GETOBJECT(,"Excel.Application")
CATCH 
 oExcel=CREATEOBJECT("Excel.Application")
ENDTRY 
WAIT CLEAR 

oexcel.UseSystemSeparators= .F.
oexcel.DecimalSeparator = '.'

oexcel.ReferenceStyle= -4150  && xlR1C1
 
oexcel.SheetsInNewWorkbook=1
oBook = oExcel.WorkBooks.Add
oexcel.Cells.Font.Name='Calibri'
oexcel.ActiveSheet.PageSetup.Orientation=2

BookName = pMail+'\'+'oms5_'+ALLTRIM(sn_pol)
oSheet = oBook.WorkSheets(1)
oSheet.Select
 
FOR iii=1 TO 12
 oexcel.Columns(iii).NumberFormat='@'
ENDFOR 

nCell = 1
orec = RECNO()

LpuName  = ''
CokrCod  = ''
CokrName = ''

m.prmcod = ''
m.lpupr = ''
m.sppr  = ''

*m.tipp = people.tipp
m.ppolis = ''

 WITH oExcel.Sheets(1)
*  .cells(1,1).Value2 = 'ЛПУ: ' + m.lpuname + ', ' + m.cokrname + ', ' + m.mcod
  .cells(1,1).Characters(1,4).Font.Bold = .t.
  .cells(1,1).Characters(1,4).Font.Italic = .t.

  .cells(2,1).Value2 = 'СМО: ' + m.qname
  .cells(2,1).Characters(1,4).Font.Bold = .t.
  .cells(2,1).Characters(1,4).Font.Italic = .t.

  .cells(4,1).Value2 = 'Пациент: ' + m.fio
  .cells(4,1).Characters(1,8).Font.Bold = .t.
  .cells(4,1).Characters(1,8).Font.Italic = .t.

  .cells(5,1).Value2 = 'Полис: ' + ALLTRIM(sn_pol)
  .cells(5,1).Characters(1,6).Font.Bold = .t.
  .cells(5,1).Characters(1,6).Font.Italic = .t.

  .cells(7,1).Value2 = 'Карта: ' + ALLTRIM(c_i)
  .cells(7,1).Characters(1,6).Font.Bold = .t.
  .cells(7,1).Characters(1,6).Font.Italic = .t.

*  .cells(8,1).Value2 = 'ЛПУ прикреления: ' + m.lpupr
  .cells(8,1).Characters(1,16).Font.Bold = .t.
  .cells(8,1).Characters(1,16).Font.Italic = .t.

*  .cells(9,1).Value2 = 'Способ прикрепления: ' + m.sppr
  .cells(9,1).Characters(1,20).Font.Bold = .t.
  .cells(9,1).Characters(1,20).Font.Italic = .t.

  .cells(10,1).Value2 = 'Счет за оказанную медицинскую помощь по Московской городской программе ОМС'
  .cells(11,3).Value2 = 'за '+ NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' года'
  .cells(10,1).Font.Size = 11
  .cells(11,1).Font.Size = 11
  .cells(10,1).Font.Bold = .T.
  .cells(11,1).Font.Bold = .T.
  .cells(10,1).HorizontalAlignment=-4108
  .cells(11,3).HorizontalAlignment=-4108

  FOR nRow=1 TO 11
   oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,11))
   oRange.Merge
  ENDFOR  
 
  .cells(13,1).Value2 = 'Дата'
  .cells(13,2).Value2 = 'Услуга'
  .cells(13,3).Value2 = 'Тип'
  .cells(13,4).Value2 = 'Диагноз'
  .cells(13,5).Value2 = 'Наименование услуги'
  .cells(13,6).Value2 = 'Кол-во'
  .cells(13,7).Value2 = 'Сумма'
  .cells(13,8).Value2 = 'ЛПУ оказ.'
  .cells(13,9).Value2 = 'ЛПУ оказ.'
  
  oal = ALIAS()
  m.polis = sn_pol
  SELECT curpaz
  DO CASE 
   CASE m.SortTip = '0'
    SET ORDER TO 
   CASE m.SortTip = '1'
    SET ORDER TO d_u
   CASE m.SortTip = '2'
    SET ORDER TO cod
   CASE m.SortTip = '3'
    SET ORDER TO ds
  ENDCASE 

  nCell = 13
  m.ttlkol = 0
  m.ttlsum = 0
  SCAN 
   IF sn_pol = m.polis
    m.cod = cod 
    nCell = nCell + 1
    .cells(nCell,1).Value2 = DTOC(d_u)
    .cells(nCell,2).Value2 = PADL(cod,6,'0')
    .cells(nCell,3).Value2 = tip
    .cells(nCell,4).Value2 = ds
    .cells(nCell,5).Value2 = ALLTRIM(uslname)
    .cells(nCell,6).Value2 = STR(k_u,3)
    .cells(nCell,7).Value2 = TRANSFORM(s_all, '99 999 999.99')
    .cells(nCell,8).Value2 = mcod
    .cells(nCell,9).Value2 = lpuname
    m.ttlkol = m.ttlkol + k_u
    m.ttlsum = m.ttlsum + s_all
   ENDIF 
  ENDSCAN 
 .cells(nCell+1,6).Value2 = STR(m.ttlkol,3)
 .cells(nCell+1,7).Value2 = TRANSFORM(m.ttlsum, '99 999 999.99')
  SELECT (oal)
 ENDWITH 

GO (orec)

FOR iii=1 TO 12
 oexcel.Columns(iii).AutoFit
ENDFOR 

m.ttlsumsay = 'ИТОГО: '+cpr(INT(m.ttlsum))+PADL(INT((m.ttlsum-INT(m.ttlsum))*100),2,'0')+' КОП.'
oExcel.Sheets(1).cells(nCell+2,1).Value2 = m.ttlsumsay
oExcel.Range(oExcel.Sheets(1).cells(nCell+2,1), oExcel.Sheets(1).cells(nCell+2,8)).Merge 

IF fso.FileExists(pMail+'\'+'oms5_'+ALLTRIM(sn_pol)+'.xls')
 fso.DeleteFile(pMail+'\'+'oms5_'+ALLTRIM(sn_pol)+'.xls')
ENDIF 

oBook.SaveAs(BookName,18)

IF IsVisible == .T. 
 oExcel.Visible = .T.
ELSE 
 oBook.Close(0)
 IF IsQuit
  oExcel.Quit
 ENDIF 
ENDIF 

RETURN 
