PROCEDURE FormSh6

IF OpenFile(pcommon+'\usvmpxx', 'usvmp', 'shar', 'cod')>0
 IF USED('usvmp')
  USE IN usvmp
 ENDIF 
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
 IF USED('usvmp')
  USE IN usvmp
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('usvmp')
  USE IN usvmp
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 RETURN 
ENDIF 

SELECT aisoms
SET RELATION TO lpuid INTO sprlpu

IF !fso.FolderExists(pout+'\'+gcperiod)
 fso.CreateFolder(pout+'\'+gcperiod)
ENDIF 

pdir = pout+'\'+gcperiod

PUBLIC oExcel AS Excel.Application

WAIT "Запуск MS Excel..." WINDOW NOWAIT 
TRY 
 oExcel=GETOBJECT(,"Excel.Application")
CATCH 
 oExcel=CREATEOBJECT("Excel.Application")
ENDTRY 
WAIT CLEAR 

WAIT "ФОРМИРОВАНИЕ ОТЧЕТА..." WINDOW NOWAIT 

oexcel.UseSystemSeparators = .F.
oexcel.DecimalSeparator = '.'

oexcel.ReferenceStyle= -4150  && xlR1C1
 
oexcel.SheetsInNewWorkbook=1
oBook = oExcel.WorkBooks.Add
oexcel.Cells.Font.Name='Calibri'
oexcel.ActiveSheet.PageSetup.Orientation=2

BookName = pdir+'\FormSh6'
oSheet = oBook.WorkSheets(1)
oSheet.Select
 
nCell = 1

WITH oExcel
 .Columns(01).ColumnWidth = 3
 .Columns(02).ColumnWidth = 30
 .Columns(02).NumberFormat = '@'
 .Columns(02).WrapText = .t.
 .Columns(03).NumberFormat = '@'
 .Columns(04).NumberFormat = "#0"
 .Columns(05).NumberFormat = "#0"
 .Columns(06).NumberFormat = "#,##0.00"
 .Columns(06).ColumnWidth  = 12
 .Columns(07).NumberFormat = "#0"
 .Columns(08).NumberFormat = "#0"
 .Columns(09).NumberFormat = "#,##0.00"
 .Columns(09).ColumnWidth  = 12
 .Columns(10).NumberFormat = "#0"
 .Columns(11).NumberFormat = "#0"
 .Columns(12).NumberFormat = "#,##0.00"
 .Columns(12).ColumnWidth  = 12
 .Columns(13).NumberFormat = "#0"
 .Columns(14).NumberFormat = "#0"
 .Columns(15).NumberFormat = "#,##0.00"
 .Columns(15).ColumnWidth  = 12
 .Columns(16).NumberFormat = "#0"
 .Columns(17).NumberFormat = "#0"
 .Columns(18).NumberFormat = "#,##0.00"
 .Columns(18).ColumnWidth  = 12
 .Columns(19).NumberFormat = "#0"
 .Columns(20).NumberFormat = "#0"
 .Columns(21).NumberFormat = "#,##0.00"
 .Columns(21).ColumnWidth  = 12
ENDWITH 

WITH oExcel.Sheets(1)
 .cells(1,1).Value2 = ''
 .cells(2,1).Value2 = ''
 .cells(3,1).Value2 = 'Сведения об оказании застрахованному лицу медицинской помощи (данные к Разделу II Формы №1 (для СМО))'
 .cells(4,1).Value2 = 'за '+ NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' года'
 .cells(1,1).Font.Size = 11
 .cells(2,1).Font.Size = 11
 .cells(3,1).Font.Size = 11
 .cells(4,1).Font.Size = 11
 .cells(1,1).Font.Bold = .T.
 .cells(2,1).Font.Bold = .T.
 .cells(3,1).Font.Bold = .T.
 .cells(4,1).Font.Bold = .T.
 .cells(1,1).HorizontalAlignment=-4108
 .cells(2,1).HorizontalAlignment=-4108
 .cells(3,1).HorizontalAlignment=-4108
 .cells(4,1).HorizontalAlignment=-4108
 FOR nRow=1 TO 4
  oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,21))
  oRange.Merge
 ENDFOR  
 
ENDWITH 

oExcel.Range(oExcel.Cells(5,04), oExcel.Cells(5,12)).Merge 
oExcel.Cells(5,04).Value='Первичная МСП'
oExcel.Cells(5,04).HorizontalAlignment=-4108

oExcel.Range(oExcel.Cells(5,13), oExcel.Cells(5,18)).Merge 
oExcel.Cells(5,13).Value='Специализированная МСП'
oExcel.Cells(5,13).HorizontalAlignment=-4108

oExcel.Range(oExcel.Cells(6,04), oExcel.Cells(6,06)).Merge 
oExcel.Range(oExcel.Cells(6,07), oExcel.Cells(6,09)).Merge 
oExcel.Range(oExcel.Cells(6,10), oExcel.Cells(6,12)).Merge 
oExcel.Range(oExcel.Cells(6,13), oExcel.Cells(6,15)).Merge 
oExcel.Range(oExcel.Cells(6,16), oExcel.Cells(6,18)).Merge 
oExcel.Range(oExcel.Cells(6,19), oExcel.Cells(6,21)).Merge 

oExcel.Cells(6,04).Value2 = 'Амб.-пол.'
oExcel.Cells(6,04).HorizontalAlignment=-4108
oExcel.Cells(6,07).Value2 = 'Стоматология'
oExcel.Cells(6,07).HorizontalAlignment=-4108
oExcel.Cells(6,10).Value2 = 'Дневной стационар'
oExcel.Cells(6,10).HorizontalAlignment=-4108

oExcel.Cells(6,13).Value2 = 'Амб.-пол.'
oExcel.Cells(6,13).HorizontalAlignment=-4108
oExcel.Cells(6,16).Value2 = 'Cтационар'
oExcel.Cells(6,16).HorizontalAlignment=-4108

oExcel.Cells(6,19).Value2 = 'Скорая помощь'
oExcel.Cells(6,19).HorizontalAlignment=-4108


oExcel.Cells(7,04).Value2 = 'Счетов'
oExcel.Cells(7,05).Value2 = 'Услуг'
oExcel.Cells(7,06).Value2 = 'Сумма'
oExcel.Cells(7,07).Value2 = 'Счетов'
oExcel.Cells(7,08).Value2 = 'Услуг'
oExcel.Cells(7,09).Value2 = 'Сумма'
oExcel.Cells(7,10).Value2 = 'Счетов'
oExcel.Cells(7,11).Value2 = 'К/дн.'
oExcel.Cells(7,12).Value2 = 'Сумма'

oExcel.Cells(7,13).Value2 = 'Счетов'
oExcel.Cells(7,14).Value2 = 'Услуг'
oExcel.Cells(7,15).Value2 = 'Сумма'
oExcel.Cells(7,16).Value2 = 'Счетов'
oExcel.Cells(7,17).Value2 = 'МЭСов'
oExcel.Cells(7,18).Value2 = 'Сумма'
oExcel.Cells(7,19).Value2 = 'Счетов'
oExcel.Cells(7,20).Value2 = 'Услуг'
oExcel.Cells(7,21).Value2 = 'Сумма'

m.fstring = 7 && Первая строка

WITH oExcel.Sheets(1)
 .Cells(m.fstring,01).Value2 = '№ п\п'
 .Cells(m.fstring,02).Value2 = 'Наименование ЛПУ'
 .Cells(m.fstring,03).Value2 = 'Код ЛПУ'
  oExcel.Range(oExcel.Cells(m.fstring-1,1), oExcel.Cells(m.fstring,1)).Merge 
  oExcel.Range(oExcel.Cells(m.fstring-1,2), oExcel.Cells(m.fstring,2)).Merge 
  oExcel.Range(oExcel.Cells(m.fstring-1,3), oExcel.Cells(m.fstring,3)).Merge 
  .Cells(m.fstring-1,5).Value2 = ''
ENDWITH 

FOR ncol=1 TO 21
 WITH oExcel.Sheets(1)
  .Cells(m.fstring,ncol).Font.Size = 8
  .Cells(m.fstring,ncol).Font.Bold = .F.
  .Cells(m.fstring,ncol).WrapText = .t.
  .Cells(m.fstring,ncol).HorizontalAlignment = -4108
  .Cells(m.fstring,ncol).VerticalAlignment = -4108
 ENDWITH 
NEXT 

m.fstring = m.fstring + 1
FOR ncolumn=1 TO 21
 WITH oExcel.Sheets(1)
  .cells(m.fstring,ncolumn).Value2 = STR(ncolumn,2)
  .cells(m.fstring,ncolumn).HorizontalAlignment = -4108
 ENDWITH 
NEXT 

oExcel.Range(oExcel.Cells(m.fstring,1), oExcel.Cells(m.fstring,21)).NumberFormat='@'

m.fstring = m.fstring + 1

m.mcod = mcod
m.lpuname = sprlpu.name
*oExcel.Sheets(1).cells(m.fstring,2).Value2 = m.lpuname
oRange = oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,21))
oRange.Merge
oRange.HorizontalAlignment = -4108
oRange.Interior.ColorIndex = 40

m.nnn = 0

m.spaz1 = 0
m.susl1 = 0
m.ssum1 = 0
m.spaz2 = 0
m.susl2 = 0
m.ssum2 = 0
m.spaz3 = 0
m.susl3 = 0
m.ssum3 = 0
m.spaz4 = 0
m.susl4 = 0
m.ssum4 = 0
m.spaz6 = 0
m.susl6 = 0
m.ssum6 = 0
m.spaz7 = 0
m.susl7 = 0
m.ssum7 = 0

SCAN FOR aisoms.s_pred>0
 m.mcod = mcod
 
 WAIT m.mcod+'...' WINDOW NOWAIT 

 lcDir = pBase + '\' + m.gcperiod + '\' + m.mcod
 IF !fso.FolderExists(lcDir)
  LOOP 
 ENDIF 
 IF !fso.FileExists(lcDir+'\talon.dbf') OR !fso.FileExists(lcDir+'\e'+mcod+'.dbf')
  LOOP 
 ENDIF 
 IF OpenFile("&lcDir\Talon", "Talon", "SHARE", 'sn_pol')>0 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  SELECT AisOms
  LOOP 
 ENDIF 
 IF OpenFile("&lcDir\people", "people", "SHARE", 'sn_pol')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  SELECT AisOms
  LOOP 
 ENDIF 
 IF OpenFile("&lcDir\e"+mcod, "Error", "SHARE", 'rid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('error')
   USE IN error
  ENDIF 
  SELECT AisOms
  LOOP 
 ENDIF 
 IF OpenFile("&lcDir\m"+mcod, "mError", "SHARE", 'recid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('error')
   USE IN error
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT AisOms
  LOOP 
 ENDIF 

 SELECT Talon
 SET RELATION TO recid INTO error 
 SET RELATION TO recid INTO merror ADDITIVE 
 SET RELATION TO cod INTO usvmp ADDITIVE 

 CREATE CURSOR paz1 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR paz2 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR paz3 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR paz4 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR paz6 (c_i c(30))
 INDEX on c_i TAG c_i
 SET ORDER TO c_i
 CREATE CURSOR paz7 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol


 SELECT AisOms

 m.nnn = m.nnn+1
 m.fstring = m.fstring + 1
 m.lpuname = sprlpu.name

 m.paz1 = 0
 m.usl1 = 0
 m.sum1 = 0
 m.paz2 = 0
 m.usl2 = 0
 m.sum2 = 0
 m.paz3 = 0
 m.usl3 = 0
 m.sum3 = 0
 m.paz4 = 0
 m.usl4 = 0
 m.sum4 = 0
 m.paz6 = 0
 m.usl6 = 0
 m.sum6 = 0
 m.paz7 = 0
 m.usl7 = 0
 m.sum7 = 0

 SELECT talon 
 SCAN 
  IF !EMPTY(error.c_err)
   LOOP 
  ENDIF 
  IF !EMPTY(merror.recid)
   LOOP 
  ENDIF 
  m.c_i = c_i
  m.sn_pol = sn_pol
  m.vmp = usvmp.vmp146

  DO CASE 
   CASE m.vmp = 1 && Первичная медико-санитарная амбулаторная помощь                                                     
    m.sum1 = m.sum1 + s_all
    m.ssum1 = m.ssum1 + s_all
    m.usl1 = m.usl1 + k_u
    m.susl1 = m.susl1 + k_u
    IF !SEEK(m.sn_pol, 'paz1')
     INSERT INTO paz1 (sn_pol) VALUES (m.sn_pol)
     m.paz1 = m.paz1 + 1
     m.spaz1 = m.spaz1 + 1
    ENDIF 

   CASE m.vmp = 2 && Первичная медико-санитарная стоматологическая помощь                                                
    m.sum2 = m.sum2 + s_all
    m.ssum2 = m.ssum2 + s_all
    m.usl2 = m.usl2 + k_u
    m.susl2 = m.susl2 + k_u
    IF !SEEK(m.sn_pol, 'paz2')
     INSERT INTO paz2 (sn_pol) VALUES (m.sn_pol)
     m.paz2 = m.paz2 + 1
     m.spaz2 = m.spaz2 + 1
    ENDIF 

   CASE m.vmp = 3 && Первичная медико-санитарная  помощь, оказанная в условиях дневных стационаров всех типов            
    m.sum3 = m.sum3 + s_all
    m.ssum3 = m.ssum3 + s_all
    m.usl3 = m.usl3 + 1
    m.susl3 = m.susl3 + 1
    IF !SEEK(m.sn_pol, 'paz3')
     INSERT INTO paz3 (sn_pol) VALUES (m.sn_pol)
     m.paz3 = m.paz3 + 1
     m.spaz3 = m.spaz3 + 1
    ENDIF 

   CASE m.vmp = 4 && Специализированная  медицинская амбулаторная  помощь                                                
    m.sum4 = m.sum4 + s_all
    m.ssum4 = m.ssum4 + s_all
    m.usl4 = m.usl4 + k_u
    m.susl4 = m.susl4 + k_u
    IF !SEEK(m.sn_pol, 'paz4')
     INSERT INTO paz4 (sn_pol) VALUES (m.sn_pol)
     m.paz4 = m.paz4 + 1
     m.spaz4 = m.spaz4 + 1
    ENDIF 

   CASE m.vmp = 6 && Специализированная  медицинская стационарная  помощь                                                
    m.sum6 = m.sum6 + s_all
    m.ssum6 = m.ssum6 + s_all
    m.usl6 = m.usl6 + k_u
    m.susl6 = m.susl6 + k_u
    IF !SEEK(m.c_i, 'paz6')
     INSERT INTO paz6 (c_i) VALUES (m.c_i)
     m.paz6  = m.paz6 + 1
     m.spaz6 = m.spaz6 + 1
    ENDIF 

   CASE m.vmp = 7 && Скорая медицинская помощь                                                                           
    m.sum7 = m.sum7 + s_all
    m.ssum7 = m.ssum7 + s_all
    m.usl7 = m.usl7 + k_u
    m.susl7 = m.susl7 + k_u
    IF !SEEK(m.sn_pol, 'paz7')
     INSERT INTO paz7 (sn_pol) VALUES (m.sn_pol)
     m.paz7  = m.paz7 + 1
     m.spaz7 = m.spaz7 + 1
    ENDIF 

  ENDCASE 


 ENDSCAN 
 SET RELATION OFF INTO error 
 SET RELATION OFF INTO merror
 SET RELATION OFF INTO usvmp
 
 SELECT AisOms

 WITH oExcel.Sheets(1)
  .Cells(m.fstring,01).Value2 = m.nnn
  .Cells(m.fstring,02).Value2 = ALLTRIM(m.lpuname)
  .Cells(m.fstring,03).Value2 = m.mcod
  .Cells(m.fstring,04).Value2 = m.paz1
  .Cells(m.fstring,05).Value2 = m.usl1
  .Cells(m.fstring,06).Value2 = m.sum1
  .Cells(m.fstring,07).Value2 = m.paz2
  .Cells(m.fstring,08).Value2 = m.usl2
  .Cells(m.fstring,09).Value2 = m.sum2
  .Cells(m.fstring,10).Value2 = m.paz3
  .Cells(m.fstring,11).Value2 = m.usl3
  .Cells(m.fstring,12).Value2 = m.sum3
  .Cells(m.fstring,13).Value2 = m.paz4
  .Cells(m.fstring,14).Value2 = m.usl4
  .Cells(m.fstring,15).Value2 = m.sum4
  .Cells(m.fstring,16).Value2 = m.paz6
  .Cells(m.fstring,17).Value2 = m.usl6
  .Cells(m.fstring,18).Value2 = m.sum6
  .Cells(m.fstring,19).Value2 = m.paz7
  .Cells(m.fstring,20).Value2 = m.usl7
  .Cells(m.fstring,21).Value2 = m.sum7

 ENDWITH 

 USE IN paz1
 USE IN paz2
 USE IN paz3
 USE IN paz4
 USE IN paz6
 USE IN paz7

 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('error')
  USE IN error
 ENDIF 
 IF USED('merror')
  USE IN merror
 ENDIF 
 SELECT AisOms

 WAIT CLEAR 

ENDSCAN 

m.fstring = m.fstring + 1

oRange = oExcel.Range(oExcel.Cells(m.fstring,1), oExcel.Cells(m.fstring,3))
oRange.Merge

WITH oExcel.Sheets(1)
 .Cells(m.fstring,01).Value2 = 'Итого:'
 .Cells(m.fstring,04).Value2 = m.spaz1
 .Cells(m.fstring,05).Value2 = m.susl1
 .Cells(m.fstring,06).Value2 = m.ssum1
 .Cells(m.fstring,07).Value2 = m.spaz2
 .Cells(m.fstring,08).Value2 = m.susl2
 .Cells(m.fstring,09).Value2 = m.ssum2
 .Cells(m.fstring,10).Value2 = m.spaz3
 .Cells(m.fstring,11).Value2 = m.susl3
 .Cells(m.fstring,12).Value2 = m.ssum3

 .Cells(m.fstring,13).Value2 = m.spaz4
 .Cells(m.fstring,14).Value2 = m.susl4
 .Cells(m.fstring,15).Value2 = m.ssum4
 .Cells(m.fstring,16).Value2 = m.spaz6
 .Cells(m.fstring,17).Value2 = m.susl6
 .Cells(m.fstring,18).Value2 = m.ssum6
 .Cells(m.fstring,19).Value2 = m.spaz7
 .Cells(m.fstring,20).Value2 = m.susl7
 .Cells(m.fstring,21).Value2 = m.ssum7
ENDWITH 

IF fso.FileExists(pDir+'\FormSh6.xls')
 fso.DeleteFile(pDir+'\FormSh6.xls')
ENDIF 

oBook.SaveAs(BookName,18)

oExcel.Visible = .T.

SET RELATION OFF INTO sprlpu
USE IN aisoms
USE IN sprlpu
USE IN usvmp

WAIT CLEAR 

RETURN 
