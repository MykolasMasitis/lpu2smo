PROCEDURE FormSh0

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
 USE IN sprlpu
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

BookName = pdir+'\Отчет для Поветко'
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
 .Columns(07).NumberFormat = "#0"
 .Columns(08).NumberFormat = "#0"
 .Columns(09).NumberFormat = "#,##0.00"
 .Columns(10).NumberFormat = "#0"
 .Columns(11).NumberFormat = "#0"
 .Columns(12).NumberFormat = "#,##0.00"
 .Columns(13).NumberFormat = "#0"
 .Columns(14).NumberFormat = "#0"
 .Columns(15).NumberFormat = "#,##0.00"
 .Columns(16).NumberFormat = "#0"
 .Columns(17).NumberFormat = "#0"
 .Columns(18).NumberFormat = "#,##0.00"
 .Columns(19).NumberFormat = "#0"
 .Columns(20).NumberFormat = "#0"
 .Columns(21).NumberFormat = "#,##0.00"
 .Columns(22).NumberFormat = "#0"
 .Columns(23).NumberFormat = "#0"
 .Columns(24).NumberFormat = "#,##0.00"
 .Columns(25).NumberFormat = "#0"
 .Columns(26).NumberFormat = "#0"
 .Columns(27).NumberFormat = "#,##0.00"
 .Columns(28).NumberFormat = "#0"
 .Columns(29).NumberFormat = "#0"
 .Columns(30).NumberFormat = "#0"
 .Columns(31).NumberFormat = "#,##0.00"
ENDWITH 

WITH oExcel.Sheets(1)
 .cells(1,1).Value2 = ''
 .cells(2,1).Value2 = ''
 .cells(3,1).Value2 = ''
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
  oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,31))
  oRange.Merge
 ENDFOR  
 
ENDWITH 

oExcel.Range(oExcel.Cells(5,04), oExcel.Cells(5,12)).Merge 
oExcel.Cells(5,04).Value='Представлено ЛПУ'
oExcel.Cells(5,04).HorizontalAlignment=-4108

oExcel.Range(oExcel.Cells(5,13), oExcel.Cells(5,21)).Merge 
oExcel.Cells(5,13).Value='После МЭК'
oExcel.Cells(5,13).HorizontalAlignment=-4108

oExcel.Range(oExcel.Cells(5,22), oExcel.Cells(5,31)).Merge 
oExcel.Cells(5,22).Value='После МЭЭ'
oExcel.Cells(5,22).HorizontalAlignment=-4108

oExcel.Range(oExcel.Cells(6,04), oExcel.Cells(6,06)).Merge 
oExcel.Range(oExcel.Cells(6,07), oExcel.Cells(6,09)).Merge 
oExcel.Range(oExcel.Cells(6,10), oExcel.Cells(6,12)).Merge 
oExcel.Range(oExcel.Cells(6,13), oExcel.Cells(6,15)).Merge 
oExcel.Range(oExcel.Cells(6,16), oExcel.Cells(6,18)).Merge 
oExcel.Range(oExcel.Cells(6,19), oExcel.Cells(6,21)).Merge 
oExcel.Range(oExcel.Cells(6,22), oExcel.Cells(6,24)).Merge 
oExcel.Range(oExcel.Cells(6,25), oExcel.Cells(6,27)).Merge 
oExcel.Range(oExcel.Cells(6,28), oExcel.Cells(6,31)).Merge 

oExcel.Cells(6,04).Value2 = 'Амб.-пол.'
oExcel.Cells(6,04).HorizontalAlignment=-4108
oExcel.Cells(6,07).Value2 = 'Дневной стационар'
oExcel.Cells(6,07).HorizontalAlignment=-4108
oExcel.Cells(6,10).Value2 = 'Стационар'
oExcel.Cells(6,10).HorizontalAlignment=-4108

oExcel.Cells(6,13).Value2 = 'Амб.-пол.'
oExcel.Cells(6,13).HorizontalAlignment=-4108
oExcel.Cells(6,16).Value2 = 'Дневной стационар'
oExcel.Cells(6,16).HorizontalAlignment=-4108
oExcel.Cells(6,19).Value2 = 'Стационар'
oExcel.Cells(6,19).HorizontalAlignment=-4108

oExcel.Cells(6,22).Value2 = 'Амб.-пол.'
oExcel.Cells(6,22).HorizontalAlignment=-4108
oExcel.Cells(6,25).Value2 = 'Дневной стационар'
oExcel.Cells(6,25).HorizontalAlignment=-4108
oExcel.Cells(6,28).Value2 = 'Стационар'
oExcel.Cells(6,28).HorizontalAlignment=-4108

oExcel.Cells(7,04).Value2 = 'Счетов'
oExcel.Cells(7,05).Value2 = 'Услуг'
oExcel.Cells(7,06).Value2 = 'Сумма'
oExcel.Cells(7,07).Value2 = 'Счетов'
oExcel.Cells(7,08).Value2 = 'К/дн.'
oExcel.Cells(7,09).Value2 = 'Сумма'
oExcel.Cells(7,10).Value2 = 'Счетов'
oExcel.Cells(7,11).Value2 = 'МЭСов'
oExcel.Cells(7,12).Value2 = 'Сумма'

oExcel.Cells(7,13).Value2 = 'Счетов'
oExcel.Cells(7,14).Value2 = 'Услуг'
oExcel.Cells(7,15).Value2 = 'Сумма'
oExcel.Cells(7,16).Value2 = 'Счетов'
oExcel.Cells(7,17).Value2 = 'К/дн.'
oExcel.Cells(7,18).Value2 = 'Сумма'
oExcel.Cells(7,19).Value2 = 'Счетов'
oExcel.Cells(7,20).Value2 = 'МЭСов'
oExcel.Cells(7,21).Value2 = 'Сумма'

oExcel.Cells(7,22).Value2 = 'Счетов'
oExcel.Cells(7,23).Value2 = 'Услуг'
oExcel.Cells(7,24).Value2 = 'Сумма'
oExcel.Cells(7,25).Value2 = 'Счетов'
oExcel.Cells(7,26).Value2 = 'К/дн.'
oExcel.Cells(7,27).Value2 = 'Сумма'
oExcel.Cells(7,28).Value2 = 'Счетов'
oExcel.Cells(7,29).Value2 = 'МЭСов'
oExcel.Cells(7,30).Value2 = 'К/дн.'
oExcel.Cells(7,31).Value2 = 'Сумма'

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

FOR ncol=1 TO 31
 WITH oExcel.Sheets(1)
  .Cells(m.fstring,ncol).Font.Size = 8
  .Cells(m.fstring,ncol).Font.Bold = .F.
  .Cells(m.fstring,ncol).WrapText = .t.
  .Cells(m.fstring,ncol).HorizontalAlignment = -4108
  .Cells(m.fstring,ncol).VerticalAlignment = -4108
 ENDWITH 
NEXT 

m.fstring = m.fstring + 1
FOR ncolumn=1 TO 31
 WITH oExcel.Sheets(1)
  .cells(m.fstring,ncolumn).Value2 = STR(ncolumn,2)
  .cells(m.fstring,ncolumn).HorizontalAlignment = -4108
 ENDWITH 
NEXT 

oExcel.Range(oExcel.Cells(m.fstring,1), oExcel.Cells(m.fstring,31)).NumberFormat='@'

m.fstring = m.fstring + 1

m.mcod = mcod
m.lpuname = sprlpu.name
*oExcel.Sheets(1).cells(m.fstring,2).Value2 = m.lpuname
oRange = oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,31))
oRange.Merge
oRange.HorizontalAlignment = -4108
oRange.Interior.ColorIndex = 40

m.nnn = 0

m.spaz_amb = 0
m.susl_amb = 0
m.ssum_amb = 0
m.spaz_dst = 0
m.skd_dst  = 0
m.ssum_dst = 0
m.spaz_st  = 0
m.sms_st   = 0
m.ssum_st  = 0

m.spaz_ambmek = 0
m.susl_ambmek = 0
m.ssum_ambmek = 0
m.spaz_dstmek = 0
m.skd_dstmek  = 0
m.ssum_dstmek = 0
m.spaz_stmek  = 0
m.sms_stmek   = 0
m.ssum_stmek  = 0

m.spaz_ambmee = 0
m.susl_ambmee = 0
m.ssum_ambmee = 0
m.spaz_dstmee = 0
m.skd_dstmee  = 0
m.ssum_dstmee = 0
m.spaz_stmee = 0
m.sms_stmee   = 0
m.skd_stmee   = 0
m.ssum_stmee  = 0

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

 CREATE CURSOR pazppl (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR pazdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazst (c_i c(25))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i
   
 CREATE CURSOR pazambmek (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazdstmek (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazstmek (c_i c(25))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR pazambmee (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazdstmee (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazstmee (c_i c(25))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR ambchkdmee (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR dstchkdmee (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR stchkdmee (c_i c(25))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 SELECT AisOms

 m.nnn = m.nnn+1
 m.fstring = m.fstring + 1
 m.lpuname = sprlpu.name

 m.paz_amb = 0
 m.usl_amb = 0
 m.sum_amb = 0
 m.paz_dst = 0
 m.kd_dst  = 0
 m.sum_dst = 0
 m.paz_st  = 0
 m.ms_st   = 0
 m.sum_st  = 0

 m.paz_ambmek = 0
 m.usl_ambmek = 0
 m.sum_ambmek = 0
 m.paz_dstmek = 0
 m.kd_dstmek  = 0
 m.sum_dstmek = 0
 m.paz_stmek  = 0
 m.ms_stmek   = 0
 m.sum_stmek  = 0

 m.paz_ambmee = 0
 m.usl_ambmee = 0
 m.sum_ambmee = 0
 m.paz_dstmee = 0
 m.kd_dstmee  = 0
 m.sum_dstmee = 0
 m.paz_stmee = 0
 m.ms_stmee   = 0
 m.kd_stmee   = 0
 m.sum_stmee  = 0

* m.spaz_amb = m.spaz_amb + krank
* m.susl_amb = m.susl_amb + usl_amb
* m.ssum_amb = m.ssum_amb + sum_kr
* m.spaz_dst = m.spaz_dst + paz_dst
* m.skd_dst  = m.skd_dst  + kd_dst
* m.ssum_dst = m.ssum_dst + sum_dst
* m.spaz_st  = m.spaz_st  + paz_st
* m.sms_st   = m.sms_st   + ms_st
* m.ssum_st  = m.ssum_st  + sum_st

 SELECT talon 
 SCAN 
  m.cod = cod
  m.c_i = c_i
  m.sn_pol = sn_pol
  m.tip = tip
  m.lIs02 = IIF(INLIST(m.cod,83010,83020,83030,83040,83050),.t.,.f.)

  IF IsUsl(m.cod)
   m.sum_amb = m.sum_amb + s_all
   m.ssum_amb = m.ssum_amb + s_all
   m.usl_amb = m.usl_amb + k_u
   m.susl_amb = m.susl_amb + k_u
   IF !SEEK(m.sn_pol, 'pazamb')
    INSERT INTO pazamb (sn_pol) VALUES (m.sn_pol)
    m.paz_amb = m.paz_amb + 1
    m.spaz_amb = m.spaz_amb + 1
   ENDIF 
  ENDIF 
   
  IF (IsMes(m.cod) OR IsVMP(m.cod))
   m.sum_st = m.sum_st + s_all
   m.ssum_st = m.ssum_st + s_all
*   m.ms_st = m.ms_st + IIF(m.lIs02, 0, 1)
*   m.sms_st = m.sms_st + IIF(m.lIs02, 0, 1)
   m.ms_st = m.ms_st + 1
   m.sms_st = m.sms_st + 1
   IF !SEEK(m.c_i, 'pazst')
    INSERT INTO pazst (c_i) VALUES (m.c_i)
    m.paz_st = m.paz_st + 1
    m.spaz_st = m.spaz_st + 1
   ENDIF 
  ENDIF 

  IF IsKd(m.cod)
   m.sum_dst = m.sum_dst + s_all
   m.ssum_dst = m.ssum_dst + s_all
*   m.kd_dst = m.kd_dst + 1
*   m.skd_dst = m.skd_dst + 1
   m.kd_dst = m.kd_dst + k_u
   m.skd_dst = m.skd_dst + k_u
   IF !SEEK(m.sn_pol, 'pazdst')
    INSERT INTO pazdst (sn_pol) VALUES (m.sn_pol)
    m.paz_dst = m.paz_dst + 1
    m.spaz_dst = m.spaz_dst + 1
   ENDIF 
  ENDIF 

  IF IsUsl(m.cod) AND EMPTY(error.c_err)
   m.sum_ambmek = m.sum_ambmek + s_all
   m.ssum_ambmek = m.ssum_ambmek + s_all
   m.usl_ambmek = m.usl_ambmek + k_u
   m.susl_ambmek = m.susl_ambmek + k_u
   IF !SEEK(m.sn_pol, 'pazambmek')
    INSERT INTO pazambmek (sn_pol) VALUES (m.sn_pol)
    m.paz_ambmek = m.paz_ambmek + 1
    m.spaz_ambmek = m.spaz_ambmek + 1
   ENDIF 
  ENDIF 
   
  IF (IsMes(m.cod) OR IsVMP(m.cod)) AND EMPTY(error.c_err)
   m.sum_stmek = m.sum_stmek + s_all
   m.ssum_stmek = m.ssum_stmek + s_all
*   m.ms_stmek = m.ms_stmek + IIF(m.lIs02, 0, 1)
*   m.sms_stmek = m.sms_stmek + IIF(m.lIs02, 0, 1)
   m.ms_stmek = m.ms_stmek + 1
   m.sms_stmek = m.sms_stmek + 1
   IF !SEEK(m.c_i, 'pazstmek')
    INSERT INTO pazstmek (c_i) VALUES (m.c_i)
    m.paz_stmek = m.paz_stmek + 1
    m.spaz_stmek = m.spaz_stmek + 1
   ENDIF 
  ENDIF 

  IF IsKd(m.cod) AND EMPTY(error.c_err)
   m.sum_dstmek = m.sum_dstmek + s_all
   m.ssum_dstmek = m.ssum_dstmek + s_all
*   m.kd_dstmek = m.kd_dstmek + 1
*   m.skd_dstmek = m.skd_dstmek + 1
   m.kd_dstmek = m.kd_dstmek + k_u
   m.skd_dstmek = m.skd_dstmek + k_u
   IF !SEEK(m.sn_pol, 'pazdstmek')
    INSERT INTO pazdstmek (sn_pol) VALUES (m.sn_pol)
    m.paz_dstmek = m.paz_dstmek + 1
    m.spaz_dstmek = m.spaz_dstmek + 1
   ENDIF 
  ENDIF 

  IF IsUsl(m.cod) AND EMPTY(error.c_err) AND ;
   (EMPTY(merror.e_cod) AND EMPTY(merror.e_ku) AND EMPTY(merror.e_tip))
   m.sum_ambmee = m.sum_ambmee + s_all
   m.ssum_ambmee = m.ssum_ambmee + s_all
   m.usl_ambmee = m.usl_ambmee + k_u
   m.susl_ambmee = m.susl_ambmee + k_u
   IF !SEEK(m.sn_pol, 'pazambmee')
    INSERT INTO pazambmee (sn_pol) VALUES (m.sn_pol)
    m.paz_ambmee = m.paz_ambmee + 1
    m.spaz_ambmee = m.spaz_ambmee + 1
   ENDIF 
  ENDIF 
   
  IF (IsMes(m.cod) OR IsVMP(m.cod)) AND EMPTY(error.c_err) AND ;
   (EMPTY(merror.e_cod) AND EMPTY(merror.e_ku) AND EMPTY(merror.e_tip))
   m.sum_stmee = m.sum_stmee + s_all
   m.ssum_stmee = m.ssum_stmee + s_all
*   m.ms_stmee = m.ms_stmee + IIF(m.lIs02, 0, 1)
*   m.sms_stmee = m.sms_stmee + IIF(m.lIs02, 0, 1)
   m.ms_stmee = m.ms_stmee + 1
   m.sms_stmee = m.sms_stmee + 1
   m.kd_stmee = m.kd_stmee + k_u
   m.skd_stmee = m.skd_stmee + k_u
   IF !SEEK(m.c_i, 'pazstmee')
    INSERT INTO pazstmee (c_i) VALUES (m.c_i)
    m.paz_stmee = m.paz_stmee + 1
    m.spaz_stmee = m.spaz_stmee + 1
   ENDIF 
  ENDIF 

  IF (IsUsl(m.cod)) AND EMPTY(error.c_err) AND ;
   (EMPTY(merror.e_cod) AND EMPTY(merror.e_ku) AND EMPTY(merror.e_tip))
   m.sum_dstmee = m.sum_dstmee + s_all
   m.ssum_dstmee = m.ssum_dstmee + s_all
   m.kd_dstmee = m.kd_dstmee + k_u
   m.skd_dstmee = m.skd_dstmee + k_u
   IF !SEEK(m.sn_pol, 'pazdstmee')
    INSERT INTO pazdstmee (sn_pol) VALUES (m.sn_pol)
    m.paz_dstmee = m.paz_dstmee + 1
    m.spaz_dstmee = m.spaz_dstmee + 1
   ENDIF 
  ENDIF 

 ENDSCAN 
 SET RELATION OFF INTO error 
 SET RELATION OFF INTO merror
 
 SELECT AisOms

 WITH oExcel.Sheets(1)
  .Cells(m.fstring,01).Value2 = m.nnn
  .Cells(m.fstring,02).Value2 = ALLTRIM(m.lpuname)
  .Cells(m.fstring,03).Value2 = m.mcod
  .Cells(m.fstring,04).Value2 = m.paz_amb
  .Cells(m.fstring,05).Value2 = m.usl_amb
  .Cells(m.fstring,06).Value2 = m.sum_amb
  .Cells(m.fstring,07).Value2 = m.paz_dst
  .Cells(m.fstring,08).Value2 = m.kd_dst
  .Cells(m.fstring,09).Value2 = m.sum_dst
  .Cells(m.fstring,10).Value2 = m.paz_st
  .Cells(m.fstring,11).Value2 = m.ms_st
  .Cells(m.fstring,12).Value2 = m.sum_st

  .Cells(m.fstring,13).Value2 = m.paz_ambmek
  .Cells(m.fstring,14).Value2 = m.usl_ambmek
  .Cells(m.fstring,15).Value2 = m.sum_ambmek
  .Cells(m.fstring,16).Value2 = m.paz_dstmek
  .Cells(m.fstring,17).Value2 = m.kd_dstmek
  .Cells(m.fstring,18).Value2 = m.sum_dstmek
  .Cells(m.fstring,19).Value2 = m.paz_stmek
  .Cells(m.fstring,20).Value2 = m.ms_stmek
  .Cells(m.fstring,21).Value2 = m.sum_stmek

  .Cells(m.fstring,22).Value2 = m.paz_ambmee
  .Cells(m.fstring,23).Value2 = m.usl_ambmee
  .Cells(m.fstring,24).Value2 = m.sum_ambmee
  .Cells(m.fstring,25).Value2 = m.paz_dstmee
  .Cells(m.fstring,26).Value2 = m.kd_dstmee
  .Cells(m.fstring,27).Value2 = m.sum_dstmee
  .Cells(m.fstring,28).Value2 = m.paz_stmee
  .Cells(m.fstring,29).Value2 = m.ms_stmee
  .Cells(m.fstring,30).Value2 = m.kd_stmee
  .Cells(m.fstring,31).Value2 = m.sum_stmee

 ENDWITH 

 USE IN pazamb
 USE IN pazdst
 USE IN pazst
 USE IN pazppl
 USE IN pazambmek
 USE IN pazdstmek
 USE IN pazstmek
 USE IN ambchkdmee
 USE IN dstchkdmee
 USE IN stchkdmee

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
 .Cells(m.fstring,04).Value2 = m.spaz_amb
 .Cells(m.fstring,05).Value2 = m.susl_amb
 .Cells(m.fstring,06).Value2 = m.ssum_amb
 .Cells(m.fstring,07).Value2 = m.spaz_dst
 .Cells(m.fstring,08).Value2 = m.skd_dst
 .Cells(m.fstring,09).Value2 = m.ssum_dst
 .Cells(m.fstring,10).Value2 = m.spaz_st
 .Cells(m.fstring,11).Value2 = m.sms_st
 .Cells(m.fstring,12).Value2 = m.ssum_st

 .Cells(m.fstring,13).Value2 = m.spaz_ambmek
 .Cells(m.fstring,14).Value2 = m.susl_ambmek
 .Cells(m.fstring,15).Value2 = m.ssum_ambmek
 .Cells(m.fstring,16).Value2 = m.spaz_dstmek
 .Cells(m.fstring,17).Value2 = m.skd_dstmek
 .Cells(m.fstring,18).Value2 = m.ssum_dstmek
 .Cells(m.fstring,19).Value2 = m.spaz_stmek
 .Cells(m.fstring,20).Value2 = m.sms_stmek
 .Cells(m.fstring,21).Value2 = m.ssum_stmek
 .Cells(m.fstring,22).Value2 = m.spaz_ambmee
 .Cells(m.fstring,23).Value2 = m.susl_ambmee
 .Cells(m.fstring,24).Value2 = m.ssum_ambmee
 .Cells(m.fstring,25).Value2 = m.spaz_dstmee
 .Cells(m.fstring,26).Value2 = m.skd_dstmee
 .Cells(m.fstring,27).Value2 = m.ssum_dstmee
 .Cells(m.fstring,28).Value2 = m.spaz_stmee
 .Cells(m.fstring,29).Value2 = m.sms_stmee
 .Cells(m.fstring,30).Value2 = m.skd_stmee
 .Cells(m.fstring,31).Value2 = m.ssum_stmee
ENDWITH 

IF fso.FileExists(pDir+'\Отчет для Поветко.xls')
 fso.DeleteFile(pDir+'\Отчет для Поветко.xls')
ENDIF 

oBook.SaveAs(BookName,18)

oExcel.Visible = .T.

SET RELATION OFF INTO sprlpu
USE IN aisoms
USE IN sprlpu

WAIT CLEAR 

RETURN 
