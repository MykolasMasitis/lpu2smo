PROCEDURE RepExp7
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
  'ОТЧЕТ ПО ЦЕЛЕВЫМ ЭКСПЕРТИЗАМ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 PUBLIC m.minperiod, m.maxperiod
 
 CREATE CURSOR curexps (period c(7), lpuid i(4), mcod c(7), lpuname c(120), ischked l)

 CREATE CURSOR curperiod (period c(6))
 INDEX on period TAG period
 SET ORDER TO period

 FOR lnmonth=1 TO tMonth
  m.lcperiod = STR(tYear,4)+PADL(lnmonth,2,'0')
  IF !SEEK(m.lcperiod, 'curperiod')
   INSERT INTO curperiod (period) VALUES (m.lcperiod)
  ENDIF 
 ENDFOR 
 
 SELECT curperiod
 GO TOP 
 m.minperiod = period
 GO BOTTOM 
 m.maxperiod = period

 FOR lnmonth=1 TO tMonth
  m.lcperiod = STR(tYear,4)+PADL(lnmonth,2,'0')
  m.lpath = pbase+'\'+m.lcperiod
  IF !fso.FolderExists(m.lpath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lpath+'\aisoms.dbf')
   LOOP 
  ENDIF 

  WAIT m.lcperiod+'...' WINDOW NOWAIT 
  =selexpsone(m.lpath)
  WAIT CLEAR 
 ENDFOR 
 
 IF fso.FileExists(m.pbase+'\'+m.gcperiod+'\nsi'+'\sprlpuxx'+'.dbf')
  IF OpenFile(m.pbase+'\'+m.gcperiod+'\nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
  ELSE 
   SELECT curexps
   SET RELATION TO lpuid INTO sprlpu
   SCAN 
    REPLACE mcod WITH sprlpu.mcod
    REPLACE lpuname WITH sprlpu.fullname
   ENDSCAN 
   SET RELATION OFF INTO sprlpu
   USE IN sprlpu 
  ENDIF 
 ENDIF 


 PUBLIC oExcel AS Excel.Application
 WAIT "Запуск MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 WAIT "ФОРМИРОВАНИЕ ОТЧЕТА..." WINDOW NOWAIT 
 
 CREATE CURSOR curcur (lpuid n(4), mcod c(7), lpuname c(120), mustexps n(6), doneexps n(6))
 INDEX on lpuid TAG lpuid 
 INDEX on mcod TAG mcod 
 SET ORDER TO lpuid
 
 SELECT curexps 
 SCAN 
  m.lpuid   = lpuid 
  m.mcod    = mcod
  m.lpuname = ALLTRIM(lpuname)
  m.ischked = ischked
  IF !SEEK(m.lpuid, 'curcur')
   INSERT INTO curcur (lpuid, mcod, lpuname, mustexps, doneexps) VALUES (m.lpuid, m.mcod, m.lpuname,1,IIF(m.ischked,1,0))
  ELSE 
   m.omust = curcur.mustexps
   m.nmust = m.omust + 1
   m.odone = curcur.doneexps
   m.ndone = m.odone + IIF(m.ischked,1,0)
   UPDATE curcur SET mustexps=m.nmust, doneexps=m.ndone WHERE lpuid=m.lpuid
  ENDIF 
 ENDSCAN 
 
 oExcel.UseSystemSeparators = .F.
 oExcel.DecimalSeparator = '.'

 oexcel.ReferenceStyle= -4150  && xlR1C1
 
 oexcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add

 oSheet = oBook.WorkSheets(1)
 oSheet.Select
 oSheet.Name = 'Сводная'

 =HeadOfTheList()

 nRow = 8
 nRnn = 1

 SELECT curcur 
 SET ORDER TO mcod 
 SCAN
 
  oExcel.Cells(nRow,01).HorizontalAlignment = -4131
  oExcel.Cells(nRow,02).HorizontalAlignment = -4131
  oExcel.Cells(nRow,03).HorizontalAlignment = -4131
  oExcel.Cells(nRow,04).HorizontalAlignment = -4131

  oExcel.Rows(nRow).RowHeight = 15
  oExcel.Rows(nRow).VerticalAlignment = -4108

  oExcel.Cells(nRow,1).Value  = nRnn
  oExcel.Cells(nRow,2).Value  = STR(lpuid,4)
  oExcel.Cells(nRow,3).Value  = mcod
  oExcel.Cells(nRow,4).Value  = lpuname
  oExcel.Cells(nRow,5).Value  = STR(mustexps,6)
  oExcel.Cells(nRow,6).Value  = STR(doneexps,6)
  oExcel.Cells(nRow,7).Value  = STR(mustexps-doneexps,6)
  
  nRnn = nRnn + 1
  nRow = nRow + 1
 ENDSCAN 

 USE IN curcur 
 
 SELECT curperiod
 IF RECCOUNT()>0
  GO TOP 
  SCAN 
  
   m.period = period
   
   WAIT m.period WINDOW NOWAIT 

   CREATE CURSOR curcur (lpuid n(4), mcod c(7), lpuname c(120), mustexps n(6), doneexps n(6))
   INDEX on lpuid TAG lpuid 
   INDEX on mcod TAG mcod 
   SET ORDER TO lpuid
 
   SELECT curexps 
   SCAN 
    IF period!=m.period
     LOOP 
    ENDIF 
    
    m.lpuid   = lpuid 
    m.mcod    = mcod
    m.lpuname = ALLTRIM(lpuname)
    m.ischked = ischked
    IF !SEEK(m.lpuid, 'curcur')
     INSERT INTO curcur (lpuid, mcod, lpuname, mustexps, doneexps) VALUES (m.lpuid, m.mcod, m.lpuname,1,IIF(m.ischked,1,0))
    ELSE 
     m.omust = curcur.mustexps
     m.nmust = m.omust + 1
     m.odone = curcur.doneexps
     m.ndone = m.odone + IIF(m.ischked,1,0)
     UPDATE curcur SET mustexps=m.nmust, doneexps=m.ndone WHERE lpuid=m.lpuid
    ENDIF 
   ENDSCAN 
   
   oSheet = oBook.WorkSheets.Add(,oexcel.ActiveSheet)
   oSheet.Name = m.period
   
   SELECT curcur 
   
   =HeadOfTheList(m.period)
   =BodyOfTheList(m.period)

   nRow = 8
   nRnn = 1
   
   USE IN curcur 
   SELECT curperiod 
   WAIT CLEAR 
  ENDSCAN 
 ENDIF 
 USE IN curexps

 WAIT CLEAR 

 BookName = 'RepExp7'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(pmee+'\'+BookName+'.xls')
  fso.DeleteFile(pmee+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pmee+'\'+BookName,18)
 oExcel.Visible = .T.

RETURN 

FUNCTION selexpsone(m.lpath)
 PRIVATE m.llcpath
 m.llcpath = m.lpath
 IF OpenFile(m.llcpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 SCAN 
  m.lpuid = lpuid
  m.mcod = mcod
  IF INT(VAL(SUBSTR(mcod,3,2)))<41
   LOOP 
  ENDIF 
  IF !fso.FolderExists(m.llcpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\nsi'+'\tarifn'+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.llcpath+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\nsi'+'\tarifn', 'tarif', 'shar', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'recid')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 

 
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO error ADDITIVE 
  SET RELATION TO recid INTO merror ADDITIVE 
  SET RELATION TO cod INTO tarif ADDITIVE 
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   m.cod = cod
   IF !IsMes(m.cod)
    LOOP 
   ENDIF 

   m.k_u = k_u
   m.n_kd = tarif.n_kd
   m.koeff = ROUND(m.k_u/m.n_kd,0)
   
   IF BETWEEN(m.koeff,0.5,1.5)
    LOOP 
   ENDIF 
   
   m.ischked = IIF(EMPTY(merror.recid), .f., .t.)

   INSERT INTO curexps (period,lpuid,mcod,ischked) VALUES ;
    (m.lcperiod,m.lpuid,m.mcod,m.ischked) 

  ENDSCAN 
  SET RELATION OFF INTO merror
  SET RELATION OFF INTO error
  SET RELATION OFF INTO people
  SET RELATION OFF INTO tarif
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('error')
   USE IN error
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
 
  SELECT aisoms

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 

FUNCTION HeadOfTheList(llperiod)
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,7))
 oRange.Merge
 oExcel.Cells(1,1).Value='ОТЧЕТ ПО ЦЕЛЕВОЙ ЭКСПЕРТИЗЕ'
 oExcel.Cells(1,1).HorizontalAlignment = -4108
 oExcel.Cells(1,1).Font.Size = 12
 oExcel.Cells(1,1).Font.Bold = .F.
 oExcel.Cells(1,1).Font.Italic = .T.
 oExcel.Rows(1).RowHeight = 30
 oExcel.Rows(1).VerticalAlignment = -4108
 
 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,3))
 oRange.Merge
 oExcel.Cells(2,1).Value = 'По периоду:'
 oRange = oExcel.Range(oExcel.Cells(2,4), oExcel.Cells(2,7))
 oRange.Merge
 oExcel.Cells(2,4).Value = IIF(EMPTY(llperiod), 'Сводная по всем периодам', llperiod)
 oExcel.Cells(2,4).Font.Size = 12
 oExcel.Cells(2,4).Font.Bold = .F.
 oExcel.Cells(2,4).Font.Italic = .T.
 oExcel.Rows(2).RowHeight = 30
 oExcel.Rows(2).VerticalAlignment = -4108

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,3))
 oRange.Merge
 oExcel.Cells(3,1).Value = 'За период:'
 oExcel.Cells(3,4).Value = 'с: '+m.minperiod+' по '+m.maxperiod
 
 oExcel.Columns(1).ColumnWidth = 3
 oExcel.Columns(2).ColumnWidth = 4
 oExcel.Columns(3).ColumnWidth = 7
 oExcel.Columns(4).ColumnWidth = 85
 oExcel.Columns(5).ColumnWidth = 5
 oExcel.Columns(6).ColumnWidth = 5
 oExcel.Columns(7).ColumnWidth = 5
 
 oExcel.Cells(7,1).Value  = '№ п/п'
 oExcel.Cells(7,2).Value  = 'lpuid'
 oExcel.Cells(7,3).Value  = 'mcod'
 oExcel.Cells(7,4).Value  = 'Наименование ЛПУ'
 oExcel.Cells(7,5).Value  = 'Подлежит экспертизе'
 oExcel.Cells(7,6).Value  = 'Проведено экспертиз'
 oExcel.Cells(7,7).Value  = 'Провести еще'

 oExcel.Rows(7).RowHeight=40
 oExcel.Rows(7).VerticalAlignment = -4108
RETURN 

FUNCTION BodyOfTheList(lcperiod)
 nRow = 8
 nRnn = 1

 SCAN
  
  oExcel.Cells(nRow,01).HorizontalAlignment = -4131
  oExcel.Cells(nRow,02).HorizontalAlignment = -4131
  oExcel.Cells(nRow,03).HorizontalAlignment = -4131
  oExcel.Cells(nRow,04).HorizontalAlignment = -4131

  oExcel.Rows(nRow).RowHeight=20
  oExcel.Rows(nRow).VerticalAlignment = -4108

  oExcel.Cells(nRow,1).Value  = nRnn
  oExcel.Cells(nRow,2).Value  = STR(lpuid,4)
  oExcel.Cells(nRow,3).Value  = mcod
  oExcel.Cells(nRow,4).Value  = lpuname
  oExcel.Cells(nRow,5).Value  = STR(mustexps,6)
  oExcel.Cells(nRow,6).Value  = STR(doneexps,6)
  oExcel.Cells(nRow,7).Value  = STR(mustexps-doneexps,6)
  
  nRnn = nRnn + 1
  nRow = nRow + 1
 ENDSCAN 

RETURN 