PROCEDURE FormSh2
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ ÏÎ ÐÎÄÄÎÌÀÌ?'+CHR(13)+CHR(10),4+32,'ÔÎÐÌÀ Ø-2')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 =CrFiles()
 
 IF OpenFile(pmee+'\RODD\sprlpu', 'sprlp', 'shar')>0
  IF USED('sprlp')
   USE IN sprlp
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pmee+'\RODD\sprcod', 'sprcod', 'shar')>0
  IF USED('sprlp')
   USE IN sprlp
  ENDIF 
  IF USED('sprcod')
   USE IN sprcod
  ENDIF 
  RETURN 
 ENDIF 
 
 PUBLIC oExcel AS Excel.Application
 WAIT "Çàïóñê MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add

 SELECT sprlp
 SCAN 
  m.lpuid = lpuid
  m.mcod  = mcod

  TRY 
   oSheet = oBook.WorkSheets(m.mcod)
  CATCH 
   IF oexcel.ActiveSheet.name!='Ëèñò1'
    oSheet = oBook.WorkSheets.Add(,oexcel.ActiveSheet)
   ELSE 
    oSheet = oexcel.ActiveSheet
   ENDIF 
   oSheet.Name = m.mcod
  ENDTRY 
  oExcel.Columns(1).NumberFormat = '@'
  oExcel.Columns(1).ColumnWidth  = 6
  oExcel.Columns(2).NumberFormat = '@'
  oExcel.Columns(2).ColumnWidth  = 70
  oExcel.Columns(3).NumberFormat = '0'
  oExcel.Columns(3).ColumnWidth  = 8
  oExcel.Columns(4).NumberFormat = '0'
  oExcel.Columns(4).ColumnWidth  = 8
  oExcel.Columns(5).NumberFormat = '0'
  oExcel.Columns(5).ColumnWidth  = 8

  crsname = 'l'+m.mcod
  CREATE CURSOR &crsname (cod n(6), name c(250), p2012 n(5), p2013 n(5), p2014 n(5))
  INDEX on cod TAG cod 
  SET ORDER TO cod
  SELECT sprcod
  SCAN 
   m.cod = cod
   m.name = name 
   INSERT INTO &crsname FROM MEMVAR 
  ENDSCAN 
  SELECT &crsname

  lcyear = '2012'
  FOR nperiod = 1 TO 6
   m.tcperiod = lcyear + PADL(nperiod,2,'0')
   IF !fso.FolderExists(pbase+'\'+m.tcperiod)
    LOOP 
   ENDIF 
   IF !fso.FolderExists(pbase+'\'+m.tcperiod+'\nsi')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.tcperiod+'\nsi\sprlpuxx.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.tcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
    LOOP 
   ENDIF
   IF !SEEK(m.lpuid, 'sprlpu')
    USE IN sprlpu 
    LOOP 
   ENDIF 
   m.lcmcod = sprlpu.mcod
   USE IN sprlpu 
   IF !fso.FolderExists(pbase+'\'+m.tcperiod+'\'+m.lcmcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.tcperiod+'\'+m.lcmcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.tcperiod+'\'+m.lcmcod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT talon 
   SCAN 
    m.cod = cod 
    IF SEEK(m.cod, '&crsname')
     m.ok_u = &crsname..p2012
     m.nk_u = m.ok_u + 1
     REPLACE p2012 WITH m.nk_u IN &crsname
    ENDIF 
   ENDSCAN 
   USE IN talon 
   
  ENDFOR 
  
  lcyear = '2013'
  FOR nperiod = 1 TO 6
   m.tcperiod = lcyear + PADL(nperiod,2,'0')
   IF !fso.FolderExists(pbase+'\'+m.tcperiod)
    LOOP 
   ENDIF 
   IF !fso.FolderExists(pbase+'\'+m.tcperiod+'\nsi')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.tcperiod+'\nsi\sprlpuxx.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.tcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
    LOOP 
   ENDIF
   IF !SEEK(m.lpuid, 'sprlpu')
    USE IN sprlpu 
    LOOP 
   ENDIF 
   m.lcmcod = sprlpu.mcod
   USE IN sprlpu 
   IF !fso.FolderExists(pbase+'\'+m.tcperiod+'\'+m.lcmcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.tcperiod+'\'+m.lcmcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.tcperiod+'\'+m.lcmcod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT talon 
   SCAN 
    m.cod = cod 
    IF SEEK(m.cod, '&crsname')
     m.ok_u = &crsname..p2013
     m.nk_u = m.ok_u + 1
     REPLACE p2013 WITH m.nk_u IN &crsname
    ENDIF 
   ENDSCAN 
   USE IN talon 
   
  ENDFOR 

  lcyear = '2014'
  FOR nperiod = 1 TO 6
   m.tcperiod = lcyear + PADL(nperiod,2,'0')
   IF !fso.FolderExists(pbase+'\'+m.tcperiod)
    LOOP 
   ENDIF 
   IF !fso.FolderExists(pbase+'\'+m.tcperiod+'\nsi')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.tcperiod+'\nsi\sprlpuxx.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.tcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
    LOOP 
   ENDIF
   IF !SEEK(m.lpuid, 'sprlpu')
    USE IN sprlpu 
    LOOP 
   ENDIF 
   m.lcmcod = sprlpu.mcod
   USE IN sprlpu 
   IF !fso.FolderExists(pbase+'\'+m.tcperiod+'\'+m.lcmcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.tcperiod+'\'+m.lcmcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.tcperiod+'\'+m.lcmcod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT talon 
   SCAN 
    m.cod = cod 
    IF SEEK(m.cod, '&crsname')
     m.ok_u = &crsname..p2014
     m.nk_u = m.ok_u + 1
     REPLACE p2014 WITH m.nk_u IN &crsname
    ENDIF 
   ENDSCAN 
   USE IN talon 
   
  ENDFOR 

  SELECT &crsname
  COPY TO pmee+'\RODD\'+crsname

  WITH oSheet
   .Cells(1,1) = 'Êîä ÌÝÑ'
   .Cells(1,2) = 'Íàèìåíîâàíèå ÌÝÑ'
   .Cells(1,3) = '1 ïîëóãîäèå 2012'
   .Cells(1,4) = '1 ïîëóãîäèå 2013'
   .Cells(1,5) = '1 ïîëóãîäèå 2014'
  ENDWITH 

  m.n = 3

  SCAN 
   WITH oSheet
   .Cells(n,1) = PADL(cod,6,'0')
   .Cells(n,2) = name
   .Cells(n,3) = p2012
   .Cells(n,4) = p2013
   .Cells(n,5) = p2014
   ENDWITH 
   m.n = m.n + 1
  ENDSCAN 
  USE 
  
  SELECT sprlp
 ENDSCAN
 
 USE IN sprlp
 USE IN sprcod
 
 IF fso.FileExists(pmee+'\RODD\rodd.xls')
  fso.DeleteFile(pmee+'\RODD\rodd.xls')
 ENDIF 
 BookName = pmee+'\RODD\rodd'
 oBook.SaveAs(BookName,18)
 oExcel.Visible = .t.

RETURN 

FUNCTION CrFiles
 IF !fso.FolderExists(pmee+'\rodd')
  fso.CreateFolder(pmee+'\RODD')
 ENDIF 
 IF !fso.FileExists(pmee+'\RODD\sprlpu.dbf')
  CREATE TABLE &pmee\rodd\sprlpu (lpuid n(4), mcod c(7))
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1989,'0341003')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1992,'0343007') && ðåîðãàíèçîâàíà
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1990,'0341008')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1928,'0343015')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (2285,'0343020')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1905,'0343029')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (2858,'0343068')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1993,'0343070')
  INSERT INTO sprlpu (lpuid,mcod) VALUES (1991,'0343072') && ðåîðãàíèçîâàíà
  USE 
 ENDIF 
 IF !fso.FileExists(pmee+'\RODD\sprcod.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
  ENDIF 
  CREATE TABLE &pmee\rodd\sprcod (cod n(6), name c(250))
  INSERT INTO sprcod (cod) VALUES (76580)
  INSERT INTO sprcod (cod) VALUES (76590)
  INSERT INTO sprcod (cod) VALUES (76600)
  INSERT INTO sprcod (cod) VALUES (76610)
  INSERT INTO sprcod (cod) VALUES (76620)
  INSERT INTO sprcod (cod) VALUES (76630)
  INSERT INTO sprcod (cod) VALUES (76640)
  INSERT INTO sprcod (cod) VALUES (76650)
  INSERT INTO sprcod (cod) VALUES (76660)
  INSERT INTO sprcod (cod) VALUES (76670)
  INSERT INTO sprcod (cod) VALUES (76680)
  INSERT INTO sprcod (cod) VALUES (76690)
  INSERT INTO sprcod (cod) VALUES (76700)
  INSERT INTO sprcod (cod) VALUES (76710)
  INSERT INTO sprcod (cod) VALUES (76720)
  INSERT INTO sprcod (cod) VALUES (76730)
  INSERT INTO sprcod (cod) VALUES (76740)
  INSERT INTO sprcod (cod) VALUES (76750)
  INSERT INTO sprcod (cod) VALUES (76760)
  INSERT INTO sprcod (cod) VALUES (76770)
  INSERT INTO sprcod (cod) VALUES (76780)
  INSERT INTO sprcod (cod) VALUES (76790)
  INSERT INTO sprcod (cod) VALUES (76800)
  INSERT INTO sprcod (cod) VALUES (76810)
  INSERT INTO sprcod (cod) VALUES (76820)
  INSERT INTO sprcod (cod) VALUES (76830)
  INSERT INTO sprcod (cod) VALUES (76840)
  INSERT INTO sprcod (cod) VALUES (76850)
  INSERT INTO sprcod (cod) VALUES (76860)
  INSERT INTO sprcod (cod) VALUES (76870)
  INSERT INTO sprcod (cod) VALUES (76880)
  INSERT INTO sprcod (cod) VALUES (76890)
  INSERT INTO sprcod (cod) VALUES (76900)
  INSERT INTO sprcod (cod) VALUES (76910)
  INSERT INTO sprcod (cod) VALUES (76920)
  IF USED('tarif')
   SELECT sprcod 
   SET RELATION TO cod INTO tarif
   REPLACE ALL name WITH tarif.name
   USE IN tarif 
  ENDIF 
  USE IN sprcod
 ENDIF 
RETURN 