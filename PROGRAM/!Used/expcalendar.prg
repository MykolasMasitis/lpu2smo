PROCEDURE ExpCalendar

 IF !fso.FileExists(pbase+'\ExpCalendar.dbf')
  WAIT "ÑÎÇÄÀÅÒÑß ÑÒÐÓÊÒÓÐÀ ÊÀËÅÍÄÀÐß..." WINDOW NOWAIT 
  =OpenFile(pcommon+'\sprlpuxx', 'sprlpu', 'shar')
  CREATE TABLE &pbase\ExpCalendar (lpuid n(6), mcod c(7), p201201 l, p201202 l, p201203 l, p201204 l, p201205 l, ;
   p201206 l, p201207 l, p201208 l, p201209 l, p201210 l, p201211 l, p201212 l)
  INDEX on lpuid TAG lpuid
  INDEX on mcod TAG mcod 
  USE 
  =OpenFile(pbase+'\expcalendar', 'expcal', 'shar')
  
  SELECT sprlpu
  SCAN 
   m.lpuid = lpu_id
   m.mcod = mcod
   INSERT INTO expcal (lpuid, mcod) VALUES (m.lpuid, m.mcod)
  ENDSCAN 
  USE 
  USE IN expcal
  WAIT CLEAR 
 
  =OpenFile(pbase+'\expcalendar', 'expcal', 'shar', 'lpuid')

  FOR nmonth=1 TO 12
   m.pgperiod = STR(YEAR(m.tdat2),4)+PADL(nmonth,2,'0')
   m.IsExp = IsExp(m.pgperiod)
  NEXT 
 ENDIF 
 
 IF m.IsNotePad = .F.
  DO FORM ViewCalendar
 ELSE 
  DO FORM ViewCalendar600
 ENDIF 

FUNCTION IsExp(_pgperiod)
 m.lcPgPeriod = _pgperiod

 IF !fso.FileExists(pBase+'\'+lcPgPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile("&pBase\&lcPgPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+lcPgPeriod+'\'+'nsi'+'\sprlpuxx', 'sprlp', 'SHARED') > 0
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 SCAN 
  m.mcod = mcod
  m.lpuid =  lpuid

  WAIT m.mcod WINDOW NOWAIT 
  
  IF !fso.FolderExists(pbase+'\'+lcPgPeriod+'\'+m.mcod)
   MESSAGEBOX(CHR(13)+CHR(10)+'ÄÈÐÅÊÒÎÐÈß '+m.mcod+' ÎÒÑÓÒÑÒÂÓÅÒ!'+CHR(13)+CHR(10),0+48,'')
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\talon.dbf')
   MESSAGEBOX(CHR(13)+CHR(10)+'ÔÀÉË TALON.DBF ÎÒÑÓÒÑÒÂÓÅÒ!'+CHR(13)+CHR(10),0+48, m.mcod)
   LOOP 
  ENDIF 
  
  IF OpenFile(pbase+'\'+lcPgPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   LOOP 
  ENDIF 
  
  m.WasExp = .f.
  SELECT talon 
  SCAN 
   IF !EMPTY(err_mee)
    m.WasExp = .t.
    EXIT 
   ENDIF 
   WAIT CLEAR 
  ENDSCAN 
  USE IN talon 
  
  IF m.WasExp==.t.
   IF SEEK(m.lpuid, 'expcal')
    fldname = 'p'+ m.lcPgPeriod
    UPDATE expcal SET &fldname = .t. WHERE expcal.lpuid = m.lpuid
   ENDIF 
  ENDIF 

  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE 
 USE IN sprlp

RETURN

