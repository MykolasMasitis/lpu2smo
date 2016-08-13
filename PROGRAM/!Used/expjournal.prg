PROCEDURE ExpJournal

 IF !fso.FileExists(pbase+'\ExpJournal.dbf')
  WAIT "ÑÎÇÄÀÅÒÑß ÑÒÐÓÊÒÓÐÀ ÆÓÐÍÀËÀ ÝÊÑÏÅÐÒÈÇ..." WINDOW NOWAIT 
  =OpenFile(pcommon+'\sprlpuxx', 'sprlpu', 'shar')
  CREATE TABLE &pbase\ExpJournal (lpuid n(6), mcod c(7), period c(6), e_period c(6),;
   pid i, sn_pol c(25), et c(1), aktsv c(50), daktsv d, aktss c(50), daktss d, sumexp n(11,2))
  INDEX on lpuid TAG lpuid
  INDEX on mcod TAG mcod 
  USE 
  =OpenFile(pbase+'\expjournal', 'expj', 'shar')
  
  SELECT sprlpu
  SCAN 
   m.lpuid = lpu_id
   m.mcod = mcod
   INSERT INTO expj (lpuid, mcod) VALUES (m.lpuid, m.mcod)
  ENDSCAN 
  USE 
  USE IN expj
  WAIT CLEAR 
 
*  =OpenFile(pbase+'\expcalendar', 'expcal', 'shar', 'lpuid')
  =OpenFile(pbase+'\expjournal', 'expj', 'shar', 'lpuid')

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
   IF SEEK(m.lpuid, 'expj')
    fldname = 'p'+ m.lcPgPeriod
    UPDATE expj SET &fldname = .t. WHERE expj.lpuid = m.lpuid
   ENDIF 
  ENDIF 

  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE 
 USE IN sprlp

RETURN

