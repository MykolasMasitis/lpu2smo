PROCEDURE ImpS6
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÈÌÏÎÐÒÈÐÎÂÀÒÜ ÄÀÍÍÛÅ?'+CHR(13)+CHR(10),4+32,'S6')=7
  RETURN 
 ENDIF 

 pUpdDir = 'd:\s6'

 SET DEFAULT TO (pUpdDir)
 csprfile = ''
 csprfile=GETFILE('dbf')
 IF EMPTY(csprfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÍÈ×ÅÃÎ ÍÅ ÂÛÁÐÀËÈ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 ospr = fso.GetFile(csprfile)
 IF LOWER(LEFT(ospr.name,3)) != 'lpu'
  MESSAGEBOX(CHR(13)+CHR(10)+'ÝÒÎ ÍÅ ÑÏÐÀÂÎ×ÍÈÊ LPU!'+CHR(13)+CHR(10),0+16,'lpu.dbf')
  RELEASE ospr 
  RETURN 
 ENDIF 

 IF OpenFile(csprfile, 'lpu', 'shar')>0
  RELEASE ospr 
  USE IN lpu 
  RETURN 
 ENDIF 
 
 m.lcperiod = RIGHT(ospr.ParentFolder.Path,6)
 m.curdir   = ospr.ParentFolder.Path
 m.mmy      = SUBSTR(m.lcperiod,5,2)+SUBSTR(m.lcperiod,4,1)
 
 IF INLIST(m.lcperiod,'201401','201402','201403','201404','201405','201406','201407',;
  '201408','201409','201410','201411','201412')
  IF MESSAGEBOX('ÂÛÁÐÀÍ ÏÅÐÈÎÄ '+m.lcperiod+CHR(13)+CHR(10)+'ÏÐÎÄÎËÆÈÒÜ?',4+32,'')=7
   RELEASE ospr 
   USE IN lpu 
   RETURN 
  ENDIF 
 
 ELSE 
  
  MESSAGEBOX('ÏÅÐÈÎÄ ÍÅ ÐÀÑÏÎÇÍÀÍ!'+CHR(13)+CHR(10),0+16,m.lcperiod)
  RELEASE ospr 
  USE IN lpu 
  RETURN 
  
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.lcperiod)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÂÛÁÐÀÍÍÛÉ ÏÅÐÈÎÄ Â LPU2SMO!'+CHR(13)+CHR(10),0+16,m.lcperiod)
  RELEASE ospr 
  USE IN lpu 
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË AISOMS Â ÂÛÁÐÀÍÍÎÌ ÏÅÐÈÎÄÅ!'+CHR(13)+CHR(10),0+16,m.lcperiod)
  RELEASE ospr 
  USE IN lpu 
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', "aisoms", "shar") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RELEASE ospr 
  USE IN lpu 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "lpu_id") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RELEASE ospr 
  USE IN lpu 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\profus', "profus", "shar", "cod") > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('profus')
   USE IN profus
  ENDIF 
  RELEASE ospr 
  USE IN lpu 
  RETURN 
 ENDIF 
 
 SELECT lpu 
 SCAN 
  m.lpu_id = lpu_id
  m.mcod   = mcod
  m.locdir = m.curdir + '\' + STR(m.lpu_id,4)
  IF !fso.FolderExists(m.locdir)
   MESSAGEBOX(m.locdir,0+64,'')
   LOOP 
  ENDIF 
  m.bfile = 'b'+m.mcod+'.'+m.mmy
  IF !fso.FileExists(m.locdir+'\'+m.bfile)
   MESSAGEBOX(m.bfile,0+64,'')
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   MESSAGEBOX(pbase+'\'+m.gcperiod+'\'+m.mcod,0+64,'')
   LOOP 
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.bfile)
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.bfile)
  ENDIF 
  
  fso.CopyFile(m.locdir+'\'+m.bfile, pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.bfile)
  
  WAIT m.mcod WINDOW NOWAIT 
  
  =AccReload(m.mcod, m.lpu_id, .f.)
  SELECT lpu 
  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar')>0
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'recid_lpu')>0
   IF USED('error')
    USE IN error
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid_lpu')>0
   IF USED('error')
    USE IN error
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.locdir+'\ctrl', 'ctrl', 'shar')>0
   IF USED('error')
    USE IN error
   ENDIF 
   IF USED('ctrl')
    USE IN ctrl
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT ctrl 
  SCAN 
   m.f     = LEFT(file,1)
   m.c_err = errors
   m.recid_lpu = recid
   IF m.f='R'
    m.rid = IIF(SEEK(m.recid_lpu, 'people'), people.recid, 0)
   ELSE 
    m.rid = IIF(SEEK(m.recid_lpu, 'talon'), talon.recid, 0)
   ENDIF 
   INSERT INTO error FROM MEMVAR 
  ENDSCAN 
  USE IN ctrl
  USE IN error
  USE IN talon 

  IF fso.FileExists(m.locdir+'\reestr.dbf')
   IF OpenFile(m.locdir+'\reestr', 'reestr', 'excl')>0
    IF USED('resstr')
     USE IN reestr
    ENDIF 
   ELSE 
    SELECT reestr 
    INDEX on recid TAG recid
    SET ORDER to recid 
    
    SELECT people 
    SET RELATION TO recid_lpu INTO reestr 
    SCAN 
     REPLACE ALL sv WITH reestr.sv
    ENDSCAN 
    SET RELATION OFF INTO reestr
    SELECT reestr 
    SET ORDER TO 
    DELETE TAG recid 
    
    IF fso.FileExists(m.locdir+'\erz_ans.dbf')
     IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.dbf')
      fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.dbf')
     ENDIF 
     IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.cdx')
      fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.cdx')
     ENDIF 
     IF fso.FileExists(m.locdir+'\erz_ans.dbf')
      fso.CopyFile(m.locdir+'\erz_ans.dbf',pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.dbf')
     ENDIF 
     IF fso.FileExists(m.locdir+'\erz_ans.cdx')
      fso.CopyFile(m.locdir+'\erz_ans.cdx',pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.cdx')
     ENDIF 
     IF OpenFile(m.locdir+'\erz_ans', 'erzans', 'excl')>0
      IF USED('erz_ans')
       USE IN erz_ans
      ENDIF 
     ELSE
      SELECT reestr 
      INDEX on id_erz TAG id_erz
      SET ORDER TO id_erz
      SET RELATION TO recid INTO people
      SELECT erzans
      SET RELATION TO recid INTO reestr ADDITIVE 
      IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers', 'answers', 'excl')>0
      ENDIF 
      SELECT answers
      DELETE TAG ALL 
      ZAP 
      SELECT erzans
      SCAN 
       SCATTER MEMVAR 
       RELEASE recid 
       m.recid = PADL(people.recid,6,'0')
       INSERT INTO answers FROM MEMVAR 
      ENDSCAN 
      SET RELATION OFF INTO reestr 
      USE IN erzans
      SELECT reestr 
      SET ORDER TO 
      DELETE TAG id_erz
      IF USED('answers')
       USE IN answers
      ENDIF 
     ENDIF 
    ENDIF 
    
    USE IN reestr 
   ENDIF 
  ENDIF 

  USE IN people 
  
  
  SELECT lpu 
  
 ENDSCAN
 
 WAIT CLEAR 
 
 RELEASE ospr 
 USE IN lpu 
 USE IN aisoms
 USE IN sprlpu
 USE IN profus

RETURN 
