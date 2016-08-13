PROCEDURE CrLicsDir

LicDir    = fso.GetParentFolderName(pcommon) + '\LICENCES'
IF !fso.FolderExists(LicDir)
 fso.CreateFolder(LicDir)
ELSE 
 RETURN 
ENDIF   

IF !fso.FileExists(pcommon+'\lpudogs.dbf')
 MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик-яопюбнвмхй LPUDOGS.DBF!'+CHR(13)+CHR(10),0+64,'')
 RETURN 
ENDIF 

IF !fso.FileExists(pcommon+'\prv002xx.dbf')
 MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик-яопюбнвмхй PRV002xx.DBF!'+CHR(13)+CHR(10),0+64,'')
 RETURN 
ENDIF 

IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar')>0
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 

IF OpenFile(pcommon+'\prv002xx', 'profs', 'shar')>0
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('profs')
  USE IN profs
 ENDIF 
 RETURN 

ENDIF 

SELECT lpudogs
SCAN
 m.mcod = mcod
 m.lpu_id = STR(lpu_id,4)
 
 WAIT 'янгдюеряъ яопюбнвмхй кхжемгхпнбюммшу опнтхкеи дкъ '+m.mcod+'...' WINDOW NOWAIT 
 
 licdbf = LicDir+'\l'+m.lpu_id
 
 CREATE TABLE &licdbf (v l, cod n(3)) 
 licalias = ALIAS()
 
 SELECT profs
 m.v = .T.
* SCAN FOR IsOms
 SCAN
  m.cod = INT(VAL(profil))
  INSERT INTO &licalias FROM MEMVAR 
 ENDSCAN 
 
 USE IN &licalias
 
 SELECT lpudogs 
 
ENDSCAN 
WAIT CLEAR 

IF USED('lpudogs')
 USE IN lpudogs
ENDIF 
IF USED('profs')
 USE IN profs
ENDIF 

RETURN 
 

*IF !fso.FolderExists(LicDir)
* fso.CreateFolder(LicDir)
* IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\prv002xx.dbf')
*  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\prv002xx', 'profs', 'shar')=0
*   IF 
   
*  ENDIF 
* ENDIF 
*ENDIF 
*RETURN 