FUNCTION  MakeCtrl(lcPath)
 fso.CopyFile(pTempl+'\Ctrl.dbf', lcPath+'\Ctrl'+m.qcod+'.dbf', .t.)

 m.mcod   = SUBSTR(lcPath, RAT('\', lcPath)+1)
 m.lpu_id = 0

 IF !USED('aisoms')
  IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shared')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   RETURN 
  ENDIF 
  m.lpu_id = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.lpuid, 0)
  USE IN aisoms
 ELSE
  m.lpu_id = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.lpuid, 0)
 ENDIF  

 EFile    = lcPath + '\e' + m.mcod
 nEFile   = lcPath + '\e'
 People   = lcPath + '\People'
 Talon    = lcPath + '\Talon'

 Ctrl = lcPath + '\Ctrl' + m.qcod
 
 pnResult = 0
 pnResult = OpenFile("&People", "People", "SHARED", "recid")
 pnResult = OpenFile("&Talon", "Talon", "SHARED", "recid")
 pnResult = OpenFile("&EFile", "Err", "SHARED")
 pnResult = OpenFile("&Ctrl", "Ctrl", "SHARED")
 pnResult = OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', "sookod", "SHARED", "er_c")
 
 SELECT Err
 SET RELATION TO RId INTO People
 SET RELATION TO RId INTO Talon ADDITIVE 
 SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 

 SCAN FOR !DELETED()
  m.mmy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
  m.file = IIF(f == 'R', 'R' + m.qcod + '.' + m.mmy, 'S' + m.qcod  + '.' + m.mmy)
  IF f == 'R'
   m.recid = People.recid_lpu
   m.fil_id = 0
  ELSE 
   m.recid = Talon.recid_lpu
   m.fil_id    = talon.fil_id
  ENDIF 
  m.errors    = c_err
  m.e_cod     = 0
  m.e_ku      = 0
  m.e_tip     = ''
  m.RefReason = sookod.refreason
  m.et230     = 1 && Если МЭК!
  m.osn230    = sookod.osn230
  
  INSERT INTO Ctrl FROM MEMVAR 
  
 ENDSCAN 
 
 SET RELATION OFF INTO Talon
 SET RELATION OFF INTO People 
 SET RELATION OFF INTO sookod
 
 USE 
 USE IN People
 USE IN Talon 
 USE IN Ctrl
 USE IN sookod

RETURN 