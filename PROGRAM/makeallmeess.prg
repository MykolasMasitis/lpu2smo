FUNCTION MakeAllMEEss(m.mcod)

 IF OpBase(m.mcod)>0
  =ClBase()
  SELECT aisoms
  RETURN
 ENDIF 
 
 SET ORDER TO recid IN talon 
 SELECT merror
 SET RELATION TO recid INTO talon 
 
 SCAN 
  m.et = et
  IF !INLIST(m.et,'2','3','7','8')
   LOOP  
  ENDIF 
  m.sn_pol = talon.sn_pol
  m.vir    = m.sn_pol+' '+m.et
  IF !SEEK(m.vir, 'curtalon')
   INSERT INTO curtalon FROM MEMVAR 
  ENDIF 
 ENDSCAN 
 
 IF RECCOUNT('curtalon')<=0
  =ClBase()
  SELECT aisoms
  MESSAGEBOX('ÏÎ ÂÛÁÐÀÍÍÎÌÓ ËÏÓ ÌÝÝ ÍÅ ÏÐÎÂÎÄÈËÀÑÜ!',0+64,'')
  RETURN 
 ENDIF 
 
 WAIT "ÇÀÏÓÑÊ WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 

 SELECT curtalon 
 SET ORDER TO 
 DELETE TAG vir
 GO TOP 

 SELECT merror 
 SET RELATION OFF INTO talon 
 
 SELECT Talon
 SET ORDER TO sn_pol
 SET RELATION TO recid INTO serror 
 SELECT people
 SET RELATION TO RecId INTO rError
 
 SELECT curtalon
 SCAN
  m.sn_pol = sn_pol
  m.et     = et
  WAIT m.sn_pol WINDOW NOWAIT 
  =MakeMEESS(m.sn_pol, .f., .f., goApp.TipAcc, m.et)
  WAIT CLEAR 
 ENDSCAN 
 
 =ClBase()
 SELECT aisoms
 
 oWord.Quit
 
 MESSAGEBOX('ÔÎÐÌÈÐÎÂÀÍÈÅ ÀÊÒÎÂ ÇÀÊÎÍ×ÅÍÎ!',0+64,'')

RETURN 

FUNCTION OpBase(m.mcod)
 PRIVATE ppath 

 m.ppath = m.pbase+'\'+m.gcperiod+'\'+m.mcod
 
 pnResult = 0
 pnResult = pnResult + OpenFile(pPath+'\people', "people", "SHARED", "sn_pol")
 pnResult = pnResult + OpenFile(pPath+'\talon', "talon", "SHARED")
 pnResult = pnResult + OpenFile(pPath+'\e'+m.mcod, "rerror", "SHARED", "rrid")
 pnResult = pnResult + OpenFile(pPath+'\e'+m.mcod, "serror", "SHARED", "rid", "again")
 pnResult = pnResult + OpenFile(pPath+'\m'+m.mcod, "merror", "SHARED", "recid")
 pnResult = pnResult + OpenFile(pPath+'\doctor', "doctor", "SHARED", "pcod")
 pnResult = pnResult + OpenFile(pPath+'\otdel', "otdel", "SHARED", "iotd")
 
 CREATE CURSOR curtalon (sn_pol c(25), et c(1))
 INDEX on sn_pol+' '+et TAG vir
 SET ORDER TO vir
 
RETURN pnResult

FUNCTION ClBase
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('doctor')
  USE IN doctor
 ENDIF 
 IF USED('otdel')
  USE IN otdel
 ENDIF 
 IF USED('rerror')
  USE IN rerror
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
 IF USED('merror')
  USE IN merror
 ENDIF 
 IF USED('curtalon')
  USE IN curtalon
 ENDIF 
RETURN

