FUNCTION MakeEkmpAll(mcod)
 pnResult = 0
 m.mcod = mcod
 m.ppath = m.pbase+'\'+m.gcperiod+'\'+m.mcod
 pnResult = pnResult + OpenFile(pPath+'\people', "people", "SHARED", "sn_pol")
 pnResult = pnResult + OpenFile(pPath+'\talon', "talon", "SHARED", "recid")
 pnResult = pnResult + OpenFile(pPath+'\e'+m.mcod, "rerror", "SHARED", "rrid")
 pnResult = pnResult + OpenFile(pPath+'\e'+m.mcod, "serror", "SHARED", "rid", "again")
 pnResult = pnResult + OpenFile(pPath+'\m'+m.mcod, "merror", "SHARED", "recid")
 pnResult = pnResult + OpenFile(pPath+'\sprsels', "sprsel", "SHARED", "recid")
 pnResult = pnResult + OpenFile(pPath+'\doctor', "doctor", "SHARED", "pcod")
 pnResult = pnResult + OpenFile(pPath+'\otdel', "otdel", "SHARED", "iotd")

 IF pnResult > 0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
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
  IF USED('doctor')
   USE IN doctor
  ENDIF 
  IF USED('otdel')
   USE IN otdel
  ENDIF 
  IF USED('sprsel')
   USE IN sprsel
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 CREATE CURSOR curpols (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT merror
 SCAN 
  IF SEEK(recid, 'talon', 'recid') AND INLIST(et,'4','5','6','9')
   m.sn_pol = talon.sn_pol
   IF !SEEK(m.sn_pol, 'curpols')
    INSERT INTO curpols FROM MEMVAR 
   ENDIF 
  ENDIF 
 ENDSCAN 
 
 SELECT curpols
* BROWSE 
 IF RECCOUNT('curpols')<=0
  MESSAGEBOX('Â ÂÛÁÐÀÍÍÎÌ ËÏÓ ÝÊÌÏ ÍÅ ÏÐÎÂÎÄÈËÀÑÜ!',0+64,'')
 ELSE 
  SCAN 
   m.sn_pol = sn_pol
   WAIT m.sn_pol WINDOW NOWAIT 
   =MakeEkmp(m.sn_pol, pbase+'\'+gcperiod+'\'+mcod, .F., .T., goApp.TipAcc, '0')
   WAIT CLEAR 
   SELECT curpols
  ENDSCAN 
 ENDIF 
 USE IN curpols
 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('talon')
  USE IN talon
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
 IF USED('doctor')
  USE IN doctor
 ENDIF 
 IF USED('otdel')
  USE IN otdel
 ENDIF 
 IF USED('sprsel')
  USE IN sprsel
 ENDIF 
 SELECT aisoms

RETURN 

