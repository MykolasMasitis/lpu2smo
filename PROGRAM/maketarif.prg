PROCEDURE MakeTarif
 IF MESSAGEBOX('"—Œ¡–¿“‹" ‘¿…À “¿–»‘¿?',4+16,'')==7
  RETURN 
 ENDIF 
 
 WAIT "—Œ«ƒ¿Õ»≈ ‘¿…À¿, ∆ƒ»“≈..." WINDOW NOWAIT 
 
 CREATE TABLE pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn' (cod n(6), vmp n(1), comment c(40), name c(250), uet1 n(6,2), uet2 n(6,2),;
  tarif n(9,2), tarif_v n(9,2), n_kd n(3), stkd n(9,2), stkdv n(9,2))
 INDEX ON cod TAG cod 
 USE 
 
 pNsi = 'd:\lpu2smo\nsi'

 tn_result = 0
 tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', "TarifN", "excl")
 tn_result = tn_result + OpenFile(pcommon+'\Tarif', "Tarif", "excl", "cod")
 tn_result = tn_result + OpenFile(pNSI+'\usvmp', "usvmp", "excl")
 tn_result = tn_result + OpenFile(pNSI+'\tarimu', "tarimu", "excl")
 tn_result = tn_result + OpenFile(pNSI+'\reesms', "reesms", "excl")
 
 SELECT tarimu
 INDEX ON cod TAG cod 
 SET ORDER TO cod

 SELECT reesms
 INDEX ON cod TAG cod
 SET ORDER TO cod
 
 SELECT usvmp
 SET RELATION TO cod INTO tarimu
 SET RELATION TO cod INTO reesms ADDITIVE 
 SET RELATION TO cod INTO tarif ADDITIVE 
 
 SCAN 
  SCATTER MEMVAR 
  m.comment = tarif.comment
  m.uet1    = tarif.uet1
  m.uet2    = tarif.uet2
  m.tarif   = tarimu.tarif
  m.tarif_v = tarimu.tarif_v
  m.stkd    = tarimu.stkd
  m.stkdv   = tarimu.stkdv
  m.n_kd    = reesms.n_kd
  
  INSERT INTO tarifn FROM MEMVAR 
  
 ENDSCAN 
 
 SET RELATION OFF INTO tarif 
 SET RELATION OFF INTO reesms
 SET RELATION OFF INTO tarimu 
 USE 

 SELECT tarif 
 USE

 SELECT reesms 
 SET ORDER TO 
 DELETE TAG ALL 
 USE 

 SELECT tarimu
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 
 SELECT tarifn
 use
 
 WAIT CLEAR 
 
RETURN 