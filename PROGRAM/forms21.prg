PROCEDURE FormS21 && ÀÂÚ‡Î¸ÌËÍ ÔÓÎÌ˚È
 IF MESSAGEBOX('¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹'+CHR(13)+CHR(10)+'Œ“◊≈“ Œ —À”◊¿ﬂ’ À≈“¿À‹ÕŒ√Œ »—’Œƒ¿'+CHR(13)+CHR(10)+;
  '¬ –¿«–≈«≈ œ–Œ‘»À≈… (œŒÀÕ€…)?',4+32,'Sovita')=7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿'+CHR(13)+CHR(10)+UPPER(pBase+'\'+m.gcPeriod),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\Talon.dbf')
  MESSAGEBOX('—¬ŒƒÕ€… —◊≈“ «¿ œ≈–»Œƒ Õ≈ —Œ¡–¿Õ!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\Nsi\sprlpuxx.dbf')
  MESSAGEBOX('‘¿…À '+UPPER(pBase+'\'+m.gcPeriod+'\Nsi\sprlpuxx.dbf')+' Õ≈ Õ¿…ƒ≈Õ!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pCommon+'\prv002xx.dbf')
  MESSAGEBOX('‘¿…À '+UPPER(pCommon+'\prv002xx.dbf')+' Õ≈ Õ¿…ƒ≈Õ!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('Talon')
   USE IN talon
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  USE IN Talon
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pCommon+'\prv002xx', 'prv002', 'shar', 'profil')>0
  USE IN sprlpu
  USE IN Talon
  IF USED('prv002')
   USE IN prv002
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdeads (period c(6), mcod c(7), lpuid c(4), lpuname c(40), sn_pol c(17), c_i c(30), cod n(6), d_u d, ;
  prv c(3), pr_name c(100), e_period c(6), docexp c(7))
 INDEX on mcod+sn_pol TAG msn_pol
 SET ORDER TO msn_pol 
 m.period = m.gcperiod
 
 WAIT "Œ“¡Œ–.." WINDOW NOWAIT 
 SELECT aisoms
 SET RELATION TO lpuid INTO sprlpu
 SCAN 
  m.mcod    = mcod 
  m.lpuid   = STR(lpuid,4)
  m.lpuname = sprlpu.name
  IF INT(VAL(SUBSTR(m.mcod,3,2)))<41
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'recid')>0
   USE IN talon 
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  SELECT talon 
  SET RELATION TO profil INTO prv002
  SET RELATION TO recid INTO merror ADDITIVE 
*  SET SKIP TO merror
  SCAN 
   IF Tip!='5'
    LOOP 
   ENDIF 
   m.sn_pol = sn_pol
   m.c_i    = c_i
   m.cod    = cod
   m.d_u    = d_u
   m.prv    = profil
   m.e_period = IIF(!EMPTY(merror.e_period), merror.e_period, '')
   m.docexp   = IIF(!EMPTY(merror.docexp), merror.docexp, '')
   m.pr_name  = prv002.pr_name
  
   m.unik = m.mcod + m.sn_pol
   IF !SEEK(m.unik, 'curdeads')
    INSERT INTO curdeads FROM MEMVAR 
   ELSE 
    IF EMPTY(curdeads.e_period) AND !EMPTY(m.e_period)
     UPDATE curdeads SET e_period=m.e_period WHERE m.mcod + m.sn_pol=m.unik
    ENDIF 
    IF EMPTY(curdeads.docexp) AND !EMPTY(m.docexp)
     UPDATE curdeads SET docexp=m.docexp WHERE m.mcod + m.sn_pol=m.unik
    ENDIF 
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO merror
  SET RELATION OFF INTO prv002
  USE IN merror 
  USE IN talon 
  
  SELECT aisoms

 ENDSCAN 
 WAIT CLEAR 
 
 SET RELATION OFF INTO sprlpu 

 USE IN aisoms
 USE IN sprlpu
 USE IN prv002

 SELECT curdeads 
 SET ORDER TO 
 INDEX on lpuid TAG lpuid 
 SET ORDER TO lpuid
 
 m.llResult = X_Report(pTempl+'\FormS21.xls', pBase+'\'+m.gcperiod+'\FormS21.xls', .T.)
 
 SELECT curdeads 
 COPY TO &pbase\&gcperiod\curdeadsfull
 USE IN curdeads
 
RETURN 