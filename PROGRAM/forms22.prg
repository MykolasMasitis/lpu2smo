PROCEDURE FormS22 && ÀÂÚ‡Î¸ÌËÍ Í‡ÚÍËÈ
 IF MESSAGEBOX('¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹'+CHR(13)+CHR(10)+'Œ“◊≈“ Œ —À”◊¿ﬂ’ À≈“¿À‹ÕŒ√Œ »—’Œƒ¿'+CHR(13)+CHR(10)+;
  '¬ –¿«–≈«≈ œ–Œ‘»À≈… ( –¿“ »…)?',4+32,'Sovita')=7
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
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\Talon', 'talon', 'shar')>0
  IF USED('Talon')
   USE IN talon
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
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
 
 CREATE CURSOR curdeads (mcod c(7), lpuid n(4), sn_pol c(17), prv c(3))
 INDEX on mcod+sn_pol TAG msn_pol
 SET ORDER TO msn_pol 
 
 WAIT "Œ“¡Œ–.." WINDOW NOWAIT 
 SELECT talon 
 SET RELATION TO mcod INTO sprlpu 
 SCAN 
  IF Tip!='5'
   LOOP 
  ENDIF 
  m.mcod   = mcod 
  m.lpuid  = sprlpu.lpu_id
  m.sn_pol = sn_pol
  m.prv    = profil
  
  m.unik = m.mcod + m.sn_pol
  IF !SEEK(m.unik, 'curdeads')
   INSERT INTO curdeads FROM MEMVAR 
  ENDIF 
  
 ENDSCAN 
 SET RELATION OFF INTO sprlpu 
 USE IN talon 
 WAIT CLEAR 
 
 SELECT lpuid,prv,COUNT(*) as cnt FROM curdeads GROUP BY lpuid,prv ORDER BY lpuid,prv INTO CURSOR curdata READWRITE 
 ALTER TABLE curdata ADD COLUMN period c(6)
 ALTER TABLE curdata ADD COLUMN lpuname c(40)
 ALTER TABLE curdata ADD COLUMN prvname c(100)
 SELECT curdata 
 SET ORDER TO lpu_id IN sprlpu 
 SET RELATION TO lpuid INTO sprlpu
 SET RELATION TO prv INTO prv002 ADDITIVE 
 REPLACE ALL period WITH m.gcperiod, lpuname WITH sprlpu.name, prvname WITH prv002.pr_name
 SET RELATION OFF INTO prv002
 SET RELATION OFF INTO sprlpu 
 
 SELECT curdeads 
 COPY TO &pbase\&gcperiod\curdeads
 USE IN curdeads
 USE IN sprlpu
 USE IN prv002

 m.llResult = X_Report(pTempl+'\FormS22.xls', pBase+'\'+m.gcperiod+'\FormS22.xls', .T.)
 
 SELECT curdata 
 COPY TO &pbase\&gcperiod\curdata
 USE IN curdata 
 
RETURN 