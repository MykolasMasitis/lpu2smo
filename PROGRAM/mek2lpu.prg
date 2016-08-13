PROCEDURE Mek2Lpu

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN
 ENDIF 
  
 SELECT AisOms
 SCAN FOR !DELETED()
  IF m.gcUser!='OMS'
   SET FILTER TO Usr == PADL(ALLTRIM(gcUser),6)
  ENDIF 

  WAIT mcod WINDOW NOWAIT 

  lcPath = pBase+'\'+m.gcperiod+'\'+mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcPath+'\people.dbf') OR !fso.FileExists(lcPath+'\talon.dbf')
   LOOP 
  ENDIF 

  m.lpath = pbase+'\'+m.gcperiod+'\'+mcod
  DocName = m.lpath + '\b_mek_' + mcod

  IF fso.FileExists(DocName)
   LOOP 
  ENDIF 
  
  =SendMek(m.lpath)
  
  SELECT aisoms

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 WAIT CLEAR 

 USE
 USE IN sprlpu

 SET ESCAPE &OldEscStatus

RETURN 
 
