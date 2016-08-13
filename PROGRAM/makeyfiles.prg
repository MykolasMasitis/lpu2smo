PROCEDURE MakeYFiles

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar', 'mcod')>0
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'tarif', 'shar', 'cod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', "sookod", "SHARED", "er_c")>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\lputpn', 'lputpn', 'shar', 'lpu_id')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 

 mmy      = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 
 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 SELECT AisOms

 SCAN
  IF s_pred <= 0
   LOOP 
  ENDIF 
  
  m.nIsDoubles = 0
  
  m.sum_st1 = 0  && Ñóììà ê îïëàòå, ïîëó÷åííàÿ âû÷èòàíèåì ÔËÊ èç ïðåäñòàâëåííûõ
  m.sum_st1 = s_pred - sum_flk

  m.sum_st2 = 0  && Ñóììà ê îïëàòå ïî äàííûì àãðåãàòêè
    
  m.mcod = mcod
  m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  m.IsPilot = IIF(SEEK(m.mcod, 'pilot'), .t., .f.)
  
  m.s_calc_pf = finval

  WAIT m.mcod WINDOW NOWAIT 

  m.lpu_id = lpuid
  m.IsLpuTpn = IIF(SEEK(m.lpu_id, 'lputpn'), .t., .f.)

  lcPath = pBase+'\'+m.gcperiod+'\'+m.mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 
  
  =MakeYFilesOne(lcPath)

  SET DEFAULT TO (pBin)

  SELECT AisOms

  IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÏÐÅÐÂÀÒÜ ÎÁÐÀÁÎÒÊÓ?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 
 
 ENDSCAN 
 WAIT CLEAR 
 USE 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 

 SET ESCAPE &OldEscStatus
 
 MESSAGEBOX("ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!",0+64,"")
 
RETURN 


