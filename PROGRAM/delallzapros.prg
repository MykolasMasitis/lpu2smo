PROCEDURE DelAllZapros
 IF MESSAGEBOX('ÁÓÄÓÒ ÓÄÀËÅÍÛ ÂÑÅ ÑÔÎÐÌÈÐÎÂÀÍÍÍÛÅ ÐÀÍÅÅ '+CHR(13)+CHR(10)+;
               'ÔÀÉËÛ ÇÀÏÐÎÑÎÂ Ê ÅÐÇ (Zapros.dbf)!'+CHR(13)+CHR(10)+;
               'ÝÒÎ ÒÎ, ×ÒÎ ÂÛ ÄÅÉÑÒÂÈÒÅËÜÍÎ ÕÎÒÈÒÅ ÑÄÅËÀÒÜ?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+48,'') != 6
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN
  m.mcod = mcod
*  m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
*  IF m.usr != m.gcUser AND m.gcUser!='OMS'
*   LOOP 
*  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 

  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+mcod+'\Zapros.dbf')
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+mcod+'\Zapros.dbf')
  ENDIF 
  
  IF fso.FileExists(m.pAisOms+'\'+m.gcUser+'\OUTPUT\berz_'+mcod)
   fso.DeleteFile(m.pAisOms+'\'+m.gcUser+'\OUTPUT\berz_'+mcod)
  ENDIF 
  
  IF fso.FileExists(m.pAisOms+'\'+m.gcUser+'\OUTPUT\derz_'+mcod)
   fso.DeleteFile(m.pAisOms+'\'+m.gcUser+'\OUTPUT\derz_'+mcod)
  ENDIF 
  
  REPLACE erz_id WITH '', erz_status WITH 0

 ENDSCAN 
 WAIT CLEAR 
 USE 
 USE IN UsrLpu

RETURN 