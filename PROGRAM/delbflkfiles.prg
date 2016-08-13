PROCEDURE DelBFlkFiles

 IF MESSAGEBOX('ÁÓÄÓÒ ÓÄÀËÅÍÛ ÂÑÅ ÑÔÎÐÌÈÐÎÂÀÍÍÍÛÅ ÐÀÍÅÅ '+CHR(13)+CHR(10)+;
               'ÔÀÉËÛ B_FLK!'+CHR(13)+CHR(10)+;
               'ÝÒÎ ÒÎ, ×ÒÎ ÂÛ ÄÅÉÑÒÂÈÒÅËÜÍÎ ÕÎÒÈÒÅ ÑÄÅËÀÒÜ?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+48,'') != 6
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN FOR !DELETED()
  WAIT mcod WINDOW NOWAIT 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+mcod+'\b_flk_'+mcod)
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+mcod+'\b_flk_'+mcod)
  ENDIF 
 ENDSCAN 
 WAIT CLEAR 
 USE 

RETURN 
