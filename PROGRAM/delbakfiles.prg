PROCEDURE DelBakFiles
 IF MESSAGEBOX('ÁÓÄÓÒ ÓÄÀËÅÍÛ ÂÑÅ BAK-ÔÀÉËÛ?'+CHR(13)+CHR(10)+;
               'ÝÒÎ ÒÎ, ×ÒÎ ÂÛ ÄÅÉÑÒÂÈÒÅËÜÍÎ ÕÎÒÈÒÅ ÑÄÅËÀÒÜ?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+48,'') != 6
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN
  m.mcod = mcod

  WAIT m.mcod WINDOW NOWAIT 

 
*  DELETE FILE pBase+'\'+m.gcperiod+'\'+mcod+'\*.pdf'
*  DELETE FILE pBase+'\'+m.gcperiod+'\'+mcod+'\*.doc'

  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\people.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\people.bak')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\talon.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\talon.bak')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\otdel.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\otdel.bak')
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcperiod+'\'+mcod+'\doctor.bak')
   fso.DeleteFile(pBase+'\'+m.gcperiod+'\'+mcod+'\doctor.bak')
  ENDIF 


 ENDSCAN 
 WAIT CLEAR 
 USE 

RETURN 