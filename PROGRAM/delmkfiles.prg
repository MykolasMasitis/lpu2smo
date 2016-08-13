PROCEDURE DelMkFiles
 IF MESSAGEBOX('ÝÒÀ ÏÐÎÖÅÄÓÐÀ ÓÄÀËßÅÒ ÂÑÅ'+CHR(13)+CHR(10)+;
  'ÑÔÎÐÌÈÐÎÂÀÍÍÛÅ ÐÀÍÅÅ Mk-ÔÀÉËÛ!'+CHR(13)+CHR(10)+'ÏÐÎÄÎËÆÈÒÜ?',4+32, '')==7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(''+CHR(13)+CHR(10)+;
  'ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?'+CHR(13)+CHR(10)+;
  ''+CHR(13)+CHR(10),4+32, '')==7
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar')>0
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 
 SCAN 
  m.mcod  = mcod
  m.lpuid = lpuid
  mkfiledoc = 'Mk'+STR(m.lpuid,4)+UPPER(m.qcod)+PADL(tmonth,2,'0')+RIGHT(STR(tYear,4),1)+'.doc'
  mkfilepdf = 'Mk'+STR(m.lpuid,4)+UPPER(m.qcod)+PADL(tmonth,2,'0')+RIGHT(STR(tYear,4),1)+'.pdf'
*  m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
*  IF m.usr != m.gcUser AND m.gcUser!='OMS'
*   LOOP 
*  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 

  lcDir = pBase + '\' + m.gcperiod + '\' + mcod
  IF !fso.FolderExists(lcDir)
   LOOP 
  ENDIF 

  IF fso.FileExists(lcDir+'\'+MkFileDoc)
   fso.DeleteFile(lcDir+'\'+MkFileDoc)
  ENDIF 

  IF fso.FileExists(lcDir+'\'+MkFilePdf)
   fso.DeleteFile(lcDir+'\'+MkFilePdf)
  ENDIF 
  
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE IN AisOms
 USE IN UsrLpu
 
RETURN 