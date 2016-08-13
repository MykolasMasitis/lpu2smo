PROCEDURE DelMcFiles
 IF MESSAGEBOX('ÝÒÀ ÏÐÎÖÅÄÓÐÀ ÓÄÀËßÅÒ ÂÑÅ'+CHR(13)+CHR(10)+;
  'ÑÔÎÐÌÈÐÎÂÀÍÍÛÅ ÐÀÍÅÅ Mc-ÔÀÉËÛ!'+CHR(13)+CHR(10)+'ÏÐÎÄÎËÆÈÒÜ?',4+32, '')==7
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
  mcfiledoc = 'Mc'+STR(m.lpuid,4)+UPPER(m.qcod)+PADL(tmonth,2,'0')+RIGHT(STR(tYear,4),1)+'.doc'
  mcfilexls = 'Mc'+STR(m.lpuid,4)+UPPER(m.qcod)+PADL(tmonth,2,'0')+RIGHT(STR(tYear,4),1)+'.xls'
  mcfilepdf = 'Mc'+STR(m.lpuid,4)+UPPER(m.qcod)+PADL(tmonth,2,'0')+RIGHT(STR(tYear,4),1)+'.pdf'
*  m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
*  IF m.usr != m.gcUser AND m.gcUser!='OMS'
*   LOOP 
*  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 

  lcDir = pBase + '\' + m.gcperiod + '\' + mcod
  IF !fso.FolderExists(lcDir)
   LOOP 
  ENDIF 

  IF fso.FileExists(lcDir+'\'+McFileDoc)
   fso.DeleteFile(lcDir+'\'+McFileDoc)
  ENDIF 
  IF fso.FileExists(lcDir+'\'+McFileXls)
   fso.DeleteFile(lcDir+'\'+McFileXls)
  ENDIF 
  IF fso.FileExists(lcDir+'\'+McFilePdf)
   fso.DeleteFile(lcDir+'\'+McFilePdf)
  ENDIF 
  
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE IN AisOms
 USE IN UsrLpu
 
RETURN 