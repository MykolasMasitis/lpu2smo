PROCEDURE DelZero
 GO TOP 
 MailView.Refresh
 
 SCAN 
  WAIT mcod WINDOW NOWAIT 
  MailView.Refresh
  lcDir = pBase + '\' + m.gcperiod + '\' + mcod
  IF !fso.FolderExists(lcDir)
   DELETE 
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcDir+'\People.dbf')
   DELETE 
   LOOP 
  ENDIF 
  IF OpenFile(lcDir+'\People', 'People', 'shar')>0
*   DELETE 
   LOOP 
  ENDIF 
  IF RECCOUNT('People')==0
   USE IN People 
   DELETE 
   LOOP 
  ENDIF 
  USE IN people 
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 GO TOP 
 MailView.Refresh
RETURN 