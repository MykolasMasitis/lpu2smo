PROCEDURE stat_erz

SELECT AisOms 
SET ORDER TO 
GO TOP 

SCAN
 lcDir = pBase + '\' + m.gcperiod + '\' + mcod
 DO CASE 
  CASE !fso.FileExists(lcDir+'\Zapros.dbf') AND !fso.FileExists(lcDir+'\Answer.dbf') 
   REPLACE erz_status WITH 0

  CASE  fso.FileExists(lcDir+'\Zapros.dbf') AND !fso.FileExists(lcDir+'\Answer.dbf') 
   REPLACE erz_status WITH 1

  CASE  fso.FileExists(lcDir+'\Zapros.dbf') AND  fso.FileExists(lcDir+'\Answer.dbf') 
   REPLACE erz_status WITH 2

  OTHERWISE 
   REPLACE erz_status WITH 0
 ENDCASE 

 MailView.refresh
ENDSCAN 

GO TOP 
MailView.refresh


