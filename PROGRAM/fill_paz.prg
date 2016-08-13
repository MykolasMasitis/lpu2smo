PROCEDURE fill_paz
SELECT AisOms 
SET ORDER TO 
GO TOP 

SCAN
 lcDir = pBase + '\' + m.gcperiod + '\' + mcod
 IF fso.FileExists(lcDir+'\people.dbf')
  tn_result = 0
  tn_result = tn_result + OpenFile("&lcDir\People", "People", "SHARE")
  IF tn_result == 0
   m.paz = RECCOUNT('people')
   USE IN People
   SELECT AisOms
   REPLACE paz WITH m.paz
  ENDIF 
 ENDIF 
 MailView.refresh
ENDSCAN 

