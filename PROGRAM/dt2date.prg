FUNCTION dt2date(lcData) && Конверитруем из [Thu,  2 Feb 2012 08:42:29 +0300] (RFC822) в datetime-формат
 IF SET("Hours")!=24
  SET HOURS TO 24
 ENDIF 

 startpos = RAT(' ',lcData,5)
 lcDay    = PADL(ALLTRIM(SUBSTR(lcData,startpos,2)),2,'0')
 lcMonthT = SUBSTR(lcData,startpos+3,3)
 DO CASE
  CASE lcMonthT = 'Jan'
   lcMonth = '01'
  CASE lcMonthT = 'Feb'
   lcMonth = '02'
  CASE lcMonthT = 'Mar'
   lcMonth = '03'
  CASE lcMonthT = 'Apr'
   lcMonth = '04'
  CASE lcMonthT = 'May'
   lcMonth = '05'
  CASE lcMonthT = 'Jun'
   lcMonth = '06'
  CASE lcMonthT = 'Jul'
   lcMonth = '07'
  CASE lcMonthT = 'Aug'
   lcMonth = '08'
  CASE lcMonthT = 'Sep'
   lcMonth = '09'
  CASE lcMonthT = 'Oct'
   lcMonth = '10'
  CASE lcMonthT = 'Nov'
   lcMonth = '11'
  CASE lcMonthT = 'Dec'
   lcMonth = '12'
  OTHERWISE 
   lcMonth = '00'
 ENDCASE 

 lcYear  =  SUBSTR(lcData, startpos+7, 4)
 lcDate  = lcDay +'.' + lcMonth + '.' + lcYear
 lcTime = SUBSTR(lcData, startpos+12, 8)
 lcRealData = CTOT(lcDate + ' ' + lcTime)

RETURN lcRealData
