PROCEDURE ResetErrs(calias)

 m.et = '2'
 DO FORM SelTipOfExp TO m.lResp 
 
 IF !m.lResp 
  RETURN 
 ENDIF 

 oal   = ALIAS()
 orecp = RECNO()
 SELECT people

 m.nlocked=0
 COUNT FOR ISRLOCKED() TO m.nlocked

 SELECT &oal
 GO (orecp)

 IF m.nlocked <= 0
  oal   = ALIAS()
  orecp = RECNO()
  m.ppolis = sn_pol
  SELECT (calias)
  SCAN FOR sn_pol = ppolis
   m.recid = recid 
   m.vvir = PADL(m.recid,6,'0')+m.et
   IF SEEK(m.vvir, 'merror', 'id_et')
    DELETE FROM merror WHERE recid=m.recid AND et=m.et
   ENDIF 
  ENDSCAN 
  SELECT &oal
  GO (orecp)
 ELSE 
  IF MESSAGEBOX('ÏÎÌÅÒÈÒÜ ÎØÈÁÎ×ÍÎÉ ÂÑÅ ÎÒÎÁÐÀÍÍÛÅ (ÄÀ)?'+CHR(13)+CHR(10)+;
   'ÈËÈ ÒÎËÜÊÎ ÒÅÊÓÙÓÞ ÇÀÏÈÑÜ? (ÍÅÒ)'+CHR(13)+CHR(10),4+32,'')=6

   SELECT people
   SCAN FOR ISRLOCKED()
    m.ppolis = sn_pol

    SELECT (calias)
    SCAN FOR sn_pol = ppolis
     m.recid = recid 
     m.vvir = PADL(m.recid,6,'0')+m.et
     IF SEEK(m.vvir, 'merror', 'id_et')
      DELETE FROM merror WHERE recid=m.recid AND et=m.et
     ENDIF 
    ENDSCAN 
    SELECT people 
   
   ENDSCAN 
   SELECT &oal
   GO (orecp)

  ELSE 
   oal   = ALIAS()
   orecp = RECNO()
   m.ppolis = sn_pol
   SELECT (calias)
   SCAN FOR sn_pol = ppolis
    m.recid = recid 
    m.vvir = PADL(m.recid,6,'0')+m.et
    IF SEEK(m.vvir, 'merror', 'id_et')
     DELETE FROM merror WHERE recid=m.recid AND et=m.et
    ENDIF 
   ENDSCAN 
   SELECT &oal
   GO (orecp)
  ENDIF 
 ENDIF 


RETURN 