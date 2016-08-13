PROCEDURE ResetErrs2(calias)

 m.et = '2'
 DO FORM SelTipOfExp TO m.lResp 
 
 IF !m.lResp 
  RETURN 
 ENDIF 

 oal   = ALIAS()
 orecp = RECNO()
 m.ppolis = sn_pol
 SELECT (calias)
 SCAN FOR V
  m.recid = recid 
  m.vvir = PADL(m.recid,6,'0')+m.et
  IF SEEK(m.vvir, 'merror', 'id_et')
   DELETE FROM merror WHERE recid=m.recid AND et=m.et
  ENDIF 
 ENDSCAN 
 SELECT &oal
 GO (orecp)

RETURN 