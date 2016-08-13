PROCEDURE RplMeApsf
 IF MESSAGEBOX('ÇÀÏÎËÍÈÒÜ ÏÎ ME-ÔÀÉËÓ?',4+32,'')=7
  RETURN 
 ENDIF 
 
 oal = ALIAS()
 orec = RECNO()

 SELECT &oal
 SCAN 
  m.lpuid  = lpuid
  m.mcod   = mcod 
  m.s_apsf = e_mee
  m.s_mefile = 0
  
  m.IsMeFile = .f.
  m.IsMeeSum = .f.
  
  m.mefile = 'ME'+UPPER(ALLTRIM(m.qcod))+STR(m.lpuid,4)
  
  IF fso.FileExists(pout+'\'+m.gcperiod+'\'+m.mefile+'.dbf')
   m.IsMeFile = .t.
  ENDIF 
  IF m.s_apsf>0
   m.IsMeeSum = .t.
  ENDIF 
  
  IF m.IsMeFile = .f. AND m.IsMeeSum = .f.
   LOOP 
  ENDIF 
  
  IF OpenFile(pout+'\'+m.gcperiod+'\'+m.mefile, 'mefl', 'shar')>0
  ELSE 
   SELECT mefl
   SUM FOR INLIST(et,'2','3','7') s_opl_e TO m.s_mefile
  ENDIF 
  IF USED('mefl')
   USE IN mefl
  ENDIF 
  
  SELECT &oal
  
  IF (m.s_apsf>0 OR  m.s_mefile>0) AND m.s_apsf!=m.s_mefile
   REPLACE e_mee WITH m.s_mefile
  ENDIF 

 ENDSCAN 
 
 SELECT &oal
 GO (orec)
RETURN 