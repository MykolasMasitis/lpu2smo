PROCEDURE RplEkmpApsf
 IF MESSAGEBOX('ÑÂÅÐÈÒÜ ÂÂÅÄÅÍÍÛÅ ÄÀÍÍÛÅ ÏÎ ÑÍßÒÈßÌ'+CHR(13)+CHR(10)+;
  'Ñ ÑÓÌÌÀÌÈ ÑÍßÒÈÉ Â ME-ÔÀÉËÀÕ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 oal = ALIAS()
 orec = RECNO()

 SELECT &oal
 SCAN 
  m.lpuid  = lpuid
  m.mcod   = mcod 
  m.s_apsf = e_ekmp
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
   SUM FOR INLIST(et,'4','5','6') s_opl_e TO m.s_mefile
  ENDIF 
  IF USED('mefl')
   USE IN mefl
  ENDIF 
  
  SELECT &oal
  
  IF (m.s_apsf>0 OR  m.s_mefile>0) AND m.s_apsf!=m.s_mefile
   REPLACE e_ekmp WITH m.s_mefile
  ENDIF 

 ENDSCAN 
 
 SELECT &oal
 GO (orec)
RETURN 