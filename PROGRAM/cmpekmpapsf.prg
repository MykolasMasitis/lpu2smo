PROCEDURE CmpEkmpApsf
 IF MESSAGEBOX('ÑÂÅÐÈÒÜ ÂÂÅÄÅÍÍÛÅ ÄÀÍÍÛÅ ÏÎ ÑÍßÒÈßÌ'+CHR(13)+CHR(10)+;
  'Ñ ÑÓÌÌÀÌÈ ÑÍßÒÈÉ Â ME-ÔÀÉËÀÕ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 oal = ALIAS()
 orec = RECNO()

 CREATE CURSOR curdif (lpuid n(4), mcod c(7), s_apsf n(11,2), s_mefile n(11,2), v c(2))
 INDEX on mcod TAG mcod
 SET ORDER TO mcod 
 
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
  
  m.v = IIF(m.s_apsf=m.s_mefile, 'OK', '!!')

  IF m.s_apsf>0 OR  m.s_mefile>0
   INSERT INTO curdif FROM MEMVAR 
  ENDIF 
  
  SELECT &oal
  
 ENDSCAN 
 
 SELECT curdif
 BROWSE 
 USE IN curdif
 
 SELECT &oal
 GO (orec)
RETURN 