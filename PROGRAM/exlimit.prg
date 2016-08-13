PROCEDURE ExLimit
 PARAMETERS ppath
 
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÇÀÁÐÀÊÎÂÀÒÜ ×ÀÑÒÜ Ñ×ÅÒÀ'+CHR(13)+CHR(10)+;
  'ÏÎ ÏÐÅÂÛØÅÍÈÞ ËÈÌÈÒÀ ÔÈÍÀÍÑÈÐÎÂÀÍÈß?',4+32,'')=7
  RETURN 
 ENDIF 
 
 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+32,'')=7
  RETURN 
 ENDIF 
 
 m.exsum      = 0
 m.deltaparam = 0 
 DO FORM exlimit
* IF m.exsum = 0
*  RETURN 
* ENDIF 
 
 WAIT "ÎÁÐÀÁÎÒÊÀ..." WINDOW NOWAIT 

 m.mcod    = mcod
 
 IF !OpBase(ppath)
  RETURN .f.
 ENDIF 

 m.s_flk = 0  

 SELECT talon 
 m.oksum = 0 
 SCAN 
  m.recid = recid
  m.s_all = s_all
  m.oksum = m.oksum + m.s_all
  IF m.exsum>0
   IF m.oksum <= m.exsum
    LOOP 
   ENDIF 
  ENDIF 
  IF m.deltaparam>0 AND m.s_all>m.deltaparam
   LOOP 
  ENDIF 
  m.oksum = m.oksum - m.s_all
  rval = InsError('S', 'PPA', m.recid)
  m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
 ENDSCAN
  
 CREATE CURSOR AllBad (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol 
 SELECT talon 
 SET ORDER TO sn_pol
 GO TOP 
  
 DO WHILE !EOF()
  m.polis = sn_pol
  m.lAllBad=.t.
  DO WHILE sn_pol = m.polis
   m.recid = recid
   IF !SEEK(m.RecId, 'sError')
    m.lAllBad = .f.
   ENDIF 
   SKIP 
  ENDDO 
  IF m.lAllBad
   IF !SEEK(m.polis, 'Allbad')
    INSERT INTO Allbad (sn_pol) VALUES (m.polis)
   ENDIF 
  ENDIF 
 ENDDO 
  
 SELECT People
 SCAN 
  m.polis = sn_pol
  m.recid = recid
  IF !SEEK(m.polis, 'allbad')
   LOOP 
  ENDIF 
  IF !SEEK(RecId, 'rError')
    =InsError('R', 'PNA', m.recid)
  ENDIF 
 ENDSCAN 
 USE IN AllBad
  
 SELECT talon 
 SUM(s_all) FOR SEEK(RecId, 'sError') TO m.s_flk
 SET RELATION OFF INTO people
  
 =ClBase()

 SELECT AisOms
  
 REPLACE sum_flk WITH m.s_flk

 WAIT CLEAR 
RETURN 

FUNCTION InsError(WFile, cError, cRecId)
 IF WFile == 'R'
  IF !SEEK(cRecId, 'rError')
   INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
  ELSE 
   IF cError != rError.c_err
    INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
   ENDIF cError != rError.c_err
  ENDIF !SEEK(cRecId, 'rError')
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(cRecId, 'sError')
   INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   ENDIF cError != sError.c_err 
  ENDIF !SEEK(cRecId, 'sError')
 ENDIF 
RETURN .F.

FUNCTION InsErrorSV(mmcod, WFile, cError, cRecId)
 IF WFile == 'R'
  IF !SEEK(mmcod+STR(cRecId,9), 'resv')
   INSERT INTO resv (mcod, f, c_err, rid) VALUES (mmcod, 'R', cError, cRecId)
  ELSE 
   IF cError != resv.c_err
    INSERT INTO resv (mcod, f, c_err, rid) VALUES (mmcod, 'R', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(mcod+STR(cRecId,9), 'sesv')
   INSERT INTO sesv (mcod, f, c_err, rid) VALUES (mmcod, 'S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sesv.c_err
    INSERT INTO sesv (mcod, f, c_err, rid) VALUES (mmcod, 'S', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
RETURN .F.

FUNCTION OpBase(ppath)
  IF OpenFile(ppath+'\people', 'people', 'share', 'sn_pol')>0
   RETURN .f.
  ENDIF 
  IF OpenFile(ppath+'\talon', 'talon', 'share', 'd_u')>0
   USE IN people
   RETURN .f.
  ENDIF 
  IF OpenFile(ppath+'\e'+m.mcod, 'rerror', 'share', 'rrid')>0
   USE IN people
   USE IN talon
   RETURN .f.
  ENDIF 
  IF OpenFile(ppath+'\e'+m.mcod, 'serror', 'share', 'rid', 'again')>0
   USE IN people
   USE IN talon
   USE IN rerror
   RETURN .f.
  ENDIF 

*  DELETE FOR SUBSTR(c_err,3,1)='A' IN rerror
*  DELETE FOR SUBSTR(c_err,3,1)='A' IN serror
RETURN .t.

FUNCTION ClBase()
 USE IN talon 
 USE IN people
 USE IN rerror
 USE IN serror
RETURN  
