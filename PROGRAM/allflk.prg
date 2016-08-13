PROCEDURE AllFlk

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
 m.totflk = MailView.get_sum_flk.Value
 SCAN FOR !IsPr
  m.locflk = sum_flk
  m.totflk = m.totflk - m.locflk
  REPLACE sum_flk WITH 0
  MailView.get_sum_flk.Value = m.totflk
  MailView.get_sum_flk.Refresh
 ENDSCAN 

 m.totflk = MailView.get_sum_flk.Value
 SCAN FOR !IsPr
  WAIT mcod WINDOW NOWAIT 
  MailView.refresh
  m.mcod  = mcod
  ppath   = m.pbase+'\'+m.gcperiod+'\'+m.mcod

  IF !fso.FolderExists(ppath)
   LOOP 
  ENDIF 
  
  IF !fso.FileExists(ppath+'\people.dbf') OR ;
     !fso.FileExists(ppath+'\talon.dbf')
   LOOP 
  ENDIF 
  
  =OneFlk(ppath)
  
  SELECT AisOms

  m.locflk = sum_flk
  m.totflk = m.totflk + m.locflk
  MailView.get_sum_flk.Value = m.totflk
  MailView.get_sum_flk.Refresh

  IF CHRSAW(0) == .T.
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus

 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!', 0+64, '')

RETURN 

*FUNCTION InsError(WFile, cError, cRecId)
* IF WFile == 'R'
*  IF !SEEK(cRecId, 'rError')
*   INSERT INTO eError (f, c_err, rid) VALUES ('R', cError, cRecId)
*  ELSE 
*   IF cError != rError.c_err
*    INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
*   ENDIF cError != rError.c_err
*  ENDIF !SEEK(cRecId, 'rError')
* ENDIF 
* IF WFile == 'S'
*  IF !SEEK(cRecId, 'sError')
*   INSERT INTO eError (f, c_err, rid) VALUES ('S', cError, cRecId)
*  ELSE 
*   IF cError != sError.c_err
*    INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
*   ENDIF cError != sError.c_err 
*  ENDIF !SEEK(cRecId, 'sError')
* ENDIF 
*RETURN 