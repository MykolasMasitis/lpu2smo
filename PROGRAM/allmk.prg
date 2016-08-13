PROCEDURE AllMK
 PUBLIC oWord as Word.Application 

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

* TRY 
*  oWord=GETOBJECT(,"Word.Application")
* CATCH 
*  oWord=CREATEOBJECT("Word.Application")
* ENDTRY 

 m.mmy = PADL(m.tMonth,2,'0') + SUBSTR(STR(m.tYear,4),4,1)

 m.t_beg = SECONDS()
 SCAN FOR !DELETED()
  MailView.refresh

  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\people.dbf') OR ;
   !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\talon.dbf')
   LOOP 
  ENDIF 
 
*  DocName = pbase+'\'+m.gcperiod+'\'+mcod+'\Mk' + STR(lpuid,4) + m.qcod + m.mmy+IIF(m.gcFormat='DOC', '.doc', '.pdf')

*  IF fso.FileExists(DocName)
*   LOOP 
*  ENDIF 

  =MkPrn2(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 

 m.t_end = SECONDS()
 m.t_last = ROUND((m.t_end - m.t_beg)/60,2)

 SET ESCAPE &OldEscStatus

 oExcel.Quit

 MESSAGEBOX(TRANSFORM(m.t_last,'99999.99'),0+64,'')

RETURN 