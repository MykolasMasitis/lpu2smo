PROCEDURE AllOms6c

* PUBLIC oWord as Word.Application

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
* TRY 
*  oWord=GETOBJECT(,"Word.Application")
* CATCH 
*  oWord=CREATEOBJECT("Word.Application")
* ENDTRY 
 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 GO TOP 

 nLpu = 0
 StartOfProc = SECONDS()
 SCAN 
  m.mcod = mcod
  WAIT m.mcod WINDOW NOWAIT 

  MailView.refresh
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\people.dbf') OR ;
   !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  
  DocName = pbase+'\'+m.gcperiod+'\'+mcod+'\Pr'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)+'.pdf'

  IF fso.FileExists(DocName)
   fso.DeleteFile(DocName)
*   LOOP 
  ENDIF 

*  =Oms6cpdf(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
*  IF !tpn
*   =Oms6cword(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
*  ELSE 
*   =Oms6cwordtpn(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
*  ENDIF 

   =Oms6cn(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

  WAIT CLEAR 
  nLpu = nLpu + 1
 ENDSCAN 
 GO TOP 
 MailView.refresh
 
 EndOfProc = SECONDS()
 LastOfProc = EndOfProc - StartOfProc
 MeanTime = LastOfProc/nLpu

 WAIT CLEAR 

 SET ESCAPE &OldEscStatus
 
 MESSAGEBOX(CHR(13)+CHR(10)+"Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!"+CHR(13)+CHR(10)+;
  "¬—≈√Œ Œ¡–¿¡Œ“¿ÕŒ Àœ”   : "+TRANSFORM(nLpu, '9999999')+CHR(13)+CHR(10)+;
  "Œ¡Ÿ≈≈ ¬–≈Ãﬂ Œ¡–¿¡Œ“ »  : "+TRANSFORM(LastOfProc,'999.999')+" ÒÂÍ."+CHR(13)+CHR(10)+;
  "—–≈ƒÕ≈≈ ¬–≈Ãﬂ Œ¡–¿¡Œ“ »: "+TRANSFORM(MeanTime,'999.999')+" ÒÂÍ."+CHR(13)+CHR(10),0+64,"")

* oWord.Quit
 oExcel.Quit

RETURN 