PROCEDURE AllOms6u

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 
 
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
  
  DocName = pbase+'\'+m.gcperiod+'\'+mcod+'\s_'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)+'.pdf'

  IF fso.FileExists(DocName)
   fso.DeleteFile(DocName)
  ENDIF 

   =Oms6un(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('бш унрхре опепбюрэ напюанрйс?',4+32,'') == 6
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
 
 MESSAGEBOX(CHR(13)+CHR(10)+"напюанрйю гюйнмвемю!"+CHR(13)+CHR(10)+;
  "бяецн напюанрюмн кос   : "+TRANSFORM(nLpu, '9999999')+CHR(13)+CHR(10)+;
  "наыее бпелъ напюанрйх  : "+TRANSFORM(LastOfProc,'999.999')+" ЯЕЙ."+CHR(13)+CHR(10)+;
  "япедмее бпелъ напюанрйх: "+TRANSFORM(MeanTime,'999.999')+" ЯЕЙ."+CHR(13)+CHR(10),0+64,"")

 oExcel.Quit

RETURN 