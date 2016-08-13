FUNCTION  ShowBFile(bbb)
 tbname = ALLTRIM(bbb)
 IF !fso.FileExists(tbname)
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик'+CHR(13)+CHR(10)+;
   tbname,0+16,'')
  RETURN .f. 
 ENDIF 
 wshell.ShellExecute('notepad', tbname)
RETURN 