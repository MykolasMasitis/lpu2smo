FUNCTION ReadFromIni(lcNameOfIniFile, lcParameter, lcDefRetValue)
 lcValueOfParameter = lcDefRetValue
 
 CFG = FOPEN (lcNameOfIniFile)
 IF CFG<0
  RETURN lcValueOfParameter
 ENDIF 
  
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  IF UPPER(READCFG) = lcParameter
   lcValueOfParameter = ALLTRIM(SUBSTR(READCFG, ATC(':',READCFG)+1))
   DO CASE 
    CASE lcValueOfParameter = '.T.'
     lcValueOfParameter = .T.
    CASE lcValueOfParameter = '.F.'
     lcValueOfParameter = .F.
    OTHERWISE 
     lcValueOfParameter = .T.
   ENDCASE 
  ENDIF 
 ENDDO
 = FCLOSE (CFG)

RETURN lcValueOfParameter
