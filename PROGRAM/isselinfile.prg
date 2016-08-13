FUNCTION IsSelInFile(par1, par2, par3)
 PRIVATE m.cselect, m.cflname, m.nlenght
 IF PARAMETERS() != 3
  RETURN .f.
 ENDIF 

 m.cselect = par1
 m.cflname = par2
 m.nlenght = par3
 
 IF m.nlenght<=0
  RETURN .f.
 ENDIF 
 
 IF LEN(m.cselect) != m.nlenght OR LEN(m.cselect) != m.nlenght
  RETURN .f.
 ENDIF 

 m.rslt = .t.
 FOR m.nchar=1 TO m.nlenght
  m.selchar = SUBSTR(m.cselect, m.nchar, 1)
  m.fllchar = SUBSTR(m.cflname, m.nchar, 1)
  
  IF !INLIST(m.selchar,'0','1') OR !INLIST(m.fllchar,'0','1')
   m.rslt = .f.
   EXIT 
  ENDIF 
  
  IF m.selchar='1' AND m.fllchar='0'
   m.rslt = .f.
   EXIT 
  ENDIF 
  
 ENDFOR 
 
RETURN m.rslt