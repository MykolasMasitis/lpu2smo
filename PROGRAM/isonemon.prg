FUNCTION IsOneMon(para1)
 PRIVATE m.flcod
 m.flcod = m.para1
 
 IF LEN(m.flcod)!=12
  RETURN .f. 
 ENDIF 
 
 m.rslt = .t.
 m.ncnt = 0 
 FOR m.nmn=1 TO 12
  m.cchar = SUBSTR(m.flcod,m.nmn,1)
  IF !INLIST(m.cchar,'0','1')
   m.rslt = .f.
   EXIT 
  ENDIF 
  IF m.cchar='1'
   m.ncnt = m.ncnt + 1
   IF m.ncnt>1
    m.rslt = .f.
    EXIT 
   ENDIF 
  ENDIF 
 ENDFOR

RETURN m.rslt