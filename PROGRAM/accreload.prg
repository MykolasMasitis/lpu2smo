FUNCTION AccReload(para1, para2, para3)
m.ppara1 = para1
m.ppara2 = para2
IF PARAMETERS()>2
 m.ppara3 = para3
ELSE 
 m.ppara3 = .t.
ENDIF 

IF tdat1<{01.05.2014}
 =AccReload1(ppara1, ppara2, ppara3)
ELSE 
 =AccReload2(ppara1, ppara2, ppara3)
ENDIF 