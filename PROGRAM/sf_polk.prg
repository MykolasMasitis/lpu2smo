PROCEDURE sf_polk
 oal = ALIAS()
 orec = RECNO()
 
 m.mcod = mcod
 m.k_u_sum = 0
 m.s_all_sum = 0
 REPORT FORM sf_polk PREVIEW  
 
 SELECT (oal)
 GO RECNO()
RETURN 