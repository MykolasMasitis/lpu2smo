FUNCTION ResetW0ekmp(ppolis)
 orecp = RECNO()
 _vfp.ActiveForm.LockScreen = .t.
 SELECT talon 
 REPLACE FOR sn_pol = ppolis AND INLIST(et,'4','5','6') err_mee WITH '', et WITH '', e_period WITH ''
* SELECT people
 GO (orecp)
 _vfp.ActiveForm.LockScreen = .f.
RETURN 