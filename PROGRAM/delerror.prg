PROCEDURE DelError
  IF !EMPTY(error.c_err)
   DELETE IN error
   _vfp.ActiveForm.LockScreen = .t.
   SET RELATION OFF INTO error
   SET RELATION TO RecID INTO error
   _vfp.ActiveForm.LockScreen = .f.
   _vfp.ActiveForm.refresh
  ENDIF 
RETURN 