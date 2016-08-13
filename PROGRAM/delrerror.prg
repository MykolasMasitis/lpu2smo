PROCEDURE DelrError
  IF !EMPTY(rerror.c_err)
   DELETE IN rerror
   _vfp.ActiveForm.LockScreen = .t.
   SET RELATION OFF INTO rerror
   SET RELATION TO RecID INTO rerror
   _vfp.ActiveForm.LockScreen = .f.
   _vfp.ActiveForm.refresh
  ENDIF 
RETURN 