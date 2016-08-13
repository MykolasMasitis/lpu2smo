PROCEDURE erz_reset
 GO TOP 
 SCAN 
  REPLACE erz_status WITH 0
  _vfp.ActiveForm.refresh
 ENDSCAN 
 GO TOP 
RETURN 