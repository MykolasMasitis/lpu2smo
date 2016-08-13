PROCEDURE v_medexp(TipOfView, vfilter)

DO CASE 
 CASE goapp.tiplpu<40 AND !INLIST(goapp.tiplpu,1,7)&& Пол-ка с дневным стационаром
  ACTIVATE POPUP mtip2
 
 CASE goapp.tiplpu>40 && Круглосуточные стационары
  ACTIVATE POPUP mtip3

 OTHERWISE && То есть, 1 и 7 - поликлиника
  DO CASE 
   CASE m.TipOfView == 1
    DO ViewExp
   CASE m.TipOfView == 2
    DO ViewTalon
   OTHERWISE
  ENDCASE 
  
ENDCASE 

RELEASE POPUPS mtip1
RELEASE POPUPS mtip2
RELEASE POPUPS mtip3

 DO CASE 
  CASE m.TipOfView == 1
   DO ViewExp
  CASE m.TipOfView == 2
   DO ViewTalon
  OTHERWISE
 ENDCASE 

RETURN 