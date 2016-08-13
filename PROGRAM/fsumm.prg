FUNCTION FSumm
PARAMETERS Usl, STip, Kol

IF SEEK(Usl, 'Tarif')
 m.norma    = Tarif.n_kd
 m.price    = Tarif.Price
 m.kd_price = ROUND(m.price/m.norma,2)

 IF Tarif.Tip == 'C'
  DO CASE 
   CASE INT(Usl/1000) = 83
    DO CASE
     CASE kol < m.norma
      summa = ROUND(m.kd_price * kol,2)
     CASE kol = m.norma
      summa = m.price
     CASE kol > m.norma and kol <= 30
      summa = round(m.kd_price * kol,2)
     CASE kol > m.norma and kol > 30
      summa = ROUND(m.kd_price * 30,2)
    ENDCASE 

   CASE int(Usl/1000) = 183
    summa = IIF(Kol<=30, Kol*m.price, m.price)

   OTHERWISE 
    IF STip = 'Ä'
     summa = m.price
    ELSE 
     summa = IIF(Kol<m.norma, round(m.kd_price * Kol,2), m.price)
    ENDIF 
  ENDCASE 
 ELSE 

 summa = Kol * m.price

 ENDIF 

 ELSE 
  Summa = 0
 ENDIF
RETURN Summa
