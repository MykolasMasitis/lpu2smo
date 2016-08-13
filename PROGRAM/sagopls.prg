PROCEDURE SagOpls
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ œŒ œŒÀ”/¬Œ«–¿—“”?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
 ENDIF 
 
 CREATE CURSOR cursv (mcod c(7), lpuid n(4), ok l, diff n(11,2),m0001 n(11,2), f0001 n(11,2), m0104 n(11,2), f0104 n(11,2), m0514 n(11,2), f0514 n(11,2),;
  m1517 n(11,2), f1517 n(11,2), m1824 n(11,2), f1824 n(11,2), m2534 n(11,2), f2534 n(11,2), m3544 n(11,2), f3544 n(11,2), m4559 n(11,2), f4554 n(11,2),;
  m6068 n(11,2), f5564 n(11,2), m6999 n(11,2), f6599 n(11,2), tsum n(11,2))

 CREATE CURSOR cursv1 (mcod c(7), lpuid n(4), ok l, diff n(11,2),m0001 n(11,2), f0001 n(11,2), m0104 n(11,2), f0104 n(11,2), m0514 n(11,2), f0514 n(11,2),;
  m1517 n(11,2), f1517 n(11,2), m1824 n(11,2), f1824 n(11,2), m2534 n(11,2), f2534 n(11,2), m3544 n(11,2), f3544 n(11,2), m4559 n(11,2), f4554 n(11,2),;
  m6068 n(11,2), f5564 n(11,2), m6999 n(11,2), f6599 n(11,2), tsum n(11,2))
 CREATE CURSOR cursv2 (mcod c(7), lpuid n(4), ok l, diff n(11,2),m0001 n(11,2), f0001 n(11,2), m0104 n(11,2), f0104 n(11,2), m0514 n(11,2), f0514 n(11,2),;
  m1517 n(11,2), f1517 n(11,2), m1824 n(11,2), f1824 n(11,2), m2534 n(11,2), f2534 n(11,2), m3544 n(11,2), f3544 n(11,2), m4559 n(11,2), f4554 n(11,2),;
  m6068 n(11,2), f5564 n(11,2), m6999 n(11,2), f6599 n(11,2), tsum n(11,2))
 CREATE CURSOR cursv3 (mcod c(7), lpuid n(4), ok l, diff n(11,2),m0001 n(11,2), f0001 n(11,2), m0104 n(11,2), f0104 n(11,2), m0514 n(11,2), f0514 n(11,2),;
  m1517 n(11,2), f1517 n(11,2), m1824 n(11,2), f1824 n(11,2), m2534 n(11,2), f2534 n(11,2), m3544 n(11,2), f3544 n(11,2), m4559 n(11,2), f4554 n(11,2),;
  m6068 n(11,2), f5564 n(11,2), m6999 n(11,2), f6599 n(11,2), tsum n(11,2))

 SELECT aisoms
 SCAN
  m.mcod  = mcod 
  m.lpuid = lpuid
  IF VAL(SUBSTR(m.mcod,3,2))<=40
   LOOP 
  ENDIF  
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF  
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF  
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('err')
    USE IN err 
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('err')
    USE IN err 
   ENDIF 
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  SELECT talon
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO err ADDITIVE 
  
  m.m0001 = 0
  m.f0001 = 0
  m.m0104 = 0
  m.f0104 = 0
  m.m0514 = 0
  m.f0514 = 0
  m.m1517 = 0
  m.f1517 = 0
  m.m1824 = 0
  m.f1824 = 0
  m.m2534 = 0
  m.f2534 = 0
  m.m3544 = 0
  m.f3544 = 0
  m.m4559 = 0
  m.f4554 = 0
  m.m6068 = 0
  m.f5564 = 0
  m.m6999 = 0
  m.f6599 = 0
  
  m.tsum = 0
  
  m.ok   = .f.

  SCAN 
   IF !EMPTY(err.rid)
    LOOP 
   ENDIF 
   m.cod = cod
   IF !(IsMes(m.cod) OR IsVMP(m.cod) OR IsKDS(m.cod))
    LOOP 
   ENDIF 

   m.dr   = people.dr
   m.w    = people.w
   m.vozr = ROUND((m.tdat2 - m.dr)/365.25,2)
   m.s_all = s_all 

   m.m0001 = m.m0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=1, m.s_all, 0)
   m.f0001 = m.f0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=2, m.s_all, 0)
   m.m0104 = m.m0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=1, m.s_all, 0)
   m.f0104 = m.f0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=2, m.s_all, 0)
   m.m0514 = m.m0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=1, m.s_all, 0)
   m.f0514 = m.f0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=2, m.s_all, 0)
   m.m1517 = m.m1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=1, m.s_all, 0)
   m.f1517 = m.f1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=2, m.s_all, 0)
   m.m1824 = m.m1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=1, m.s_all, 0)
   m.f1824 = m.f1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=2, m.s_all, 0)
   m.m2534 = m.m2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=1, m.s_all, 0)
   m.f2534 = m.f2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=2, m.s_all, 0)
   m.m3544 = m.m3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=1, m.s_all, 0)
   m.f3544 = m.f3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=2, m.s_all, 0)
   m.m4559 = m.m4559 + IIF(BETWEEN(m.vozr,45,59.99) and m.w=1, m.s_all, 0)
   m.f4554 = m.f4554 + IIF(BETWEEN(m.vozr,45,54.99) and m.w=2, m.s_all, 0)
   m.m6068 = m.m6068 + IIF(BETWEEN(m.vozr,60,68.99) and m.w=1, m.s_all, 0)
   m.f5564 = m.f5564 + IIF(BETWEEN(m.vozr,55,64.99) and m.w=2, m.s_all, 0)
   m.m6999 = m.m6999 + IIF(m.vozr>=69 and m.w=1, m.s_all, 0)
   m.f6599 = m.f6599 + IIF(m.vozr>=65 and m.w=2, m.s_all, 0)

   m.tsum = m.tsum + s_all

  ENDSCAN 
  
  INSERT INTO cursv FROM MEMVAR 
  
  m.m0001 = 0
  m.f0001 = 0
  m.m0104 = 0
  m.f0104 = 0
  m.m0514 = 0
  m.f0514 = 0
  m.m1517 = 0
  m.f1517 = 0
  m.m1824 = 0
  m.f1824 = 0
  m.m2534 = 0
  m.f2534 = 0
  m.m3544 = 0
  m.f3544 = 0
  m.m4559 = 0
  m.f4554 = 0
  m.m6068 = 0
  m.f5564 = 0
  m.m6999 = 0
  m.f6599 = 0
  
  m.tsum = 0
  
  m.ok   = .f.

  SCAN 
   IF !EMPTY(err.rid)
    LOOP 
   ENDIF
   m.cod = cod  
   IF !IsUsl(m.cod)
    LOOP 
   ENDIF 

   m.dr   = people.dr
   m.w    = people.w
   m.vozr = ROUND((m.tdat2 - m.dr)/365.25,2)
   m.s_all = s_all 

   m.m0001 = m.m0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=1, m.s_all, 0)
   m.f0001 = m.f0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=2, m.s_all, 0)
   m.m0104 = m.m0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=1, m.s_all, 0)
   m.f0104 = m.f0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=2, m.s_all, 0)
   m.m0514 = m.m0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=1, m.s_all, 0)
   m.f0514 = m.f0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=2, m.s_all, 0)
   m.m1517 = m.m1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=1, m.s_all, 0)
   m.f1517 = m.f1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=2, m.s_all, 0)
   m.m1824 = m.m1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=1, m.s_all, 0)
   m.f1824 = m.f1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=2, m.s_all, 0)
   m.m2534 = m.m2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=1, m.s_all, 0)
   m.f2534 = m.f2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=2, m.s_all, 0)
   m.m3544 = m.m3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=1, m.s_all, 0)
   m.f3544 = m.f3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=2, m.s_all, 0)
   m.m4559 = m.m4559 + IIF(BETWEEN(m.vozr,45,59.99) and m.w=1, m.s_all, 0)
   m.f4554 = m.f4554 + IIF(BETWEEN(m.vozr,45,54.99) and m.w=2, m.s_all, 0)
   m.m6068 = m.m6068 + IIF(BETWEEN(m.vozr,60,68.99) and m.w=1, m.s_all, 0)
   m.f5564 = m.f5564 + IIF(BETWEEN(m.vozr,55,64.99) and m.w=2, m.s_all, 0)
   m.m6999 = m.m6999 + IIF(m.vozr>=69 and m.w=1, m.s_all, 0)
   m.f6599 = m.f6599 + IIF(m.vozr>=65 and m.w=2, m.s_all, 0)

   m.tsum = m.tsum + s_all

  ENDSCAN 
  
  INSERT INTO cursv1 FROM MEMVAR 

  m.m0001 = 0
  m.f0001 = 0
  m.m0104 = 0
  m.f0104 = 0
  m.m0514 = 0
  m.f0514 = 0
  m.m1517 = 0
  m.f1517 = 0
  m.m1824 = 0
  m.f1824 = 0
  m.m2534 = 0
  m.f2534 = 0
  m.m3544 = 0
  m.f3544 = 0
  m.m4559 = 0
  m.f4554 = 0
  m.m6068 = 0
  m.f5564 = 0
  m.m6999 = 0
  m.f6599 = 0
  
  m.tsum = 0
  
  m.ok   = .f.

  SCAN 
   IF !EMPTY(err.rid)
    LOOP 
   ENDIF
   m.cod = cod  
   IF !IsKd(m.cod)
    LOOP 
   ENDIF 

   m.dr   = people.dr
   m.w    = people.w
   m.vozr = ROUND((m.tdat2 - m.dr)/365.25,2)
   m.s_all = s_all 

   m.m0001 = m.m0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=1, m.s_all, 0)
   m.f0001 = m.f0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=2, m.s_all, 0)
   m.m0104 = m.m0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=1, m.s_all, 0)
   m.f0104 = m.f0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=2, m.s_all, 0)
   m.m0514 = m.m0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=1, m.s_all, 0)
   m.f0514 = m.f0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=2, m.s_all, 0)
   m.m1517 = m.m1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=1, m.s_all, 0)
   m.f1517 = m.f1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=2, m.s_all, 0)
   m.m1824 = m.m1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=1, m.s_all, 0)
   m.f1824 = m.f1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=2, m.s_all, 0)
   m.m2534 = m.m2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=1, m.s_all, 0)
   m.f2534 = m.f2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=2, m.s_all, 0)
   m.m3544 = m.m3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=1, m.s_all, 0)
   m.f3544 = m.f3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=2, m.s_all, 0)
   m.m4559 = m.m4559 + IIF(BETWEEN(m.vozr,45,59.99) and m.w=1, m.s_all, 0)
   m.f4554 = m.f4554 + IIF(BETWEEN(m.vozr,45,54.99) and m.w=2, m.s_all, 0)
   m.m6068 = m.m6068 + IIF(BETWEEN(m.vozr,60,68.99) and m.w=1, m.s_all, 0)
   m.f5564 = m.f5564 + IIF(BETWEEN(m.vozr,55,64.99) and m.w=2, m.s_all, 0)
   m.m6999 = m.m6999 + IIF(m.vozr>=69 and m.w=1, m.s_all, 0)
   m.f6599 = m.f6599 + IIF(m.vozr>=65 and m.w=2, m.s_all, 0)

   m.tsum = m.tsum + s_all

  ENDSCAN 
  
  INSERT INTO cursv2 FROM MEMVAR 

  m.m0001 = 0
  m.f0001 = 0
  m.m0104 = 0
  m.f0104 = 0
  m.m0514 = 0
  m.f0514 = 0
  m.m1517 = 0
  m.f1517 = 0
  m.m1824 = 0
  m.f1824 = 0
  m.m2534 = 0
  m.f2534 = 0
  m.m3544 = 0
  m.f3544 = 0
  m.m4559 = 0
  m.f4554 = 0
  m.m6068 = 0
  m.f5564 = 0
  m.m6999 = 0
  m.f6599 = 0
  
  m.tsum = 0
  
  m.ok   = .f.

  SCAN 
   IF !EMPTY(err.rid)
    LOOP 
   ENDIF
   m.cod = cod  
   IF !(IsMes(m.cod) OR IsVmp(m.cod))
    LOOP 
   ENDIF 

   m.dr   = people.dr
   m.w    = people.w
   m.vozr = ROUND((m.tdat2 - m.dr)/365.25,2)
   m.s_all = s_all 

   m.m0001 = m.m0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=1, m.s_all, 0)
   m.f0001 = m.f0001 + IIF(BETWEEN(m.vozr,0,0.99) and m.w=2, m.s_all, 0)
   m.m0104 = m.m0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=1, m.s_all, 0)
   m.f0104 = m.f0104 + IIF(BETWEEN(m.vozr,1,4.99) and m.w=2, m.s_all, 0)
   m.m0514 = m.m0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=1, m.s_all, 0)
   m.f0514 = m.f0514 + IIF(BETWEEN(m.vozr,5,14.99) and m.w=2, m.s_all, 0)
   m.m1517 = m.m1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=1, m.s_all, 0)
   m.f1517 = m.f1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=2, m.s_all, 0)
   m.m1824 = m.m1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=1, m.s_all, 0)
   m.f1824 = m.f1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=2, m.s_all, 0)
   m.m2534 = m.m2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=1, m.s_all, 0)
   m.f2534 = m.f2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=2, m.s_all, 0)
   m.m3544 = m.m3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=1, m.s_all, 0)
   m.f3544 = m.f3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=2, m.s_all, 0)
   m.m4559 = m.m4559 + IIF(BETWEEN(m.vozr,45,59.99) and m.w=1, m.s_all, 0)
   m.f4554 = m.f4554 + IIF(BETWEEN(m.vozr,45,54.99) and m.w=2, m.s_all, 0)
   m.m6068 = m.m6068 + IIF(BETWEEN(m.vozr,60,68.99) and m.w=1, m.s_all, 0)
   m.f5564 = m.f5564 + IIF(BETWEEN(m.vozr,55,64.99) and m.w=2, m.s_all, 0)
   m.m6999 = m.m6999 + IIF(m.vozr>=69 and m.w=1, m.s_all, 0)
   m.f6599 = m.f6599 + IIF(m.vozr>=65 and m.w=2, m.s_all, 0)

   m.tsum = m.tsum + s_all

  ENDSCAN 
  
  INSERT INTO cursv3 FROM MEMVAR 

  SET RELATION OFF INTO err
  SET RELATION OFF INTO people
  USE 
  USE IN err
  USE IN people 
  WAIT CLEAR 
  
  SELECT aisoms

 ENDSCAN 
 USE 
 
 SELECT cursv
 SCAN 
  m.tsum = tsum
  m.rsum = m0001+f0001+m0104+f0104+m0514+f0514+m1517+f1517+m1824+f1824+;
           m2534+f2534+m3544+f3544+m4559+f4554+m6068+f5564+m6999+f6599
  m.ok = IIF(m.tsum=m.rsum, .t., .f.)
  m.diff = m.tsum - m.rsum
  REPLACE ok WITH m.ok, diff WITH m.diff
 ENDSCAN 
 BROWSE 
 COPY TO &pout\&gcperiod\sagopl
 USE  

 SELECT cursv1
 SCAN 
  m.tsum = tsum
  m.rsum = m0001+f0001+m0104+f0104+m0514+f0514+m1517+f1517+m1824+f1824+;
           m2534+f2534+m3544+f3544+m4559+f4554+m6068+f5564+m6999+f6599
  m.ok = IIF(m.tsum=m.rsum, .t., .f.)
  m.diff = m.tsum - m.rsum
  REPLACE ok WITH m.ok, diff WITH m.diff
 ENDSCAN 
 BROWSE 
 COPY TO &pout\&gcperiod\sagopl1
 USE  
 SELECT cursv2
 SCAN 
  m.tsum = tsum
  m.rsum = m0001+f0001+m0104+f0104+m0514+f0514+m1517+f1517+m1824+f1824+;
           m2534+f2534+m3544+f3544+m4559+f4554+m6068+f5564+m6999+f6599
  m.ok = IIF(m.tsum=m.rsum, .t., .f.)
  m.diff = m.tsum - m.rsum
  REPLACE ok WITH m.ok, diff WITH m.diff
 ENDSCAN 
 BROWSE 
 COPY TO &pout\&gcperiod\sagopl2
 USE  
 SELECT cursv3
 SCAN 
  m.tsum = tsum
  m.rsum = m0001+f0001+m0104+f0104+m0514+f0514+m1517+f1517+m1824+f1824+;
           m2534+f2534+m3544+f3544+m4559+f4554+m6068+f5564+m6999+f6599
  m.ok = IIF(m.tsum=m.rsum, .t., .f.)
  m.diff = m.tsum - m.rsum
  REPLACE ok WITH m.ok, diff WITH m.diff
 ENDSCAN 
 BROWSE 
 COPY TO &pout\&gcperiod\sagopl3
 USE  
 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
RETURN 