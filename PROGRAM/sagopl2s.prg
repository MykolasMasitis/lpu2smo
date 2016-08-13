PROCEDURE SagOpl2s
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ œŒ œŒÀ”/¬Œ«–¿—“”?'+CHR(13)+CHR(10),4+32,'—“¿÷»ŒÕ¿–')=7
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
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF  
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rrid')>0
   IF USED('err')
    USE IN err 
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('err')
    USE IN err 
   ENDIF 
   LOOP 
  ENDIF 
  
  CREATE CURSOR curppl (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  WAIT m.mcod WINDOW NOWAIT 
  SELECT people
  SET RELATION TO recid INTO err
  
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
   m.tip_p = tip_p
   IF m.tip_p<=1
    LOOP 
   ENDIF 
   m.sn_pol = sn_pol
   IF SEEK(m.sn_pol,'curppl')
    LOOP 
   ENDIF 
   INSERT INTO curppl FROM MEMVAR 

   m.dr   = dr
   m.w    = w
   m.vozr = ROUND((m.tdat2 - m.dr)/365.25,2)

   m.m0001 = m.m0001 + IIF(BETWEEN(m.vozr,0,0.99)   and m.w=1, 1, 0)
   m.f0001 = m.f0001 + IIF(BETWEEN(m.vozr,0,0.99)   and m.w=2, 1, 0)
   m.m0104 = m.m0104 + IIF(BETWEEN(m.vozr,1,4.99)   and m.w=1, 1, 0)
   m.f0104 = m.f0104 + IIF(BETWEEN(m.vozr,1,4.99)   and m.w=2, 1, 0)
   m.m0514 = m.m0514 + IIF(BETWEEN(m.vozr,5,14.99)  and m.w=1, 1, 0)
   m.f0514 = m.f0514 + IIF(BETWEEN(m.vozr,5,14.99)  and m.w=2, 1, 0)
   m.m1517 = m.m1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=1, 1, 0)
   m.f1517 = m.f1517 + IIF(BETWEEN(m.vozr,15,17.99) and m.w=2, 1, 0)
   m.m1824 = m.m1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=1, 1, 0)
   m.f1824 = m.f1824 + IIF(BETWEEN(m.vozr,18,24.99) and m.w=2, 1, 0)
   m.m2534 = m.m2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=1, 1, 0)
   m.f2534 = m.f2534 + IIF(BETWEEN(m.vozr,25,34.99) and m.w=2, 1, 0)
   m.m3544 = m.m3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=1, 1, 0)
   m.f3544 = m.f3544 + IIF(BETWEEN(m.vozr,35,44.99) and m.w=2, 1, 0)
   m.m4559 = m.m4559 + IIF(BETWEEN(m.vozr,45,59.99) and m.w=1, 1, 0)
   m.f4554 = m.f4554 + IIF(BETWEEN(m.vozr,45,54.99) and m.w=2, 1, 0)
   m.m6068 = m.m6068 + IIF(BETWEEN(m.vozr,60,68.99) and m.w=1, 1, 0)
   m.f5564 = m.f5564 + IIF(BETWEEN(m.vozr,55,64.99) and m.w=2, 1, 0)
   m.m6999 = m.m6999 + IIF(m.vozr>=69 and m.w=1, 1, 0)
   m.f6599 = m.f6599 + IIF(m.vozr>=65 and m.w=2, 1, 0)

  ENDSCAN 
  
  INSERT INTO cursv FROM MEMVAR 
  
  SET RELATION OFF INTO err
  USE 
  USE IN err
  WAIT CLEAR 
  
  SELECT aisoms

 ENDSCAN 
 USE 
 
 SELECT cursv
 
 m.m0001sv=0
 m.f0001sv=0
 m.m0104sv=0
 m.f0104sv=0
 m.m0514sv=0
 m.f0514sv=0
 m.m1517sv=0
 m.f1517sv=0
 m.m1824sv=0
 m.f1824sv=0
 m.m2534sv=0
 m.f2534sv=0
 m.m3544sv=0
 m.f3544sv=0
 m.m4559sv=0
 m.f4554sv=0
 m.m6068sv=0
 m.f5564sv=0
 m.m6999sv=0
 m.f6599sv=0
 
 SCAN 
  m.m0001sv = m.m0001sv + m0001 
  m.f0001sv = m.f0001sv + f0001 
  m.m0104sv = m.m0104sv + m0104 
  m.f0104sv = m.f0104sv + f0104 
  m.m0514sv = m.m0514sv + m0514 
  m.f0514sv = m.f0514sv + f0514 
  m.m1517sv = m.m1517sv + m1517 
  m.f1517sv = m.f1517sv + f1517 
  m.m1824sv = m.m1824sv + m1824 
  m.f1824sv = m.f1824sv + f1824 
  m.m2534sv = m.m2534sv + m2534 
  m.f2534sv = m.f2534sv + f2534 
  m.m3544sv = m.m3544sv + m3544 
  m.f3544sv = m.f3544sv + f3544 
  m.m4559sv = m.m4559sv + m4559 
  m.f4554sv = m.f4554sv + f4554 
  m.m6068sv = m.m6068sv + m6068 
  m.f5564sv = m.f5564sv + f5564 
  m.m6999sv = m.m6999sv + m6999 
  m.f6599sv = m.f6599sv + f6599 

  m.rsum = m0001+f0001+m0104+f0104+m0514+f0514+m1517+f1517+m1824+f1824+;
           m2534+f2534+m3544+f3544+m4559+f4554+m6068+f5564+m6999+f6599
  m.tsum = m.rsum
  m.ok = IIF(m.tsum=m.rsum, .t., .f.)
  m.diff = m.tsum - m.rsum
  REPLACE ok WITH m.ok, diff WITH m.diff, tsum WITH m.tsum
 ENDSCAN 
 APPEND BLANK 
 GO BOTTOM 

 REPLACE m0001 WITH m.m0001sv, f0001 WITH m.f0001sv, m0104 WITH m.m0104sv, f0104 WITH m.f0104sv, ;
  m0514 WITH m.m0514sv, f0514 WITH m.f0514sv, m1517 WITH m.m1517sv, f1517 WITH m.f1517sv, m1824 WITH m.m1824sv,;
  f1824 WITH m.f1824sv, m2534 WITH m.m2534sv, f2534 WITH m.f2534sv, m3544 WITH m.m3544sv, f3544 WITH m.f3544sv, ;
  m4559 WITH m.m4559sv, f4554 WITH m.f4554sv, m6068 WITH m.m6068sv, f5564 WITH m.f5564sv, m6999 WITH m.m6999sv,;
  f6599 WITH m.f6599sv

 BROWSE 
 COPY TO &pout\&gcperiod\sagopl2
 USE  
 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
RETURN 