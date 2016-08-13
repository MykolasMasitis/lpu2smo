PROCEDURE selcrosss
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÎÁÐÀÒÜ ÏÅÐÅÑÅ×ÅÍÈß?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR curgosps (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(25), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, ds c(6), dss c(3), k_u n(5),;
  otd c(4), pcod c(10))
 INDEX on c_i TAG c_i
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO c_i 

 CREATE CURSOR curpolks (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(25), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), cod n(6), d_u d, ds c(6), k_u n(5),;
  otd c(4), pcod c(10))
 INDEX on c_i TAG c_i
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO c_i 

 FOR lnmonth=1 TO 12
  m.lcperiod = STR(tYear,4)+PADL(lnmonth,2,'0')
  m.lpath = pbase+'\'+m.lcperiod
  IF !fso.FolderExists(m.lpath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lpath+'\aisoms.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.lcperiod+'...' WINDOW NOWAIT 
  =selgosps(m.lpath)
  WAIT CLEAR 

 NEXT 
 
 WAIT "ÎÔÎÐÌËÅÍÈÅ ÐÅÇÓËÜÒÀÒÎÂ..." WINDOW NOWAIT 

 outfile = pmee+'\slcrosses'
 CREATE TABLE &outfile (period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(25), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, cod n(6), ds c(6), k_u n(5),;
  otd c(4), pcod c(10), d_pos d, d_vip d, gosp c(7))
 USE 

 =OpenFile(outfile, 'outfl', 'shar')
 
 SELECT curgosps
 SET ORDER TO sn_pol

 SELECT curpolks
 SET ORDER TO sn_pol
 SCAN 
  SCATTER MEMVAR 
  IF !SEEK(m.sn_pol, 'curgosps')
   LOOP 
  ENDIF 

  SELECT curgosps
  DO WHILE sn_pol=m.sn_pol
   IF !BETWEEN(m.d_u, d_u-k_u+1, d_u-1)
    SKIP 
    LOOP 
   ENDIF 

   m.d_pos = d_u-k_u
   m.d_vip = d_u
   m.gosp  = mcod

   INSERT INTO outfl FROM MEMVAR 

   SKIP 
  ENDDO 
  SELECT curpolks

 ENDSCAN 

 USE IN curgosps
 USE IN curpolks
 USE IN outfl
 
 WAIT CLEAR 

 MESSAGEBOX('ÃÎÒÎÂÎ!',0+64,'')

RETURN 

FUNCTION selgosps(m.lpath)
 PRIVATE m.llcpath
 m.llcpath = m.lpath
 IF OpenFile(m.llcpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 SCAN 
  m.lpuid = lpuid
  m.mcod = mcod
  IF !fso.FolderExists(m.llcpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(m.llcpath+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llcpath+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('error')
    USE IN error
   ENDIF 
   LOOP 
  ENDIF 
 
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO error ADDITIVE 
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 

   m.sn_pol = sn_pol
   m.c_i    = c_i
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.dr     = people.dr 
   m.w      = people.w
   m.d_u    = d_u
   m.cod    = cod
   m.ds     = ds
   m.dss    = LEFT(ds,3)
   m.k_u    = k_u 
   m.otd    = otd
   m.pcod   = pcod
   
   IF !EMPTY(tip)

    IF !SEEK(m.c_i, 'curgosps')
     INSERT INTO curgosps (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,d_u,k_u,ds,dss,otd,pcod) VALUES ;
      (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,m.d_u,m.k_u,m.ds,m.dss,m.otd,m.pcod) 
    ELSE 
      m.ok_u = curgosps.k_u
      m.nk_u = m.ok_u + m.k_u
     IF m.d_u > curgosps.d_u
      UPDATE curgosps SET d_u=m.d_u, ds=m.ds, dss=m.dss, k_u=m.nk_u WHERE c_i=m.c_i
     ELSE 
      UPDATE curgosps SET k_u=m.nk_u WHERE c_i=m.c_i
     ENDIF 
    ENDIF 
   
   ELSE 

     INSERT INTO curpolks (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,d_u,k_u,ds,cod,otd,pcod) VALUES ;
      (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,m.d_u,m.k_u,m.ds,m.cod,m.otd,m.pcod) 
   
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO error
  SET RELATION OFF INTO people
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('error')
   USE IN error
  ENDIF 
 
  SELECT aisoms

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 

FUNCTION IsGosp(lcmcod)
 m.lnlputip = INT(VAL(SUBSTR(lcmcod,3,2)))
RETURN IIF(BETWEEN(m.lnlputip,40,67), .t., .f.)