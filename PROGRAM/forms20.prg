PROCEDURE FormS20
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ'+CHR(13)+CHR(10)+'ÑÂÎÄÍÓÞ ÂÅÄÎÌÎÑÒÜ ÇÀ ÌÅÑßÖ?',4+32,'SOVITA')=7
  RETURN
 ENDIF 
 IF !fso.FileExists(pTempl+'\FormS20.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ÎÒ×ÅÒÀ'+CHR(13)+CHR(10)+UPPER(pTempl+'\FormS20.xls'),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÏÅÐÈÎÄÀ!',0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË AISOMS.DBF!',0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  USE IN aisoms
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\FormS20.dbf')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\FormS20.dbf')
 ENDIF 
 
 CREATE CURSOR curdata (nrec i AUTOINC, lpuid n(4), mcod c(7), lpuname c(40), krank n(7), paz_dst n(7), paz_st n(7),;
  ambplmee n(5), dstplmee n(5), stplmee n(5), ambplekmp n(5), dstplekmp n(5), stplekmp n(5), ndeads n(5), ngspdbls n(5),;
  nobrdbls n(5), sandlong n(5), uniks n(5))
 INDEX on nrec TAG nrec
 SET ORDER TO nrec 
 
 CREATE CURSOR curgosps (period c(7), mcod c(7), sn_pol c(25), c_i c(30), d_u d, ds c(6), dss c(3), k_u n(5), otd c(4))
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON sn_pol + dss TAG unik
 SET ORDER TO c_i 
 
 CREATE CURSOR curobr (period c(7), mcod c(7), sn_pol c(25), d_u d, cod n(6), ds c(6))
 INDEX ON sn_pol TAG sn_pol
 INDEX ON sn_pol+ds TAG unik
 SET ORDER TO unik
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\curobr.dbf') && Îòáîð äàííûõ äëÿ ïîâòîðíûõ îáðàùåíèé!
  WAIT "ÎÒÁÎÐ ÏÅÐÈÎÄÎÂ..." WINDOW NOWAIT 
  FOR m.nmm=1 TO 1
   m.lcperiod = LEFT(DTOS(GOMONTH(m.tdat2,-m.nmm)),6)
   IF !fso.FolderExists(pBase+'\'+m.lcperiod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+m.lcperiod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+m.lcperiod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
 
   SELECT talon 
   SCAN
    m.cod = cod  
    IF !IsObr(m.cod)
     LOOP 
    ENDIF 
  
    m.mcod   = mcod
    m.sn_pol = sn_pol
    m.d_u    = d_u
    m.ds     = ds
    
    m.unik = m.sn_pol+m.ds
   
    IF !SEEK(m.unik, 'curobr')
     INSERT INTO curobr (period,mcod,sn_pol,d_u,cod,ds) VALUES ;
      (m.lcperiod,m.mcod,m.sn_pol,m.d_u,m.cod,m.ds) 
     IF m.d_u > curobr.d_u
      UPDATE curobr SET d_u=m.d_u WHERE sn_pol+ds = m.unik
     ENDIF 
    ENDIF 
  
   ENDSCAN 
   USE IN talon 

  ENDFOR 
  SELECT curobr
  COPY TO &pbase\&gcperiod\curobr CDX 
  WAIT CLEAR 
 ELSE 
  IF OpenFile(pbase+'\'+m.gcperiod+'\curobr', 'cobr', 'shar')<=0
   SELECT cobr
   SCAN 
    SCATTER MEMVAR 
    INSERT INTO curobr FROM MEMVAR 
   ENDSCAN 
   USE IN cobr
  ENDIF 
 ENDIF 

 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\curgosps.dbf') && Îòáîð äàííûõ äëÿ ïîâòîðíûõ ãîñïèòàëèçàöèé
  WAIT "ÎÒÁÎÐ ÏÅÐÈÎÄÎÂ..." WINDOW NOWAIT 
  FOR m.nmm=1 TO 3
   m.lcperiod = LEFT(DTOS(GOMONTH(m.tdat2,-m.nmm)),6)
   IF !fso.FolderExists(pBase+'\'+m.lcperiod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+m.lcperiod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+m.lcperiod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
 
   SELECT talon 
   SCAN 
    IF EMPTY(tip)
     LOOP 
    ENDIF 
  
    m.mcod   = mcod
    m.sn_pol = sn_pol
    m.c_i    = c_i
    m.d_u    = d_u
    m.ds     = ds
    m.dss    = LEFT(ds,3)
    m.k_u    = k_u 
    m.otd    = otd
    m.pcod   = pcod
   
    IF !SEEK(m.c_i, 'curgosps')
     INSERT INTO curgosps (period,mcod,sn_pol,c_i,d_u,k_u,ds,dss,otd) VALUES ;
      (m.lcperiod,m.mcod,m.sn_pol,m.c_i,m.d_u,m.k_u,m.ds,m.dss,m.otd) 
    ELSE 
     m.ok_u = curgosps.k_u
     m.nk_u = m.ok_u + m.k_u
     IF m.d_u > curgosps.d_u
      UPDATE curgosps SET d_u=m.d_u, ds=m.ds, dss=m.dss, k_u=m.nk_u WHERE c_i=m.c_i
     ELSE 
      UPDATE curgosps SET k_u=m.nk_u WHERE c_i=m.c_i
     ENDIF 
    ENDIF 
  
   ENDSCAN 
   USE IN talon 

  ENDFOR 
  SELECT curgosps
  COPY TO &pbase\&gcperiod\curgosps CDX 
  SET ORDER TO unik
  WAIT CLEAR 
 ELSE 
  IF OpenFile(pbase+'\'+m.gcperiod+'\curgosps', 'cgosps', 'shar')<=0
   SELECT cgosps
   SCAN 
    SCATTER MEMVAR 
    INSERT INTO curgosps FROM MEMVAR 
   ENDSCAN 
   USE IN cgosps 
  ENDIF 
 ENDIF 

 SELECT curgosps
* COPY TO &pbase\&gcperiod\curgosps CDX 
 SET ORDER TO unik

 CREATE CURSOR curdbls (mcod c(7), sn_pol c(25), c_i c(30))
 
 SELECT aisoms
 SET RELATION TO lpuid INTO sprlpu
 SCAN 
  m.lpuid   = lpuid
  m.mcod    = mcod
  m.lpuname = sprlpu.name
  m.krank   = krank
  m.paz_dst = paz_dst
  m.paz_st  = paz_st + paz_vmp
  m.ambplmee  = ROUND(0.008*m.krank,0)
  m.dstplmee  = ROUND(0.08*m.paz_dst,0)
  m.stplmee   = ROUND(0.08*m.paz_st,0)
  m.ambplekmp = ROUND(0.005*m.krank,0)
  m.dstplekmp = ROUND(0.03*m.paz_dst,0)
  m.stplekmp  = ROUND(0.05*m.paz_st,0)
  m.sandlong = 0
  m.ndeads   = 0 
  m.ngspdbls = 0
  m.nobrdbls = 0
*  m.uniks    = 0 

  INSERT INTO curdata FROM MEMVAR 
  
  IF INT(VAL(SUBSTR(m.mcod,3,2)))<41
   LOOP 
  ENDIF 
  
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'serror', 'shar', 'rid')>0
   USE IN talon 
   IF USED('serror')
    USE IN serror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  CREATE CURSOR curss (c_i c(30))
  INDEX on c_i TAG c_i
  SET ORDER TO c_i
  SELECT talon 
  SET RELATION TO recid INTO serror
  SET RELATION TO PADR(sn_pol,25) + LEFT(ds,3) INTO curgosps ADDITIVE 
  SET RELATION TO PADR(sn_pol,25) + ds INTO curobr ADDITIVE 
  SCAN 
   IF !EMPTY(serror.rid)
    LOOP 
   ENDIF 
   m.c_i    = c_i
   m.sn_pol = sn_pol
   IF !EMPTY(Tip) AND !BETWEEN(ROUND(k_u/n_kd,0),0.5,1.5) AND !INLIST(FLOOR(cod/1000), 83, 183) AND !IsVmp(cod)
    m.sandlong = m.sandlong + 1
   ENDIF 
   IF Tip='5'
    m.ndeads = m.ndeads + 1
   ENDIF 
   IF !EMPTY(Tip) AND !EMPTY(curgosps.c_i)
    IF !SEEK(m.c_i, 'curss')
     INSERT INTO curdbls FROM MEMVAR 
     INSERT INTO curss FROM MEMVAR 
     m.ngspdbls = m.ngspdbls + 1
    ENDIF 
   ENDIF 
   IF IsObr(cod) AND !EMPTY(curobr.sn_pol)
    m.nobrdbls = m.nobrdbls + 1
   ENDIF 
*   IF INLIST(Tip,'8','9')
*    m.uniks = m.uniks + 1
*   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO serror
  SET RELATION OFF INTO curgosps
  SET RELATION OFF INTO curobr
  USE IN talon 
  USE IN serror
  USE IN curss
  SELECT aisoms 
  
  IF m.sandlong>0
   UPDATE curdata SET sandlong=m.sandlong WHERE mcod=m.mcod
  ENDIF 
  IF m.ndeads>0
   UPDATE curdata SET ndeads=m.ndeads WHERE mcod=m.mcod
  ENDIF 
  IF m.ngspdbls>0
   UPDATE curdata SET ngspdbls=m.ngspdbls WHERE mcod=m.mcod
  ENDIF 
  IF m.nobrdbls>0
   UPDATE curdata SET nobrdbls=m.nobrdbls WHERE mcod=m.mcod
  ENDIF 
*  IF m.uniks>0
*   UPDATE curdata SET uniks=m.uniks WHERE mcod=m.mcod
*  ENDIF 
  
  WAIT CLEAR 
  
 ENDSCAN 
 SET RELATION OFF INTO sprlpu
 USE IN aisoms
 USE IN sprlpu
 
 USE IN curgosps
 SELECT curdbls 
 COPY TO &pbase\&gcperiod\curdbls_

 SELECT curdata 
 COPY TO &pbase\&gcperiod\FormS20

 m.llResult = X_Report(pTempl+'\FormS20.xls', pBase+'\'+m.gcperiod+'\FormS20.xls', .T.)

 USE 
 
RETURN 