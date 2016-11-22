PROCEDURE StatFN14
 IF MESSAGEBOX('ÕÎÒÈÒÅ ÐÀÑ×ÈÒÀÒÜ ÌÀÊÐÎÏÎÊÀÇÀÒÅËÈ?'+CHR(13)+CHR(10)+;
  ''+CHR(13)+CHR(10),4+32, '')==7
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+gcPeriod+'\stat.dbf')
  CREATE TABLE &pBase\&gcPeriod\stat (lpuid n(4), mcod c(7), cokr c(2), period c(6), ;
   obr_amb1 n(7), pos_amb1 n(7), k_pos1 n(5,2), obr_dom1 n(7), pos_dom1 n(7), usl1 n(7), k_uslobr1 n(5,2), k_uslpos1 n(5,2),;
   st_obr1 n(7), st_pos1 n(7), st_usl1 n(7), paz_dst1 n(7), kd1 n(7), ;
   obr_amb2 n(7), pos_amb2 n(7), k_pos2 n(5,2), obr_dom2 n(7), pos_dom2 n(7), usl2 n(7), k_uslobr2 n(5,2), k_uslpos2 n(5,2),;
   st_obr2 n(7), st_pos2 n(7), st_usl2 n(7), paz_dst2 n(7), kd2 n(7), ;
   usl51 n(7), usl52 n(7), usl53 n(7), usl54 n(7), usl55 n(7), usl56 n(7), ;
   st_paz1 n(7), st_paz2 n(7), amb_paz1 n(7), amb_paz2 n(2))
  INDEX ON mcod TAG mcod 
  USE 
 ENDIF
 
 IF OpenFile("&pBase\&gcPeriod\stat", "stat", "shar", "mcod") > 0
  USE IN aisoms
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\usl_pos', "uslpos", "shar", "cod") > 0
  USE IN aisoms
  USE IN stat 
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\usl_obr', "uslobr", "shar", "cod") > 0
  USE IN aisoms
  USE IN stat 
  USE IN uslpos
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\pos_dom', "posdom", "shar", "cod") > 0
  USE IN aisoms
  USE IN stat 
  USE IN uslpos
  USE IN uslobr
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod") > 0
  USE IN aisoms
  USE IN stat 
  USE IN uslpos
  USE IN uslobr
  USE IN posdom
  RETURN 
 ENDIF 

 SELECT AisOms
 
 SCAN
  m.mcod = mcod
  m.lpuid = lpuid
  m.cokr = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.cokr, '')
  WAIT m.mcod WINDOW NOWAIT 
  
  IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   LOOP 
  ENDIF 
  
  CREATE CURSOR pazdst1 (sn_pol c(25))
  SELECT pazdst1
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  CREATE CURSOR pazdst2 (sn_pol c(25))
  SELECT pazdst2
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol

  CREATE CURSOR ambpaz1 (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  CREATE CURSOR ambpaz2 (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol

  CREATE CURSOR stpaz1 (c_i c(30))
  INDEX ON c_i TAG c_i
  SET ORDER TO c_i

  CREATE CURSOR stpaz2 (c_i c(30))
  INDEX ON c_i TAG c_i
  SET ORDER TO c_i

 SELECT talon 
  m.obr_amb1 = 0 
  m.obr_amb2 = 0 
  m.pos_amb1 = 0 
  m.pos_amb2 = 0 
  m.k_pos1   = 0
  m.k_pos2   = 0
  m.obr_dom1 = 0
  m.obr_dom2 = 0
  m.pos_dom1 = 0
  m.pos_dom2 = 0
  m.st_obr1 = 0
  m.st_pos1 = 0
  m.st_usl1 = 0
  m.st_obr2 = 0
  m.st_pos2 = 0
  m.st_usl2 = 0
  m.usl51 = 0
  m.usl52 = 0
  m.usl53 = 0
  m.usl54 = 0
  m.usl55 = 0
  m.usl56 = 0
  m.k_uslobr1 = 0
  m.k_uslpos1 = 0
  m.k_uslobr2 = 0
  m.k_uslpos2 = 0
  m.usl1 = 0
  m.usl2 = 0
  m.paz_dst1 = 0
  m.kd1 = 0
  m.paz_dst2 = 0
  m.kd2 = 0
  m.st_paz1 = 0
  m.st_paz2 = 0
  m.amb_paz1 = 0
  m.amb_paz2 = 0
  m.st_kd1 = 0
  m.st_kd2 = 0

  SCAN 
   m.sn_pol = sn_pol
   m.c_i = c_i
   m.cod = cod 
   m.k_u = k_u
   m.tip = tip
   m.obr_amb1 = m.obr_amb1 + IIF(SEEK(m.cod, 'uslobr') AND BETWEEN(m.cod,1001,1999), m.k_u, 0)
   m.pos_amb1 = m.pos_amb1 + IIF(SEEK(m.cod, 'uslpos') AND BETWEEN(m.cod,1001,1999), m.k_u, 0)
   m.obr_amb2 = m.obr_amb2 + IIF(SEEK(m.cod, 'uslobr') AND BETWEEN(m.cod,101001,101999), m.k_u, 0)
   m.pos_amb2 = m.pos_amb2 + IIF(SEEK(m.cod, 'uslpos') AND BETWEEN(m.cod,101001,101999), m.k_u, 0)
   m.pos_dom1 = m.pos_dom1 + IIF(SEEK(m.cod, 'posdom') AND BETWEEN(m.cod,1001,1999), m.k_u, 0)
   m.pos_dom2 = m.pos_dom2 + IIF(SEEK(m.cod, 'posdom') AND BETWEEN(m.cod,101001,101999), m.k_u, 0)
   m.st_pos1 = m.st_pos1 + IIF(BETWEEN(m.cod,1211,1254), m.k_u, 0)
   m.st_pos2 = m.st_pos2 + IIF(BETWEEN(m.cod,101191,101200), m.k_u, 0)
   m.kd1 = m.kd1 + IIF(BETWEEN(m.cod,97000,99999), m.k_u, 0)
   m.kd2 = m.kd2 + IIF(BETWEEN(m.cod,197000,199999), m.k_u, 0)
   m.st_kd1 = m.st_kd1 + IIF(!EMPTY(m.tip) AND m.cod<99999, m.k_u, 0)
   m.st_kd2 = m.st_kd2 + IIF(!EMPTY(m.tip) AND m.cod>99999, m.k_u, 0)
   m.tip = tip
   
   IF BETWEEN(m.cod, 0, 60017)
    IF !SEEK(m.sn_pol, 'ambpaz1')
     INSERT INTO ambpaz1 (sn_pol) VALUES (m.sn_pol)
     m.amb_paz1 = m.amb_paz1 + 1
    ENDIF 
   ENDIF 
   
   IF BETWEEN(m.cod, 101001, 160017)
    IF !SEEK(m.sn_pol, 'ambpaz2')
     INSERT INTO ambpaz2 (sn_pol) VALUES (m.sn_pol)
     m.amb_paz2 = m.amb_paz2 + 1
    ENDIF 
   ENDIF 
   
   IF !EMPTY(m.tip) AND m.cod<99999
    IF !SEEK(m.c_i, 'stpaz1')
     INSERT INTO stpaz1 (c_i) VALUES (m.c_i)
     m.st_paz1 = m.st_paz1 + 1
    ENDIF 
   ENDIF 

   IF !EMPTY(m.tip) AND m.cod>99999
    IF !SEEK(m.c_i, 'stpaz2')
     INSERT INTO stpaz2 (c_i) VALUES (m.c_i)
     m.st_paz2 = m.st_paz2 + 1
    ENDIF 
   ENDIF 

   IF BETWEEN(m.cod, 97000, 99999)
    IF !SEEK(m.sn_pol, 'pazdst1')
     INSERT INTO pazdst1 (sn_pol) VALUES (m.sn_pol)
     m.paz_dst1 = m.paz_dst1 + 1
    ENDIF 
   ENDIF 

   IF BETWEEN(m.cod, 197000, 199999)
    IF !SEEK(m.sn_pol, 'pazdst2')
     INSERT INTO pazdst2 (sn_pol) VALUES (m.sn_pol)
     m.paz_dst2 = m.paz_dst2 + 1
    ENDIF 
   ENDIF 

   IF BETWEEN(m.cod,1,60999)
    IF BETWEEN(m.cod,1211,1254) OR BETWEEN(m.cod,9001,9362) OR BETWEEN(m.cod,50001,50324)
     m.st_usl1 = m.st_usl1 + m.k_u
    ELSE 
     m.usl1 = m.usl1 + m.k_u
    ENDIF 
   ENDIF 

   IF BETWEEN(m.cod,101000,160999)
    IF BETWEEN(m.cod,101191,101200) OR BETWEEN(m.cod,109001,109633) OR BETWEEN(m.cod,150625,150636)
     m.st_usl2 = m.st_usl2 + m.k_u
    ELSE 
     m.usl2 = m.usl2 + m.k_u
    ENDIF 
   ENDIF 
   
   m.param = FLOOR(m.cod/1000)
   DO CASE 
    CASE m.param = 51
     m.usl51 = m.usl51 + 1
    CASE m.param = 52
     m.usl52 = m.usl52 + 1
    CASE m.param = 53
     m.usl53 = m.usl53 + 1
    CASE m.param = 54
     m.usl54 = m.usl54 + 1
    CASE m.param = 55
     m.usl55 = m.usl55 + 1
    CASE m.param = 56
     m.usl56 = m.usl56 + 1
   ENDCASE 
  ENDSCAN 
  USE 
  
  m.k_pos1 = IIF(m.obr_amb1>0, m.pos_amb1/m.obr_amb1, 0)
  m.k_pos2 = IIF(m.obr_amb2>0, m.pos_amb2/m.obr_amb2, 0)

  m.k_uslobr1 = IIF(m.obr_amb1>0, m.usl1/m.obr_amb1, 0)
  m.k_uslpos1 = IIF(m.pos_amb1>0, m.usl1/m.pos_amb1, 0)

  m.k_uslobr2 = IIF(m.obr_amb2>0, m.usl2/m.obr_amb2, 0)
  m.k_uslpos2 = IIF(m.pos_amb2>0, m.usl2/m.pos_amb2, 0)

  IF !SEEK(m.mcod, 'stat')
   INSERT INTO stat ;
    (lpuid, mcod, period, cokr, pos_amb1, obr_amb1, k_pos1, pos_amb2, obr_amb2, k_pos2, pos_dom1, pos_dom2, ;
     usl51, usl52, usl53, usl54, usl55, usl56, st_pos1, st_usl1, st_pos2, st_usl2, usl1, usl2, ;
     k_uslobr1, k_uslpos1, k_uslobr2, k_uslpos2, paz_dst1, kd1, paz_dst2, kd2, st_paz1, st_paz2, amb_paz1, amb_paz2, st_kd1, st_kd2) ;
     VALUES ;
    (m.lpuid, m.mcod, m.gcPeriod, m.cokr, m.pos_amb1, m.obr_amb1, m.k_pos1, m.pos_amb2, m.obr_amb2, m.k_pos2, m.pos_dom1, m.pos_dom2, ;
     m.usl51, m.usl52, m.usl53, m.usl54, m.usl55, m.usl56, m.st_pos1, m.st_usl1, m.st_pos2, m.st_usl2, m.usl1, m.usl2, ;
     m.k_uslobr1, m.k_uslpos1, m.k_uslobr2, m.k_uslpos2, m.paz_dst1, m.kd1, m.paz_dst2, m.kd2, m.st_paz1, m.st_paz2, m.amb_paz1, m.amb_paz2, m.st_kd1, m.st_kd2)
  ELSE 
   REPLACE stat.pos_amb1 WITH m.pos_amb1, stat.pos_amb2 WITH m.pos_amb2, stat.k_pos1 WITH m.k_pos1, ;
    stat.obr_amb1 WITH m.obr_amb1, stat.obr_amb2 WITH m.obr_amb2, stat.k_pos2 WITH m.k_pos2, ;
    stat.pos_dom1 WITH m.pos_dom1, stat.pos_dom2 WITH m.pos_dom2, stat.usl1 WITH m.usl1, ;
    stat.usl2 WITH m.usl2, stat.k_uslobr1 WITH m.k_uslobr1, stat.k_uslpos1 WITH m.k_uslpos1, ;
    stat.k_uslobr2 WITH m.k_uslobr2, stat.k_uslpos2 WITH m.k_uslpos2, ;
    stat.paz_dst1 WITH m.paz_dst1, stat.kd1 WITH m.kd1, stat.paz_dst2 WITH m.paz_dst2, stat.kd2 WITH m.kd2,;
    stat.st_paz1 WITH m.st_paz1, stat.st_paz2 WITH m.st_paz2, stat.amb_paz1 WITH m.amb_paz1, stat.amb_paz2 WITH m.amb_paz2,;
    stat.st_kd1 WITH m.st_kd1, stat.st_kd2 WITH m.st_kd2 IN stat
  ENDIF 

  SELECT aisoms 

 ENDSCAN 

 WAIT CLEAR 
 USE 
 USE IN stat 
 USE IN uslpos
 USE IN uslobr
 USE IN posdom
 USE IN sprlpu
 
RETURN 