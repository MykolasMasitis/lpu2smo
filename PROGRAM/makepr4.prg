PROCEDURE MakePr4
 IF MESSAGEBOX(CHR(13)+CHR(10)+'СФОРМИРОВАТЬ ПРИЛОЖЕНИЕ 4?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 m.lPath = m.pbase+'\'+m.gcperiod
 IF !fso.FileExists(m.lPath+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ AISOMS.DBF'+CHR(13)+CHR(10),0+16,m.lpath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\pilot.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ PILOT.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(m.lPath+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\lputpn', 'lputpn', 'shar', 'lpu_id')>0
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')>0
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 WAIT "ПРОВЕРКА ФАЙЛА PILOT..." WINDOW NOWAIT 
 SELECT pilot
 SCAN 
  m.lpu_id = lpu_id
  IF SEEK(m.lpu_id, 'aisoms', 'lpuid')
   m.nmcod = aisoms.mcod
  ELSE 
   m.nmcod = ''
   MESSAGEBOX('ВНИМАНИЕ!'+CHR(13)+CHR(10)+'lpu_id='+STR(m.lpu_id,4)+' ОТСУТСТВУЕТ В ФАЙЛЕ AISOMS.DBF!'+CHR(13)+CHR(10),;
    0+16,'')
  ENDIF 
  
  DO CASE 
   CASE EMPTY(m.nmcod)
    MESSAGEBOX('ВНИМАНИЕ!'+CHR(13)+CHR(10)+'mcod='+mcod+' ОТСУТСТВУЕТ В ФАЙЛЕ AISOMS.DBF!'+CHR(13)+CHR(10),0+64,'')
   CASE mcod != m.nmcod
    SET STEP ON 
    MESSAGEBOX('MCOD '+mcod+' ЗАМЕНЕН НА '+m.nmcod+'!'+CHR(13)+CHR(10),0+64,'')
    REPLACE mcod WITH m.nmcod
   OTHERWISE 
  ENDCASE 
  
 ENDSCAN 
 WAIT CLEAR 
 
 m.lcperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 
 CREATE CURSOR curpr4 (lpuid n(4), mcod c(7), adnorm n(11,2), chnorm n(11,2), adults n(6), childs n(6),;
  adsum n(11,2), chsum n(11,2), pazval n(6), finval n(13,2), paz_all n(6), s_all n(11,2), s_pred_pf n(11,2), ;
  paz_empty n(6), s_empty n(11,2), ;
  paz_own n(6), s_own n(11,2), paz_guests n(6), s_guests n(11,2), paz_others n(6), s_others n(11,2),;
  paz_npilot n(6), s_npilot n(11,2), paz_bad n(6), s_bad n(11,2), s_kompl n(11,2), s_dst n(11,2))

 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 m.adnorm = 0
 m.chnorm = 0
 IF fso.FileExists(pcommon+'\pnorm.dbf')
  IF OpenFile(pcommon+'\pnorm.dbf', 'pnorm', 'shar', 'period')<=0
   SELECT pnorm 
   IF SEEK(m.gcperiod, 'pnorm')
   ELSE 
    GO BOTTOM 
   ENDIF 
   SCATTER FIELDS EXCEPT period MEMVAR 
*   m.adnorm = adnorm
*   m.chnorm = chnorm
   USE IN pnorm
  ENDIF 
 ENDIF 
 
 WAIT "СОЗДАНИЕ СПРАВОЧНИКА..." WINDOW NOWAIT 
 SELECT aisoms
 SCAN 
  m.sumok = s_pred - sum_flk
  m.mcod = mcod 
  m.lpuid = lpuid
  IF !SEEK(m.lpuid, 'pilot')
*   IF !SEEK(m.lpuid, 'horlpu')
*    LOOP 
*   ENDIF 
   LOOP 
  ENDIF 
*  IF m.sumok<=0
*   LOOP 
*  ENDIF 
  IF !SEEK(m.mcod, 'curpr4')
   INSERT INTO curpr4 (lpuid, mcod) VALUES (m.lpuid, m.mcod)
  ENDIF 
 ENDSCAN 
 WAIT CLEAR 
 
 SCAN 
  m.mcod = mcod 
  m.lpuid = lpuid
  IF !SEEK(m.lpuid, 'pilot')
   IF !SEEK(m.lpuid, 'horlpu')
    LOOP 
   ENDIF 
*   LOOP 
  ENDIF 
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  m.llpath = m.lpath+'\'+m.mcod
  IF !fso.FileExists(m.llpath+'\people.dbf') OR !fso.FileExists(m.llpath+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.llpath+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llpath+'\talon', 'talon', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.llpath+'\e'+m.mcod, 'errs', 'shar', 'rid')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('errs')
    USE IN errs
   ENDIF 
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  CREATE CURSOR cpazall (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazempty (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazown (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazguests (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpaznopilot (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR cpazbad (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR curo (mcod c(7), sall n(11,2), npaz n(6))
  INDEX on mcod TAG mcod
  SET ORDER TO mcod 
  CREATE CURSOR curp (mcod c(7), sn_pol c(25))
  INDEX ON mcod TAG mcod
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol

  SELECT talon 
  SET RELATION TO recid INTO errs
  SET RELATION TO sn_pol INTO people ADDITIVE 
  
  m.s_all    = 0
  m.s_empty  = 0
  m.s_own    = 0
  m.s_guests = 0
  m.s_npilot = 0
  m.s_bad    = 0
  m.s_kompl  = 0
  m.s_dst    = 0
  
  m.paz_all    = 0
  m.paz_empty  = 0
  m.paz_own    = 0
  m.paz_guests = 0
  m.paz_npilot = 0
  m.paz_bad    = 0
  
  m.s_pred_pf  = 0
  
  SCAN 
   IF !EMPTY(errs.rid)
    LOOP 
   ENDIF 
   m.cod    = cod
   m.otd    = SUBSTR(otd,2,2)
   m.d_type = d_type
*   m.IsTpnR = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod) OR IsEko(m.cod)), .T., .F.)
   m.IsTpnR = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
   m.ord    = ord
   m.lpu_ord = lpu_ord
   
*   IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) OR IsEKO(m.cod) && стационар
   IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) && стационар
    LOOP 
   ENDIF 
   IF m.IsTpnR OR INLIST(m.otd,'70','73','93') OR m.d_type='s' && допуслуги
    LOOP 
   ENDIF 
   IF (INLIST(m.otd,'01','90') AND IsStac(m.mcod)) AND people.mcod!=people.prmcod && допуслуги
    LOOP 
   ENDIF 
   IF (m.ord=7 AND m.lpu_ord=7665) AND people.mcod!=people.prmcod && допуслуги
    LOOP 
   ENDIF 

   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)

   m.sn_pol  = sn_pol
   m.s_all   = m.s_all + s_all
   m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
   m.otd     = SUBSTR(otd,2,2)

   m.fil_id = fil_id

   IF !SEEK(m.sn_pol, 'cpazall')
    INSERT INTO cpazall (sn_pol) VALUES (m.sn_pol)
   ENDIF 

   m.mcod2  = people.prmcod
   m.paztip = TipOfPaz(m.mcod, m.mcod2) && 0 (не прикреплен),1 (прикреплен по месту обращения),2 (к пилоту),3 (не к пилоту)
   
   IF m.paztip=0
*    IF m.IsLpuTpn = .T.
*     IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*      LOOP 
*     ENDIF 
*    ENDIF 
    m.s_empty  = m.s_empty + s_all
    IF !SEEK(m.sn_pol, 'cpazempty')
     INSERT INTO cpazempty (sn_pol) VALUES (m.sn_pol)
    ENDIF 
    LOOP 
   ENDIF 
   
   DO CASE 
    CASE m.paztip = 1 && свой у себя
*     IF m.IsLpuTpn = .T.
*      IF !SEEK(m.fil_id, 'lputpn', 'fil_id') AND INLIST(m.otd,'01','70','73')
*       LOOP 
*      ENDIF 
*     ENDIF 
     m.s_own = m.s_own + s_all
     m.s_pred_pf = m.s_pred_pf
     IF !SEEK(m.sn_pol, 'cpazown')
      INSERT INTO cpazown (sn_pol) VALUES (m.sn_pol)
     ENDIF 

    CASE m.paztip = 2 && чужой пилот
*     IF m.IsLpuTpn = .T.
*      IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*       LOOP 
*      ENDIF 
*     ENDIF 
    
     IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(m.otd,'08','92'))) && если есть направление или скорая помощь

      IF !SEEK(m.sn_pol, 'cpazguests')
       INSERT INTO cpazguests (sn_pol) VALUES (m.sn_pol)
      ENDIF 

      IF !SEEK(m.mcod2, 'curo')
       INSERT INTO curo (mcod) VALUES (m.mcod2)
      ENDIF 
      IF !SEEK(m.sn_pol, 'curp')
       INSERT INTO curp (mcod, sn_pol) VALUES (m.mcod2, m.sn_pol)
      ENDIF 

      m.s_guests = m.s_guests + s_all

      =SEEK(m.mcod2, 'curo')
      m.osall = curo.sall
      m.nsall = m.osall + s_all
      UPDATE curo SET sall = m.nsall WHERE mcod = m.mcod2
      
*      REPLACE vz WITH .t. 
     
     ELSE && если без направления и не скорая помощь
     
      m.s_bad = m.s_bad + s_all
*      m.s_pred_pf = m.s_pred_pf + IIF(IsUsl(m.cod) OR IsKDP(m.cod), s_all, 0)
      IF !SEEK(m.sn_pol, 'cpazbad')
       INSERT INTO cpazbad (sn_pol) VALUES (m.sn_pol)
      ENDIF 

     ENDIF 

    CASE m.paztip = 3 && чужой непилот
     IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(m.otd,'08','92'))) && если есть направление или скорая помощь

     IF !SEEK(m.sn_pol, 'cpaznopilot')
      INSERT INTO cpaznopilot (sn_pol) VALUES (m.sn_pol)
     ENDIF 
     m.s_npilot  = m.s_npilot  + s_all
     
     ELSE  && если без направления и не скорая помощь
     
      m.s_bad = m.s_bad + s_all
*      m.s_pred_pf = m.s_pred_pf + IIF(IsUsl(m.cod) OR IsKDP(m.cod), s_all, 0)
      IF !SEEK(m.sn_pol, 'cpazbad')
       INSERT INTO cpazbad (sn_pol) VALUES (m.sn_pol)
      ENDIF 

     ENDIF 

   ENDCASE 

  ENDSCAN 
  
  m.paz_all     = RECCOUNT('cpazall')
  m.paz_empty   = RECCOUNT('cpazempty')
  m.paz_own     = RECCOUNT('cpazown')
  m.paz_guests  = RECCOUNT('cpazguests')
  m.paz_nopilot = RECCOUNT('cpaznopilot')
  m.paz_bad     = RECCOUNT('cpazbad')

  UPDATE curpr4 SET paz_all=m.paz_all, s_all=m.s_all, paz_empty=m.paz_empty,;
   s_empty=m.s_empty, paz_own=m.paz_own, s_own=m.s_own, s_pred_pf=m.s_pred_pf,;
   paz_guests=m.paz_guests, s_guests=m.s_guests, paz_npilot=m.paz_nopilot,;
   paz_bad = m.paz_bad, s_bad=m.s_bad,;
   s_npilot=m.s_npilot, s_kompl=m.s_kompl, s_dst=m.s_dst WHERE mcod = m.mcod
  
  SET RELATION OFF INTO people
  SET RELATION OFF INTO errs
  
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('errs')
   USE IN errs
  ENDIF 

  IF USED('curo')  
   SELECT curo
   SCAN 
    m.mcodd = mcod 
    IF !SEEK(m.mcodd, 'pilot', 'mcod')
     IF !SEEK(m.mcodd, 'horlpu', 'mcod')
      LOOP 
     ENDIF 
*     LOOP 
    ENDIF 
    m.sall = sall
    =SEEK(m.mcodd, 'curpr4')
    m.os_others = curpr4.s_others
    UPDATE curpr4 SET s_others = m.os_others + m.sall WHERE mcod = m.mcodd
   ENDSCAN 
   USE IN curo
  ENDIF

  IF USED('curp')
   SELECT curp
   SCAN 
    m.mcodd = mcod
    =SEEK(m.mcodd, 'curpr4')
    m.opaz_others = curpr4.paz_others
    UPDATE curpr4 SET paz_others = m.opaz_others + 1 WHERE mcod = m.mcodd
   ENDSCAN 
   USE IN curp
  ENDIF 
  
  IF USED('cpazall')
   USE IN cpazall
  ENDIF  
  IF USED('cpazempty')
   USE IN cpazempty
  ENDIF  
  IF USED('cpazown')
   USE IN cpazown
  ENDIF  
  IF USED('cpazguests')
   USE IN cpazguests
  ENDIF  
  IF USED('cpaznopilot')
   USE IN cpaznopilot
  ENDIF  
  IF USED('cpazbad')
   USE IN cpazbad
  ENDIF  

  SELECT aisoms
  
 ENDSCAN 

 WAIT CLEAR 

 USE IN aisoms
 USE IN pilot
* USE IN rcodes
 USE IN lputpn 
 USE IN tarif
 USE IN horlpu
 
 m.IsAttPplOk = .T.
 m.attbase = ''
 IF EMPTY(m.pattppl)
  m.IsAttPplOk = .F.
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FolderExists(m.pattppl)
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FileExists(m.pattppl+'\attppl.cfg')
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF OpenFile(m.pattppl+'\attppl.cfg', 'attcfg', 'shar')>0
   IF USED('attcfg')
    USE IN attcfg
   ENDIF 
   m.IsAttPplOk = .F.
  ELSE 
   SELECT attcfg
   m.attbase = ALLTRIM(pbase)
   USE IN attcfg
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FolderExists(m.attbase)
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FolderExists(m.attbase+'\'+m.lcperiod)
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF !fso.FileExists(m.attbase+'\'+m.lcperiod+'\aisoms.dbf')
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 IF m.IsAttPplOk = .T.
  IF OpenFile(m.attbase+'\'+m.lcperiod+'\aisoms', 'attais', 'shar', 'lpuid')>0
   IF USED('attais')
    USE IN attais
   ENDIF 
   m.IsAttPplOk = .F.
  ENDIF 
 ENDIF 
 
 SELECT curpr4
 REPLACE ALL adnorm WITH m.adnorm, chnorm WITH m.chnorm
 IF m.IsAttPplOk = .T.
  SET RELATION TO lpuid INTO attais
  REPLACE ALL adults WITH attais.ad_mgf, childs WITH attais.ch_mgf
  REPLACE ALL adsum WITH adnorm*adults, chsum WITH chnorm*childs

  SCAN 

   m.pazval = attais.ch01m+attais.ch01f+attais.ch14m+attais.ch14f+;
    attais.ch514m+attais.ch514f+attais.ch1517m+attais.ch1517f+;
    attais.m1824+attais.f1824+attais.m2534+attais.f2534+;
    attais.m3544+attais.f3544+attais.m4559+attais.f4559+;
    attais.m6068+attais.f5564+attais.m69+attais.f65

   m.finval = attais.ch01m*m.m0001+attais.ch01f*m.f0001+attais.ch14m*m.m0104+attais.ch14f*m.f0104+;
    attais.ch514m*m.m0514+attais.ch514f*m.f0514+attais.ch1517m*m.m1517+attais.ch1517f*m.f1517+;
    attais.m1824*m.m1824+attais.f1824*m.f1824+attais.m2534*m.m2534+attais.f2534*m.f2534+;
    attais.m3544*m.m3544+attais.f3544*m.f3544+attais.m4559*m.m4559+attais.f4559*m.f4554+;
    attais.m6068*m.m6068+attais.f5564*m.f5564+attais.m69*m.m6999+attais.f65*m.f6599

   REPLACE finval WITH m.finval, pazval WITH m.pazval

  ENDSCAN 
  SET RELATION OFF INTO attais
  USE IN attais
 ENDIF 
 COPY TO m.lPath+'\pr4'
 USE 
 
 IF OpenFile(m.lPath+'\pr4', 'pr4', 'excl')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
 ELSE 
  SELECT pr4
  INDEX on lpuid TAG lpuid
  INDEX on mcod TAG mcod
  USE 
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ОБРАБОТКА ЗАКОНЧЕНА!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 

*FUNCTION TipOfPaz(amcod,bmcod)
* PRIVATE lcmcod, lcprmcod, IsPilot, m.paztip
* m.lcmcod   = amcod
* m.lcprmcod = bmcod
* m.IsPilot = IIF(SEEK(m.lcprmcod, 'pilot', 'mcod'), .t., .f.)

* m.paztip = 0

* DO CASE 
*  CASE EMPTY(m.lcprmcod) && не прикреплен
*   m.paztip = 0
*  CASE m.lcmcod = m.lcprmcod && прикреплен по месту обращения
*   m.paztip = 1
*  CASE m.lcmcod != m.lcprmcod AND m.IsPilot=.t. && прикреплен к пилоту не по месту обращения
*   m.paztip = 2
*  CASE m.lcmcod != m.lcprmcod AND m.IsPilot=.f. && прикреплен к не пилоту не по месту обращения
*   m.paztip = 3
*  OTHERWISE 
*  m.paztip = 0
* ENDCASE 

*RETURN m.paztip

