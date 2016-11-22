FUNCTION OneFlk(ppath)

 IF IsPr
  RETURN 
 ENDIF 

 LocalErrIniFile    = ppath + '\errors.ini'
 IsLocalErrIniFile  = fso.FileExists(LocalErrIniFile)
 GlobalErrIniFile   = pbin + '\errors.ini'
 IsGlobalErrIniFile = fso.FileExists(GlobalErrIniFile)
 
 WorkIniFile = ''
 IF IsLocalErrIniFile
  WorkIniFile = LocalErrIniFile
 ELSE 
  IF IsGlobalErrIniFile
   WorkIniFile = GlobalErrIniFile
  ENDIF 
 ENDIF 
 
 IF EMPTY(WorkIniFile)
  M.PSA = .T.
  M.ERA = .T.
  M.ECA = .T.
  M.E1A = .T.
  M.E2A = .T.
  M.E4A = .T.
  M.E7A = .T.
  M.E8A = .T.

  M.H6A = .T.
  M.COA = .T.
  M.HCA = .T.
  M.DUA = .T.
  M.H8A = .T.
  M.HEA = .T.
  M.CSA = .T.
  M.TVA = .T.
  M.NLA = .T.
  M.MDA = .T.
  M.H3A = .T.
  M.SOA = .T.
  M.R1A = .T.
  M.R2A = .T.
  M.R3A = .T. 
  M.UVA = .T.
  M.DVA = .T.
  M.UOA = .T.
  M.NOA = .T.
  M.NMA = .T.
  M.NUA = .T.
  M.NSA = .T.
  M.SMA = .T.
  M.DIA = .T.
  M.DDA = .T.
  M.HNA = .T.
  M.DLA = .T.
  M.DRA = .T.
  M.POA = .T.
  M.VDA = .T.
  M.TFA = .T.
  M.PPA = .F.
  M.G1A = .F.
  M.G2A = .F.
  M.G3A = .F.
  M.G4A = .F.
  M.NRA = .F.
  M.KEA = .F.
  M.D2A = .F.
  M.THA = .T.
  M.TLA = .T.
  M.HOA = .T.
 ELSE 
  M.PSA = ReadFromIni(WorkIniFile, 'PSA', .T.)
  M.ERA = ReadFromIni(WorkIniFile, 'ERA', .T.)
  M.ECA = ReadFromIni(WorkIniFile, 'ECA', .T.)
  M.E1A = ReadFromIni(WorkIniFile, 'E1A', .T.)
  M.E2A = ReadFromIni(WorkIniFile, 'E2A', .T.)
  M.E4A = ReadFromIni(WorkIniFile, 'E4A', .T.)
  M.E7A = ReadFromIni(WorkIniFile, 'E7A', .T.)
  M.E8A = ReadFromIni(WorkIniFile, 'E8A', .T.)

  M.H6A = ReadFromIni(WorkIniFile, 'H6A', .T.)
  M.COA = ReadFromIni(WorkIniFile, 'COA', .T.)
  M.HCA = ReadFromIni(WorkIniFile, 'HCA', .T.)
  M.DUA = ReadFromIni(WorkIniFile, 'DUA', .T.)
  M.H8A = ReadFromIni(WorkIniFile, 'H8A', .T.)
  M.HEA = ReadFromIni(WorkIniFile, 'HEA', .T.)
  M.CSA = ReadFromIni(WorkIniFile, 'CSA', .T.)
  M.TVA = ReadFromIni(WorkIniFile, 'TVA', .T.)
  M.NLA = ReadFromIni(WorkIniFile, 'NVA', .T.)
  M.MDA = ReadFromIni(WorkIniFile, 'MDA', .T.)
  M.H3A = ReadFromIni(WorkIniFile, 'H3A', .T.)
  M.SOA = ReadFromIni(WorkIniFile, 'SOA', .T.)
  M.R1A = ReadFromIni(WorkIniFile, 'R1A', .T.)
  M.R2A = ReadFromIni(WorkIniFile, 'R2A', .T.)
  M.R3A = ReadFromIni(WorkIniFile, 'R3A', .T.)
  M.UVA = ReadFromIni(WorkIniFile, 'UVA', .T.)
  M.DVA = ReadFromIni(WorkIniFile, 'DVA', .T.)
  M.UOA = ReadFromIni(WorkIniFile, 'UOA', .T.)
  M.NOA = ReadFromIni(WorkIniFile, 'NOA', .T.)
  M.NMA = ReadFromIni(WorkIniFile, 'NMA', .T.)
  M.NUA = ReadFromIni(WorkIniFile, 'NUA', .T.)
  M.NSA = ReadFromIni(WorkIniFile, 'NSA', .T.)
  M.SMA = ReadFromIni(WorkIniFile, 'SMA', .T.)
  M.DIA = ReadFromIni(WorkIniFile, 'DIA', .T.)
  M.DDA = ReadFromIni(WorkIniFile, 'DDA', .T.)
  M.HNA = ReadFromIni(WorkIniFile, 'HNA', .T.)
  M.DLA = ReadFromIni(WorkIniFile, 'DLA', .T.)
  M.DRA = ReadFromIni(WorkIniFile, 'DRA', .T.)
  M.POA = ReadFromIni(WorkIniFile, 'POA', .T.)
  M.VDA = ReadFromIni(WorkIniFile, 'VDA', .T.)
  M.TFA = ReadFromIni(WorkIniFile, 'TFA', .T.)
  M.PPA = ReadFromIni(WorkIniFile, 'PPA', .F.)
  M.G1A = ReadFromIni(WorkIniFile, 'G1A', .F.)
  M.G2A = ReadFromIni(WorkIniFile, 'G2A', .F.)
  M.G3A = ReadFromIni(WorkIniFile, 'G3A', .F.)
  M.G4A = ReadFromIni(WorkIniFile, 'G4A', .F.)
  M.NRA = ReadFromIni(WorkIniFile, 'NRA', .F.)
  M.KEA = ReadFromIni(WorkIniFile, 'KEA', .F.)
  M.D2A = ReadFromIni(WorkIniFile, 'D2A', .F.)
  M.THA = ReadFromIni(WorkIniFile, 'THA', .T.)
  M.TLA = ReadFromIni(WorkIniFile, 'TLA', .T.)
  M.HOA = ReadFromIni(WorkIniFile, 'HOA', .T.)
 ENDIF 
 
  m.lpuid   = lpuid
  m.mcod    = mcod
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')+', '+sprlpu.cokr+', '+sprlpu.mcod
  m.period  = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
  m.dat1 = CTOD('01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4))
  m.dat2 = GOMONTH(m.dat1,1)-1
  m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
  m.IsStac = IIF(VAL(SUBSTR(m.mcod,3,2))<41 or m.IsLpuTpn, .F., .T.)
  m.IsPilot = IIF(SEEK(m.lpuid, 'pilot'), .t., .f.)
  m.IsHor = IIF(SEEK(m.lpuid, 'horlpu'), .t., .f.)
  
  M.G1A = IIF(m.IsHor=.t., M.G1A, .F.)
  M.G2A = IIF(m.IsHor=.t., M.G2A, .F.)
  M.G3A = IIF(m.IsHor=.t., M.G3A, .F.)
  M.G4A = IIF(m.IsHor=.t., M.G4A, .F.)
  M.NRA = IIF(m.IsHor=.t., M.NRA, .F.)
  
  m.lIsDspExists = .f.
  m.dspfile1 = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
  m.dspfile2 = pbase +'\'+ STR(tyear-2,4)+'12'+'\dsp'
  m.dspfile3 = pbase +'\'+ STR(tyear-3,4)+'12'+'\dsp'
  IF tmonth>1
   m.dspfile = pbase +'\'+ STR(tyear,4)+PADL(tmonth-1,2,'0')+'\dsp'
  ELSE
   m.dspfile = pbase +'\'+ STR(tyear-1,4)+'12'+'\dsp'
  ENDIF 

*  m.dspfile = m.dspfile1

  IF fso.FileExists(m.dspfile+'.dbf')
   m.lIsDspExists = .t.
  ELSE 
   m.lIsDspExists = .f.
  ENDIF 

  oal = ALIAS()
  IF OpenFile(m.dspfile, 'cdsp', 'shar')>0 && mcod+sn_pol+padl(cod,6,'0')
   IF USED('cdsp')
    USE IN cdsp
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
   SELECT * FROM cdsp INTO CURSOR dspp READWRITE 
   USE IN cdsp
   SELECT dspp
   INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag
   SET ORDER TO exptag
   IF fso.FileExists(m.dspfile1+'.dbf') AND tmonth>1
    APPEND FROM &dspfile1
   ENDIF 
   IF fso.FileExists(m.dspfile2+'.dbf')
    APPEND FROM &dspfile2
   ENDIF 
   IF fso.FileExists(m.dspfile3+'.dbf')
    APPEND FROM &dspfile3
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .t.
  ENDIF  

  IF 1=2
   m.nindkol = ATAGINFO(atttag)
   IF m.nindkol<=1
    USE IN dsp 
    IF OpenFile(m.dspfile, 'dsp', 'excl')>0
     IF USED('dsp')
      USE IN dsp
     ENDIF 
     SELECT (oal)
     m.lIsDspExists = .f.
    ELSE 
     SELECT dsp 
     INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag
     SET ORDER TO exptag
     SELECT (oal)
     m.lIsDspExists = .t.
    ENDIF 
   ELSE 
    SET ORDER TO exptag IN dsp 
    SELECT (oal)
    m.lIsDspExists = .t.
   ENDIF 
  ENDIF 
  
  IF OpenFile(m.dspfile, 'dspyear', 'shar')>0
   IF USED('dspyear')
    USE IN dspyear
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
  SELECT dspyear
  =ATAGINFO(taginf)
  IF ASCAN(taginf,'EXPTAG')>0
   SET ORDER TO exptag
   m.lIsDspExists = .t.
  ELSE 
  IF OpenFile(m.dspfile, 'dspyear', 'excl')>0 && mcod+sn_pol+padl(cod,6,'0')
   IF USED('dspyear')
    USE IN dspyear
   ENDIF 
   SELECT (oal)
   m.lIsDspExists = .f.
  ELSE 
   SELECT dspyear
   INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag
   SET ORDER TO exptag
   USE IN dspyear
   =OpenFile(m.dspfile, 'dspyear', 'shar', 'exptag')
   m.lIsDspExists = .t.
  ENDIF 
  ENDIF 
  ENDIF 

  M.D2A = IIF(m.lIsDspExists = .t., M.D2A, .f.)
  
*  MESSAGEBOX(IIF(m.d2a=.t., '.T.','.F.'), 0+64,'')

  lcError = ppath+'\e'+m.mcod
  IF !fso.FileExists(lcError+'.dbf')
   CREATE TABLE (lcError) (f c(1), c_err c(3), rid i)
   INDEX FOR UPPER(f)='R' ON rid TAG rrid
   INDEX FOR UPPER(f)='S' ON rid TAG rid
   USE 
  ENDIF 
  
  IF !OpBase(ppath)
   RETURN .f.
  ENDIF 

  SELECT sn_pol, cod, SUM(k_u) AS k_u, d_u, SUM(s_all) AS s_all ;
   FROM Talon GROUP BY sn_pol, d_u, cod WHERE d_type != '2' ;
   INTO CURSOR day_gr

*  SELECT sn_pol AS sn_pol, a.cod AS cod, ;
   k_u AS k_u, 000 as cntr, d_u AS d_u, s_all AS s_all, IIF(!IsStac, mdayp, mdays) AS in_day,;
   IIF(!IsStac, mdayp, mdays) as k_norm;
   FROM day_gr a, codku b ;
   WHERE a.cod=b.cod AND k_u > IIF(!IsStac, mdayp, mdays) ;
   INTO CURSOR e_day ORDER BY a.sn_pol, a.d_u, a.cod READWRITE 
  SELECT sn_pol AS sn_pol, a.cod AS cod, ;
   k_u AS k_u, 000 as cntr, d_u AS d_u, s_all AS s_all, mdayp AS in_day,;
   mdayp as k_norm;
   FROM day_gr a, codku b ;
   WHERE a.cod=b.cod AND k_u > mdayp ;
   INTO CURSOR e_day ORDER BY a.sn_pol, a.d_u, a.cod READWRITE 
  SELECT e_day
  INDEX ON sn_pol + STR(cod,6) + DTOS(d_u) TAG ExpTag
  SET ORDER TO ExpTag

  SELECT sn_pol, cod, SUM(k_u) AS k_u, MIN(d_u) AS d_u, SUM(s_all) AS s_all ;
   FROM Talon  WHERE d_type!='2' GROUP BY sn_pol, cod ;
   INTO CURSOR month_gr

*  SELECT sn_pol as sn_pol, a.cod as cod, k_u as k_u, 000 as cntr, IIF(!IsStac, mmsp, mmss) as k_norm, s_all as s_all, ;
   IIF(!IsStac, b.mmsp, b.mmss) as in_month ;
   FROM month_gr a, codku b ;
   WHERE a.cod=b.cod and k_u > IIF(!IsStac, mmsp, mmss) ;
   INTO CURSOR e_month ORDER BY sn_pol, a.cod READWRITE 
  SELECT sn_pol as sn_pol, a.cod as cod, k_u as k_u, 000 as cntr, mmsp as k_norm, s_all as s_all, ;
   b.mmsp as in_month ;
   FROM month_gr a, codku b ;
   WHERE a.cod=b.cod and k_u > mmsp ;
   INTO CURSOR e_month ORDER BY sn_pol, a.cod READWRITE 
  SELECT e_month
  INDEX ON sn_pol + STR(cod,6) TAG ExpTag
  SET ORDER TO ExpTag

  SELECT * FROM Talon ORDER BY c_i, d_u DESC, k_u DESC INTO CURSOR Gosp
  
  m.s_flk = 0  

  IF M.DRA == .T.
   SELECT recid, sn_pol FROM people WHERE sn_pol IN ;
   (SELECT sn_pol FROM people GROUP BY sn_pol HAVING coun(*)>1) INTO CURSOR dblppl
   IF _tally>0
    SELECT dblppl
    INDEX on sn_pol TAG sn_pol UNIQUE 
    SET ORDER TO sn_pol
    SELECT talon
    SET RELATION TO sn_pol INTO dblppl
    SCAN 
     IF !EMPTY(dblppl.sn_pol)  
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval = InsError('S', 'PKA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
     ENDIF 
    ENDSCAN 
    SET RELATION OFF INTO dblppl
    SELECT dblppl
    SET ORDER TO 
    SCAN 
     m.recid = recid
     =InsError('R', 'DRA', m.recid)
*      InsErrorSV(m.mcod, 'R', 'DRA', m.recid)
    ENDSCAN 
   ENDIF 
  USE IN dblppl
  ENDIF 
  
  SELECT talon
  SET RELATION TO sn_pol INTO people 
  
  SCAN

   IF EMPTY(people.sn_pol)               && �������� PS
    m.polis = sn_pol
    DO WHILE sn_pol == m.polis
     m.recid = recid
     rval = InsError('S', 'PSA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'PSA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     SKIP +1 
    ENDDO 
   ENDIF 

  IF people.IsPr==.F. && ���������� ���������� ������ ��������!

   IF M.ERA == .T. && �������� ER
    IF !EMPTY(people.sv)  
     m.IsGood = IIF(SEEK(people.sv, 'osoerz') AND osoerz.kl == 'y', .T., .F.)
     IF IsVS(people.sn_pol) AND LEFT(people.sn_pol,2)=m.qcod
      IF USED('kms')
       m.vvs = INT(VAL(SUBSTR(ALLTRIM(people.sn_pol),7)))
       IF SEEK(m.vvs, 'kms')
        m.IsGood = .t.
       ENDIF 
      ENDIF 
     ENDIF 
     IF IsGood == .f.
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval = InsError('S', 'PKA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'ERA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.ECA == .T. && �������� EC
    IF !EMPTY(people.sv)
     m.IsGood = IIF(people.qq = m.qcod, .T., .F.)
     IF IsVS(people.sn_pol) AND LEFT(people.sn_pol,2)=m.qcod
      IF USED('kms')
       m.vvs = INT(VAL(SUBSTR(ALLTRIM(people.sn_pol),7)))
       IF SEEK(m.vvs, 'kms')
        m.IsGood = .t.
       ENDIF 
      ENDIF 
     ENDIF 
     IF IsGood == .f.                 
      m.polis = sn_pol
      m.recsproc = 0 
      DO WHILE sn_pol == m.polis
       m.recid = recid
       rval =InsError('S', 'PKA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       SKIP +1 
       m.recsproc = m.recsproc + 1
      ENDDO 
      SKIP -1*(m.recsproc)
      m.recid = people.recid
      =InsError('R', 'ECA', m.recid)
*      InsErrorSV(m.mcod, 'R', 'ECA', m.recid)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.E1A == .T.  && �������� E1
    IF !SEEK(people.d_type, 'osoree')
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E1A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E1A', m.recid)
    ENDIF 
   ENDIF 
   
   IF M.E2A == .T. && �������� E2
    IF (!IsKms(people.sn_pol) AND !IsVS(people.sn_pol) AND !IsENP(people.sn_pol))
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E2A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E2A', m.recid)
    ENDIF 
   ENDIF 

   IF  M.E4A == .T. && �������� E4
    IF ((INLIST(RIGHT(PADL(ALLTRIM(People.fam),25),2),'��','��','��') AND INLIST(RIGHT(PADL(ALLTRIM(People.ot),20),2),'��','��') AND People.w!=2) OR ;
       (INLIST(RIGHT(PADL(ALLTRIM(People.fam),25),2),'��','��','��')  AND INLIST(RIGHT(PADL(ALLTRIM(People.ot),20),2),'��','��') AND People.w!=1))
     m.polis = sn_pol 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E4A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E4A', m.recid)
    ENDIF 
   ENDIF 
   
   IF M.E7A == .T. && �������� E7
    IF (!INLIST(people.w,1,2) OR (IsKms(people.sn_pol) AND SUBSTR(people.sn_pol,5,2)!='77' AND (people.w != IIF(VAL(SUBSTR(people.sn_pol,12,2))>50, 1, 2))))
     m.polis = sn_pol
     m.recsproc = 0 
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E7A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E7A', m.recid)
    ENDIF 
   ENDIF 

   IF M.E7A == .T.
    m.sn_pol = people.sn_pol                && �������� E7
    Dtt = CTOD(IIF(VAL(SUBSTR(m.sn_pol,12,2))>50, ;
         PADL(INT(VAL(SUBSTR(m.sn_pol,12,2))-50),2,'0'), ;
         SUBSTR(m.sn_pol,12,2))+'.'+IIF(VAL(SUBSTR(m.sn_pol,14,2))>40, ;
         PADL(INT(VAL(SUBSTR(m.sn_pol,14,2))-40),2,'0')+'.20', ;
         SUBSTR(m.sn_pol,14,2)+'.19')+SUBSTR(m.sn_pol,16,2))
    IF (IsKms(m.sn_pol) AND !INLIST(SUBSTR(m.sn_pol,5,2),'50','51') AND (people.dr != Dtt))
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E7A', m.recid)
*     InsErrorSV(m.mcod, 'R', 'E7A', m.recid)
    ENDIF 
   ENDIF 

   IF M.E8A == .T.
    m.sn_pol = people.sn_pol                && �������� E8
    IF (people.dr=={} OR (dat1-IIF(!EMPTY(people.dr), people.dr, {01.01.1850}))/365.25>120 OR ;
     IIF(!EMPTY(people.dr), people.dr, {01.01.1850}) > m.dat2)
     m.polis = sn_pol
     m.recsproc = 0 
     DO WHILE sn_pol == m.polis
      m.recid = recid
      rval =InsError('S', 'PKA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'PKA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      SKIP +1 
      m.recsproc = m.recsproc + 1
     ENDDO 
     SKIP -1*(m.recsproc)
     m.recid = people.recid
     =InsError('R', 'E8A', m.recid)
*     InsErrorSV(m.mcod,'R', 'E8A', m.recid)
    ENDIF 
   ENDIF 
   
  ENDIF && ���������� ���������� ������ ��������!

  IF talon.IsPr == .F. && ���������� ���������� ������ �����!
   
   && ����� ������� ��������� �������� �����!

   IF M.H6A == .T. && �������� H6
    m.polis=''
    DO CASE 
     CASE IsKms(sn_pol)
      m.polis = SUBSTR(sn_pol,8)
     CASE IsVs(sn_pol)
      m.polis = SUBSTR(sn_pol,7)
     OTHERWISE 
      m.polis = sn_pol
    ENDCASE 
    IF EMPTY(c_i)
     m.recid = recid
     rval =InsError('S', 'H6A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'H6A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945))
     IF SUBSTR(c_i,1,6)!='�����_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF INLIST(cod,101946,101947,101948)
     IF SUBSTR(c_i,1,6)!='�����_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF INLIST(cod,101949,101950)
     IF SUBSTR(c_i,1,4)!='���_' OR ALLTRIM(SUBSTR(c_i,5))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF (BETWEEN(cod,1900,1905) OR BETWEEN(cod,101929,101932))
     IF !INLIST(SUBSTR(c_i,1,3),'��_','��_') OR ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,1906,1909)
     IF SUBSTR(c_i,1,6)!='�����_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+ALLTRIM(sn_pol)+CHR(13)+CHR(10)+ALLTRIM(c_i),0+64,'H6A')
     ENDIF 
    ENDIF 
    
    IF people.d_type='9' AND OCCURS('#', ALLTRIM(c_i))>0
     m.c_i = ALLTRIM(c_i)
     IF OCCURS('#', m.c_i)!=3
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      MESSAGEBOX('H6A',0+64,m.c_i)
     ELSE 
      DO CASE 
       CASE !INLIST(SUBSTR(m.c_i,AT('#',m.c_i)+1,1),'1','2')
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*        MESSAGEBOX('H6A',0+64,m.c_i)
       CASE EMPTY(CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4)))
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*        MESSAGEBOX('H6A',0+64,m.c_i)
       CASE !INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6')
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*        MESSAGEBOX('H6A',0+64,m.c_i)
      ENDCASE 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.COA == .T. && �������� CO 
*    IF EMPTY(otd) OR LEN(ALLTRIM(otd))!=4 OR ;
     (!ISDIGIT(SUBSTR(otd,1,1)) OR !ISDIGIT(SUBSTR(otd,2,1)) OR !ISDIGIT(SUBSTR(otd,3,1)))
    IF EMPTY(otd) OR ;
     (!ISDIGIT(SUBSTR(otd,1,1)) OR !ISDIGIT(SUBSTR(otd,2,1)) OR !ISDIGIT(SUBSTR(otd,3,1)))
     m.recid = recid
     rval =InsError('S', 'COA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'COA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.HCA == .T. && �������� HC
    IF k_u <= 0
     m.recid = recid
     rval = InsError('S', 'HCA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HCA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.DUA == .T. && ���� ��������� DU
    IF (people.tip_p==1 AND MONTH(d_u)!=tMonth)
     m.recid = recid
     rval = InsError('S', 'DUA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'DUA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.H8A == .T. && ���� ��������� H8
    IF !SEEK(ds, 'mkb10')
     m.recid = recid
     rval = InsError('S', 'H8A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'H8A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE 
     IF INLIST(LEFT(ds,3),'B95','B96','B97')
      m.recid = recid
      rval = InsError('S', 'H8A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'H8A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     m.cod = cod
     m.ds = ds
     IF (LEFT(m.ds,1)='Z' AND !INLIST(m.ds,'Z13.8','Z01.7','Z20','Z34','Z35')) AND INLIST(FLOOR(m.cod/1000),25,26,27,28,29,30,125,126,127,128,129,130)
      m.recid = recid
      rval = InsError('S', 'H8A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'H8A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF
   ENDIF 

   IF M.HEA == .T.  && �������� HE
    m.sex = IIF(OCCURS('#',c_i)==2, SUBSTR(c_i, AT('#',c_i,1)+1, 1), STR(people.w,1))
    IF (SEEK(ds, 'mkb10') AND !EMPTY(mkb10.sex)) AND m.sex != mkb10.sex
     m.recid = recid
     rval = InsError('S', 'HEA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HEA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 
   
   IF M.CSA == .T.
    IF !SEEK(cod, 'tarif')
     m.recid = recid
     rval = InsError('S', 'CSA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'CSA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.TFA == .T.
    IF SEEK(cod, 'tarif')
*     IF !EMPTY(Tarif.n_kd) AND EMPTY(Tip)
     IF !EMPTY(Tarif.n_kd) AND !SEEK(Tip, 'kpresl')
      m.recid = recid
      rval = InsError('S', 'TFA', m.recid)
 *     InsErrorSV(m.mcod, 'S', 'TFA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   IF M.TLA == .T.
    m.cod = cod
    m.k_u = k_u
    IF (INLIST(m.cod,97041,97013,197013) OR BETWEEN(cod, 84000, 84999)) AND m.k_u>1
     m.recid = recid
     rval = InsError('S', 'TLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF M.THA == .T.
    m.cod = cod
    m.otd = otd
    m.c_i    = c_i
    m.sn_pol = sn_pol
    m.tip = tip
    IF INLIST(m.tip,'0','8')
    IF IsHOOtd(m.otd) AND !SEEK(m.cod, 'noth')
     m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
     IF !USED('ho')
      m.recid = recid
      rval = InsError('S', 'THA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      IF !SEEK(m.vir, 'ho')
       m.recid = recid
       rval = InsError('S', 'THA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
    ENDIF 
   ENDIF 

   IF M.HOA == .T.
   ENDIF 

   IF M.TVA == .T.
    m.tiplpu = SUBSTR(m.mcod,2,1)
    IF m.tiplpu!='3'
     IF SEEK(cod, 'codwdr') AND (!EMPTY(codwdr.kp) AND m.tiplpu!=codwdr.kp)
      m.recid = recid
      rval = InsError('S', 'TVA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'TVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   IF M.NLA == .T.
    m.tiplpu = IIF(VAL(SUBSTR(m.mcod,3,2))<41, 'p', 's')
    IF SEEK(cod, 'codwdr') AND !EMPTY(codwdr.stac) AND LOWER(codwdr.stac) != m.tiplpu
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'NLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF

    * ��������������� �����-�����
    IF INLIST(cod,101929,101930,101931,101932)
    IF !SEEK(m.lpuid, 'dsdisp')
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
    ENDIF 
    * ��������������� �����-�����
    
    IF IsVMP(cod) AND !SEEK(m.lpuid, 'movmp')
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    * ��������������� ��������
    IF INLIST(cod,1900,1901,1902,1903,1904,1905)
     IF !SEEK(m.lpuid, 'spidd')
      m.recid = recid
      rval = InsError('S', 'NLA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
    * ��������������� ��������

   ENDIF 

   IF M.MDA == .T.
    m.sex = IIF(OCCURS('#',c_i)==2, SUBSTR(c_i, AT('#',c_i,1)+1, 1), STR(people.w,1))
    IF (SEEK(cod, 'codwdr') AND !EMPTY(codwdr.sex)) AND m.sex != codwdr.sex
     m.recid = recid
     rval = InsError('S', 'MDA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'MDA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.H3A == .T.
    IF !SEEK(d_type, 'ososch')
     m.recid = recid
     rval = InsError('S', 'H3A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'H3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE 
     IF (IsMes(cod) AND ((Tip='5' AND d_type!='5') OR (Tip!='5' AND  d_type='5'))) OR ;
      cod=1561 AND d_type!='5'
      m.recid = recid
      rval = InsError('S', 'H3A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'H3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF
   ENDIF 
   
   IF M.SOA == .T.
    m.tiplpu = IIF(VAL(SUBSTR(m.mcod,3,2))<41, 'p', 's')
    IF !SEEK(SUBSTR(otd,2,2), 'profot')
     m.recid = recid
     rval = InsError('S', 'SOA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'SOA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
*     IF profot.stac='s' AND m.tiplpu=='p'
*      m.recid = recid
*      rval = InsError('S', 'SOA', m.recid)
**      InsErrorSV(m.mcod, 'S', 'SOA', m.recid)
*      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     ENDIF 
    ENDIF 
   ENDIF 

   IF M.R1A == .T.
    IF EMPTY(ishod)
     m.recid = recid
     rval = InsError('S', 'R1A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'R1A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(ishod, 'isv012')
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
*    IF (BETWEEN(cod,1900,1909) OR BETWEEN(cod,101927,101951)) AND ishod!=304
*     m.recid = recid
*     rval = InsError('S', 'R1A', m.recid)
*     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
**     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(ishod,3)+CHR(13)+CHR(10),0+64,'R1A')
*    ENDIF 
*    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945)) AND INLIST(rslt,326,332,333,334,335,336)
*     m.recid = recid
*     rval = InsError('S', 'R1A', m.recid)
*     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*    ENDIF 
*    IF BETWEEN(cod,101946,101948) AND !(BETWEEN(rslt,337,341) OR rslt=396)
*     m.recid = recid
*     rval = InsError('S', 'R1A', m.recid)
*     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*    ENDIF 
    IF BETWEEN(cod,101927,101928) OR BETWEEN(cod,101933,101945) OR cod=101951 && ����������������
     IF !BETWEEN(rslt,332,336) AND rslt!=326
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,101946,101948) && ���������������
     IF !BETWEEN(rslt,337,341) AND rslt!=396
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,1900,1905) && ��������������� ��������
     IF !BETWEEN(rslt,317,318) AND !BETWEEN(rslt,352,353) AND !BETWEEN(rslt,355,358) AND !BETWEEN(rslt,321,325) AND rslt!=320
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF BETWEEN(cod,101929,101932) && ��������������� �����-����� � ����������
     IF !BETWEEN(rslt,317,318) AND !BETWEEN(rslt,352,353) AND !BETWEEN(rslt,355,358) AND ;
     !BETWEEN(rslt,321,325) AND !BETWEEN(rslt,347,351) AND rslt!=320 AND rslt!=390
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 

   ENDIF 

   IF M.R2A == .T.
    IF EMPTY(rslt)
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(rslt, 'rsv009')
      m.recid = recid
      rval = InsError('S', 'R2A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'R2A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    IF BETWEEN(cod,1900,1905) AND !(BETWEEN(rslt,317,319) OR BETWEEN(rslt,352,354) OR BETWEEN(rslt,355,358))
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
    ENDIF 
    
    IF BETWEEN(cod,101929,101932) AND !(BETWEEN(rslt,321,325) OR BETWEEN(rslt,347,351) OR INLIST(rslt,320,390))
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
    ENDIF 
    
    IF BETWEEN(cod,1906,1909) AND !BETWEEN(rslt,343,345)
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
    ENDIF 
    
*    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945)) AND !(BETWEEN(rslt,332,336) OR rslt=326)
*     m.recid = recid
*     rval = InsError('S', 'R2A', m.recid)
*     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
**     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
*    ENDIF 
    
    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945)) AND rslt=304
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
    ENDIF 
    
*    IF BETWEEN(cod,101946,101948) AND !(BETWEEN(rslt,337,341) OR rslt=396)
*     m.recid = recid
*     rval = InsError('S', 'R2A', m.recid)
*     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
**     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
*    ENDIF 
    
    IF BETWEEN(cod,101946,101948) AND rslt=304
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
    ENDIF 
    

    IF BETWEEN(cod,101949,101950) AND rslt!=342
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*     MESSAGEBOX(PADL(cod,6,'0')+CHR(13)+CHR(10)+STR(rslt,3)+CHR(13)+CHR(10),0+64,'R2A')
    ENDIF 

   ENDIF 

   IF M.R3A == .T.
    IF EMPTY(prvs)
     m.recid = recid
     rval = InsError('S', 'R3A', m.recid)
*     InsErrorSV(m.mcod, 'S', 'R3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(prvs, 'kspec')
      m.recid = recid
      rval = InsError('S', 'R3A', m.recid)
*      InsErrorSV(m.mcod, 'S', 'R3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.NRA == .T.
    IF !INLIST(m.qcod, 'P2', 'S7', 'I3')
     m.cod = cod 
     m.sn_pol = sn_pol
     IF !(IsMes(m.cod) OR IsVmp(m.cod))
      IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p')
       IF (SEEK(m.sn_pol, 'people') AND !EMPTY(people.prmcod))
        m.recid = recid
        m.ord   = ord
        IF m.ord=0
         rval = InsError('S', 'NRA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 
      ENDIF 

     ELSE  && ���� ���

      m.recid = recid
      m.ord   = ord
      m.lpu_ord = lpu_ord
      IF INLIST(m.ord,1,2,6) AND m.lpu_ord=0
       rval = InsError('S', 'NRA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 

    ELSE && ���� �������, ��� ��� ������

     m.cod     = cod 
     m.sn_pol  = sn_pol
     m.facotd  = SUBSTR(otd,2,2)
     m.profil  = profil
     m.lpu_ord = lpu_ord

     DO CASE 
      CASE IsMes(m.cod)
       m.recid = recid
       m.ord   = ord
       m.lpu_ord = lpu_ord
       IF !INLIST(m.ord,1,2,3,6)
        rval = InsError('S', 'NRA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

      CASE IsVmp(m.cod)
       m.recid = recid
       m.ord   = ord
       m.lpu_ord = lpu_ord
       IF !INLIST(m.ord,1,2)
        rval = InsError('S', 'NRA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

      OTHERWISE 

       IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p') AND !INLIST(m.facotd,'01','08', '91', '92', '70', '73') ;
        AND m.profil!='100' AND !IsPat(m.cod) AND !IsSimult(m.cod)
        IF (SEEK(m.sn_pol, 'people') AND !EMPTY(people.prmcod) AND people.prmcod!=people.mcod)
         m.recid = recid
         m.ord   = ord
         IF !INLIST(m.ord,4,6,8,9) AND !(m.ord=7 AND m.lpu_ord=7665)
          rval = InsError('S', 'NRA', m.recid)
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ELSE 
         m.recid = recid
         m.ord   = ord
         IF !INLIST(m.ord,0,4,6,8,9) AND !(m.ord=7 AND m.lpu_ord=7665)
          rval = InsError('S', 'NRA', m.recid)
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ENDIF 
       ENDIF 

     ENDCASE 

    ENDIF  && ���� �������, ��� ��� ������

   ENDIF 

   IF M.G1A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    m.lpu_ord = lpu_ord
    m.ordmcod = IIF(SEEK(m.lpu_ord, 'sprlpu'), sprlpu.mcod, '')
    m.facotd  = SUBSTR(otd,2,2)
    m.profil  = profil
    IF !(IsMes(m.cod) OR IsVmp(m.cod))
     IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p') AND !INLIST(m.facotd,'01','08', '91', '70', '73') ;
        AND m.profil!='100' AND !IsPat(m.cod) AND !IsSimult(m.cod)
      m.ord     = ord
      m.lpu_ord = lpu_ord
      m.recid   = recid
      DO CASE 
       CASE m.ord = 4
        IF IIF(m.qcod='P2', m.ordmcod!=people.prmcod, !SEEK(m.lpu_ord, 'sprlpu')) AND m.lpu_ord!=4708
*        IF !SEEK(m.lpu_ord, 'sprlpu') AND m.lpu_ord!=4708
         rval = InsError('S', 'G1A', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       CASE INLIST(m.ord,6,9)
        IF m.lpu_ord!=9999
         rval = InsError('S', 'G1A', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       CASE m.ord =8
        IF m.lpu_ord!=8888
         rval = InsError('S', 'G1A', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       OTHERWISE 
      ENDCASE 
     ENDIF 
    ELSE  && ���� ��� ��� ���
     m.ord     = ord
     m.lpu_ord = lpu_ord
     m.recid   = recid
     DO CASE 
      CASE m.ord = 1
*       IF IIF(m.qcod='P2', m.ordmcod!=people.prmcod, !SEEK(m.lpu_ord, 'sprlpu')) AND m.lpu_ord!=4708
       IF !SEEK(m.lpu_ord, 'sprlpu') AND  m.lpu_ord!=9999
        rval = InsError('S', 'G1A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      CASE m.ord = 2
       IF !INLIST(m.lpu_ord,4708,9999)
        rval = InsError('S', 'G1A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      CASE m.ord=6
       IF m.lpu_ord!=9999
        rval = InsError('S', 'G1A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      OTHERWISE 
     ENDCASE 
    ENDIF 
   ENDIF 

   IF M.G2A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    
    IF qcod!='P2'
    
    IF !(IsMes(m.cod) OR IsVmp(m.cod))
     IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p')
      m.ord      = ord
      m.date_ord = date_ord
      m.recid    = recid
      m.d_u      = d_u
      IF INLIST(m.ord,4,6,8,9)
       IF EMPTY(m.date_ord) OR (!EMPTY(m.date_ord) AND m.date_ord>m.d_u)
        rval = InsError('S', 'G2A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
     ENDIF 
    ELSE 
     m.ord      = ord
     m.date_ord = date_ord
     m.recid    = recid
     m.d_u      = d_u
     IF INLIST(m.ord,1,2,6)
      IF EMPTY(m.date_ord) OR (!EMPTY(m.date_ord) AND m.date_ord>m.d_u)
       rval = InsError('S', 'G2A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
    
    ELSE 
    
     m.ord      = ord
     m.date_ord = date_ord
     m.recid    = recid
     m.d_u      = d_u
     IF INLIST(m.ord,1,4,6,8,9)
      IF EMPTY(m.date_ord) OR (!EMPTY(m.date_ord) AND m.date_ord>m.d_u)
       rval = InsError('S', 'G2A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 

    ENDIF 
    
   ENDIF 

*   IF M.G3A == .T.
    m.ord     = ord
    m.cod     = cod 
    m.sn_pol  = sn_pol
    m.recid   = recid
    m.lpu_ord = lpu_ord
    m.n_u     = ALLTRIM(n_u)
    
    IF FIELD('n_vmp')='N_VMP'
     IF IsVmp(m.cod)
      m.n_vmp   = ALLTRIM(n_vmp)
      IF m.ord=1 AND EMPTY(m.n_vmp)
       rval = InsError('S', 'G3A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
*   ENDIF 

   IF M.G3A == .T.
    m.cod    = cod 
    m.sn_pol = sn_pol
    m.recid   = recid
    m.lpu_ord = lpu_ord
    m.n_u     = ALLTRIM(n_u)
    
*    IF FIELD('n_vmp')='N_VMP'
*     IF IsVmp(m.cod)
*      m.n_vmp   = ALLTRIM(n_vmp)
*      IF m.ord=1 AND EMPTY(m.n_vmp)
*       rval = InsError('S', 'G3A', m.recid)
*       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*      ENDIF 
*     ENDIF 
*    ENDIF 

    IF qcod!='P2'

    IF !(IsMes(m.cod) OR IsVmp(m.cod))
     IF (SEEK(m.cod, 'tarif') AND tarif.tpn='p')
      IF INLIST(m.ord,4,6,8,9) AND m.lpu_ord = 4708 AND EMPTY(m.n_u)
       rval = InsError('S', 'G3A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ELSE 
     IF INLIST(m.ord,1,2,6) AND m.lpu_ord = 4708 AND EMPTY(m.n_u)
      rval = InsError('S', 'G3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 

    ELSE 

    IF IsMes(m.cod) OR IsVmp(m.cod)
     IF m.lpu_ord = 4708 AND EMPTY(m.n_u)
      rval = InsError('S', 'G3A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 

    ENDIF 

   ENDIF 

   IF M.G4A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    IF IsMes(m.cod) OR IsVmp(m.cod)
     m.recid = recid
     m.ord   = ord
     m.ds_0 = ALLTRIM(ds_0)
     IF INLIST(m.ord,1,2,6) AND EMPTY(m.ds_0)
      rval = InsError('S', 'G4A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.UVA == .T.
    m.ldr = people.dr
    m.ddr = IIF(OCCURS('#',c_i)>=2, ;
     CTOD(SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),1,4)), ;
     people.dr)
    IF !EMPTY(m.ddr)
     nmonthes = ((m.dat2-m.ddr)/365.25)*12
*     nmonthes = ((m.ldr-m.ddr)/365.25)*12
     IF !BETWEEN(nmonthes, CodWDr.min_ms, CodWDr.max_ms) AND ;
       (!INLIST(d_type,'1','2','5') AND ;
       (!SEEK(ds, 'nocodr', 'ds1') AND !SEEK(ds, 'nocodr', 'ds2') AND !SEEK(ds, 'nocodr', 'ds3')))
      m.recid = recid
      rval = InsError('S', 'UVA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'UVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.DVA == .T.
    IF !EMPTY(Tip) AND !INLIST(SUBSTR(PADL(Cod,6,'0'),2,2), '83', '84')
     IF IsMes(Cod)
      m.perem = IIF(!ISDIGIT(SUBSTR(Ds,5,1)), STR(Cod,6)+' '+LEFT(Ds,3)+'   ', STR(Cod,6)+' '+Ds)
      IF (!SEEK(m.perem,'MesMkb') AND !INLIST(d_type,'1','5'))
       m.recid = recid
       rval = InsError('S', 'DVA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ELSE && ���� ���
      m.IsErr=.t.
      FOR m.opl=0 TO 3
       m.perem = STR(Cod,6)+' '+SUBSTR(Ds,1,6-m.opl)+SPACE(m.opl)
       IF SEEK(m.perem,'MesMkb')
        m.IsErr=.f.
        EXIT 
       ENDIF 
      ENDFOR 
      IF m.IsErr=.t.
       m.recid = recid
       rval = InsError('S', 'DVA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   && �������� �� "y" && ������ ����� ������ � ���� ���������
   IF M.UOA == .T.
    SET ORDER TO notd IN CodOtd
    IsCheck = IIF(SEEK(SUBSTR(otd,2,2), 'CodOtd', 'notd'), .T., .F.)
    IF IsCheck AND d_type!='2'
     IsOk = .f.
     DO WHILE SUBSTR(otd,2,2) = CodOtd.otd
      IF cod = CodOtd.Cod
       IsOk = .t.
       EXIT
      ENDIF
      SKIP IN CodOtd
     ENDDO 
    
     IF IsOk = .f.
      m.recid = recid
      rval = InsError('S', 'UOA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'UOA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
      
   IF M.NOA == .T.
    m.perem = sn_pol+str(cod,6)+dtos(d_u)
    IF SEEK(m.perem, 'e_day') AND d_type!='2'
     m.ocntr = e_day.cntr
     m.ncntr = m.ocntr + k_u
     IF m.ncntr<=e_day.k_norm
      REPLACE e_day.cntr WITH m.ncntr IN e_day
     ELSE 
     m.recid = recid
     rval = InsError('S', 'NOA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'NOA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.NMA == .T.
    m.perem = sn_pol+str(cod,6)
    IF SEEK(m.perem, 'e_month') AND d_type!='2'
     m.ocntr = e_month.cntr
     m.ncntr = m.ocntr + k_u
     IF m.ncntr<=e_month.k_norm
      REPLACE e_month.cntr WITH m.ncntr IN e_month
     ELSE 
     m.recid = recid
     rval = InsError('S', 'NMA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'NMA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    ENDIF 
   ENDIF 

   IF M.D2A == .T.
    m.cod = cod
    m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)

    IF m.dsptip > 0
    
    DO CASE 
     CASE m.dsptip = 1  && �������������� ��������, tip=1
      m.lastt = 12*3
     CASE m.dsptip = 2 && ����������� ��������, tip=2
      m.lastt = 12*2
     CASE m.dsptip = 3 && ��������������� �����, tip=3
      m.lastt = 12
     CASE m.dsptip = 4 && ����������� �����, tip=4
      m.lastt = IIF(m.qcod != 'P2', 12, 3)
     CASE m.dsptip = 5 && ���������������, tip=5
      m.lastt = 3
     CASE m.dsptip = 6 && �������������, tip=6
      m.lastt = 3
     OTHERWISE 
      m.lastt = 0
    ENDCASE 

    IF BETWEEN(cod,101933,101945) OR cod=101951
     FOR m.varcod = 101927 TO 101932
      m.perem = m.mcod+LEFT(sn_pol,17)+PADL(m.varcod,6,'0')
      IF m.qcod!='P2'
       IF SEEK(m.perem, 'dspyear')
        m.recid = recid
        rval = InsError('S', 'D2A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ELSE
       IF SEEK(m.perem, 'dspp') AND (d_u - dspp.d_u)/30<m.lastt
        m.recid = recid
        rval = InsError('S', 'D2A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
     ENDFOR 
    ENDIF 

    IF m.lastt>0
     m.perem = m.mcod+LEFT(sn_pol,17)+PADL(cod,6,'0')
     IF SEEK(m.perem, 'dspp') AND (d_u - dspp.d_u)/30<m.lastt
      m.recid = recid
      rval = InsError('S', 'D2A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
   ENDIF && IF m.dsptip>0
   
   ENDIF && IF M.D2A == .T.

   IF M.NUA == .T.
    SET ORDER TO ncod IN sovmno
    IF SEEK(cod, 'sovmno') && �������� NU - ������������� ������ 
     DO WHILE sovmno.cod == cod
      IF SEEK(sn_pol+STR(sovmno.cod_1,6)+DTOS(d_u), 'talon_exp')
       IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
        (IsStac AND UPPER(sovmno.Stac)='S')) AND (d_type != '2' OR talon_exp.d_type != '2')
        m.recid = recid
        rval = InsError('S', 'NUA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'NUA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        m.recid = talon_exp.recid
        rval = InsError('S', 'NUA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'NUA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
      SKIP +1 IN sovmno 
     ENDDO 
    ENDIF
   ENDIF  

   IF M.NSA == .T.
    SET ORDER TO scod IN sovmno
    IF SEEK(cod, 'sovmno') && �������� NS - ������������� ������ 
     IsSovmUsl = .F.
     DO WHILE sovmno.cod == cod
      IF SEEK(sn_pol+STR(sovmno.cod_1,6)+DTOS(d_u), 'talon_exp')
       IsSovmUsl = .T.
       EXIT 
      ENDIF 
      SKIP +1 IN sovmno 
     ENDDO 
     IF !IsSovmUsl
      IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
         (IsStac AND UPPER(sovmno.Stac)='S')) AND d_type != '2'
       m.recid = recid
       rval = InsError('S', 'NSA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'NSA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.HNA == .T.
    m.cod    = cod
    m.sn_pol = sn_pol
    m.d_u    = d_u
    IF INLIST(m.cod,15001,115001) AND SEEK(m.sn_pol, 'polic_h') AND m.d_u - polic_h.d_u < 365
     m.recid = recid
     rval = InsError('S', 'HNA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HNA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.HNA == .T.
    m.cod    = cod
    m.sn_pol = sn_pol
    m.d_u    = d_u
    IF INLIST(m.cod,101927,101928) AND SEEK(m.sn_pol, 'polic_dp')
     m.recid = recid
     rval = InsError('S', 'HNA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'HNA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.POA == .T. && ��������������� � ������������ ���
    m.cod    = cod
*    IF INLIST(m.cod, 101927, 101928) AND !SEEK(m.mcod, 'lpu_m')
    IF INLIST(m.cod, 101927, 101928)
     m.recid = recid
     rval = InsError('S', 'POA', m.recid)
*     InsErrorSV(m.mcod, 'S', 'POA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.VDA == .T. AND m.tdat1>={01.05.2014} && ������������ ����������
    m.pcod = pcod
*    m.prvs = ALLTRIM(prvs)
    m.prvs  = prvs
    m.d_ser = {}
    m.d_u   = d_u
    IF SEEK(m.pcod, 'doctor')
       m.d_ser = doctor.d_ser
*     DO CASE 
*      CASE m.prvs == ALLTRIM(doctor.prvs_1)
*       m.d_ser = doctor.d_ser_1
*      CASE m.prvs == ALLTRIM(doctor.prvs_2)
*       m.d_ser = doctor.d_ser_2
*      CASE m.prvs == ALLTRIM(doctor.prvs_3)
*       m.d_ser = doctor.d_ser_3
*      CASE m.prvs == ALLTRIM(doctor.prvs_4)
*       m.d_ser = doctor.d_ser_4
*     ENDCASE 
    ENDIF 
    IF !EMPTY(m.d_ser)
     IF m.d_u-m.d_ser > 365.25*5
      m.recid = recid
      rval = InsError('S', 'VDA', m.recid)
*      InsErrorSV(m.mcod, 'S', 'VDA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

  ENDIF && ���������� ���������� ������ �����!

  ENDSCAN

  IF talon.IsPr == .F. && ���������� ���������� ������ �����!
  IF M.SMA = .T.
  SELECT Gosp && �������� �� ���������� DU,SM,DI,DD.
  DO WHILE !EOF()
   Karta = c_i
   DO WHILE c_i = Karta
    DO CASE
     CASE !BETWEEN(d_u, dat1, dat2) And !EMPTY(Tip) && ���� ������� ��� �������
       DO WHILE c_i = Karta
        m.recid = recid
        rval = InsError('S', 'DUA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'DUA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        SKIP
       ENDDO 

     CASE BETWEEN(d_u, dat1, dat2) And !EMPTY(Tip) && ���� ������� � �������
      DVip  = d_u
      DPost = d_u
      kMS   = 1
      DO WHILE c_i = Karta
       DO CASE
        CASE !EMPTY(Tip) And d_u = DPost And !kMs=1 And !(kMS=2 And Cod=83010 And k_u=1) && ��� ��� �� ���� �������
         m.recid = recid
         rval = InsError('S', 'SMA', m.recid)
*         InsErrorSV(m.mcod, 'S', 'SMA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        
        CASE !EMPTY(Tip) And d_u < DVip && ������
         m.recid = recid
         rval = InsError('S', 'DIA', m.recid)
*         InsErrorSV(m.mcod, 'S', 'DIA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

*        CASE !EMPTY(Tip) And !INLIST(Cod,83010,83020,83030,83040,83050,183010,183020) And d_u > DVip && �����������
        CASE !EMPTY(Tip) And !INLIST(INT(Cod/1000),83,183) And d_u > DVip && �����������
         m.recid = recid
         rval = InsError('S', 'DDA', m.recid)
*         InsErrorSV(m.mcod, 'S', 'DDA', m.recid)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

        OTHERWISE
         DVip  = DVip - k_u
         DPost = d_u

       ENDCASE
       SKIP
       kMS = kMS + 1
     ENDDO 
    OTHERWISE 
     SKIP 
    ENDCASE
   ENDDO 
  ENDDO  
  SELECT Talon
  ENDIF  && ���������� ������ SMA
  ENDIF && ���������� ���������� ������ �����!
  
  IF talon.IsPr == .F. && ���������� ���������� ������ �����!
   IF M.DLA == .T.
    SET ORDER TO Unik
    GO TOP
    DO WHILE !EOF()
     Vir = sn_pol+otd+ds+Padl(cod,6,'0')+DToC(d_u)
     SKIP
     jjj = .T.
     DO WHILE sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) = vir
      IF jjj = .T.
       jjj = .F.
       SKIP -1
       IF d_type != '2'
        m.recid = recid
        rval = InsError('S', 'DLA', m.recid)
*        InsErrorSV(m.mcod, 'S', 'DLA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       SKIP 1
      ENDIF
      IF d_type != '2'
       m.recid = recid
       rval = InsError('S', 'DLA', m.recid)
*       InsErrorSV(m.mcod, 'S', 'DLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      SKIP 
     ENDDO  
    ENDDO 
   ENDIF 
  ENDIF 

  IF talon.IsPr == .F. && ���������� ���������� ������ �����!
*   MESSAGEBOX('OK',0+64,'')
   IF M.PPA == .T.
    IF SEEK(m.lpuid, 'lpudogs')
     m.syear = 0 
     m.syear = lpudogs.kv01+lpudogs.kv02+lpudogs.kv03+lpudogs.kv04
     IF m.syear>0
      m.sfact = 0
      m.beginm = 1
      DO CASE 
       CASE BETWEEN(m.tmonth,1,3)
*        m.beginm = 1
        m.kvlimit = lpudogs.kv01
       CASE BETWEEN(m.tmonth,4,6)
*        m.beginm = 4
        m.kvlimit = lpudogs.kv01 + lpudogs.kv02
       CASE BETWEEN(m.tmonth,7,9)
*        m.beginm = 7
        m.kvlimit = lpudogs.kv01 + lpudogs.kv02 + lpudogs.kv03
       CASE BETWEEN(m.tmonth,10,12)
*        m.beginm = 10
        m.kvlimit = lpudogs.kv01 + lpudogs.kv02 + lpudogs.kv03 + lpudogs.kv04
      ENDCASE 
      FOR m.nmon = m.beginm TO m.tmonth-1
       m.lcperiod = STR(m.tyear,4)+PADL(m.nmon,2,'0')
       m.lppath = pbase+ '\'+m.lcperiod
       IF fso.FolderExists(m.lppath)
        IF fso.FileExists(m.lppath+'\aisoms.dbf')
         IF OpenFile(m.lppath+'\aisoms', 'lcais', 'shar', 'lpuid')=0
          IF SEEK(m.lpuid, 'lcais')
           m.sfact = m.sfact + (lcais.s_pred-lcais.sum_flk)
          ENDIF 
          USE IN lcais
         ELSE 
          IF USED('lcais')
           USE IN lcais
          ENDIF 
         ENDIF 
        ENDIF 
       ENDIF 
      NEXT 

*      m.sfact = m.sfact + (aisoms.s_pred-aisoms.sum_flk)
      
*      IF m.sfact > m.kvlimit
       oord = ORDER()
       SET ORDER TO d_u
       SCAN 
        m.sfact = m.sfact + s_all
        IF m.sfact<=m.kvlimit
         LOOP 
        ENDIF 
        m.recid = recid
        rval = InsError('S', 'PPA', m.recid)
       ENDSCAN
       SET ORDER TO &oord

*      ENDIF 
     
     ENDIF 
    ENDIF 
   ENDIF 
  ENDIF 
  
  
  CREATE CURSOR AllBad (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol 
  SELECT talon 
  SET ORDER TO sn_pol
  GO TOP 
  
  DO WHILE !EOF()
   m.polis = sn_pol
   m.lAllBad=.t.
   DO WHILE sn_pol = m.polis
    m.recid = recid
    IF !SEEK(m.RecId, 'sError')
     m.lAllBad = .f.
*     MESSAGEBOX(TRANSFORM(m.RecId,'999999'),0+64,sn_pol)
    ENDIF 
    SKIP 
   ENDDO 
   IF m.lAllBad
    IF !SEEK(m.polis, 'Allbad')
     INSERT INTO Allbad (sn_pol) VALUES (m.polis)
    ENDIF 
   ENDIF 
  ENDDO 
  
*  SELECT allbad
*  BROWSE 
  
  SELECT People
  SCAN 
   m.polis = sn_pol
   m.recid = recid
   IF !SEEK(m.polis, 'allbad')
    LOOP 
   ENDIF 
   IF !SEEK(RecId, 'rError')
     =InsError('R', 'PNA', m.recid)
   ENDIF 
  ENDSCAN 
  USE IN AllBad
  
  SELECT talon 
  SUM(s_all) FOR SEEK(RecId, 'sError') TO m.s_flk
  SET RELATION OFF INTO people
  
  =ClBase()

  SELECT AisOms
  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
  SELECT AisOms
  
  REPLACE sum_flk WITH m.s_flk

 WAIT CLEAR 
* MESSAGEBOX('��������� ���������!', 0+64, '')

RETURN 

FUNCTION InsError(WFile, cError, cRecId)
 IF WFile == 'R'
  IF !SEEK(cRecId, 'rError')
   INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
  ELSE 
*   IF cError != rError.c_err
*    INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
*   ENDIF cError != rError.c_err
  ENDIF !SEEK(cRecId, 'rError')
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(cRecId, 'sError')
   INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   ENDIF cError != sError.c_err 
  ENDIF !SEEK(cRecId, 'sError')
 ENDIF 
RETURN .F.

FUNCTION InsErrorSV(mmcod, WFile, cError, cRecId)
 IF WFile == 'R'
  IF !SEEK(mmcod+STR(cRecId,9), 'resv')
   INSERT INTO resv (mcod, f, c_err, rid) VALUES (mmcod, 'R', cError, cRecId)
  ELSE 
   IF cError != resv.c_err
    INSERT INTO resv (mcod, f, c_err, rid) VALUES (mmcod, 'R', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(mcod+STR(cRecId,9), 'sesv')
   INSERT INTO sesv (mcod, f, c_err, rid) VALUES (mmcod, 'S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sesv.c_err
    INSERT INTO sesv (mcod, f, c_err, rid) VALUES (mmcod, 'S', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
RETURN .F.

FUNCTION OpBase(ppath)
 tnresult = 0
 tnresult = tnresult + OpenFile(ppath+'\people', 'people', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(ppath+'\talon', 'talon', 'share')
 tnresult = tnresult + OpenFile(ppath+'\doctor', 'doctor', 'share', 'pcod')
 IF fso.FileExists(ppath+'\ho'+m.qcod+'.dbf')
  tnresult = tnresult + OpenFile(ppath+'\ho'+m.qcod, 'ho', 'share', 'unik')
 ENDIF 
 tnresult = tnresult + OpenFile(ppath+'\e'+m.mcod, 'rerror', 'share', 'rrid')
 tnresult = tnresult + OpenFile(ppath+'\e'+m.mcod, 'serror', 'share', 'rid', 'again')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoree', 'osoree', 'share', 'd_type')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mkb10', 'mkb10', 'share', 'ds')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'share', 'cod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodWDr', 'CodWDr', 'share', 'cod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\ososch', 'ososch', 'share', 'd_type')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\profot', 'profot', 'share', 'otd')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\isv012', 'isv012', 'share', 'ishod')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\rsv009', 'rsv009', 'share', 'rslt')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\'+IIF(m.tdat1>={01.05.2014}, 'spv015', 'kspec'), 'kspec', 'share', IIF(m.tdat1>={01.05.2014}, 'code', 'prvs'))
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\NoCodR', 'NoCodR', 'share')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\ms_mkb', 'MesMkb', 'share', 'ds_ms')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodOtd', 'CodOtd', 'share')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\CodKU', 'CodKU', 'share')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', 'OsoERZ', 'Shar', 'ans_r')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sovmno', 'sovmno', 'Shar', 'ncod')
 tnresult = tnresult + OpenFile(ppath+'\talon', 'talon_exp', 'share', 'ExpTag', 'again')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_h', 'polic_h', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_dp', 'polic_dp', 'share', 'sn_pol')
 tnresult = tnresult + OpenFile(pcommon+'\dsdisp', 'dsdisp', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\kpresl', 'kpresl', 'share', 'tip')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\mo_vmp', 'movmp', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spi_lpu_dd', 'spidd', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\noth', 'noth', 'share', 'cod')
 tnresult = tnresult + OpenFile(pcommon+'\lpudogs', 'lpudogs', 'share', 'lpu_id')
 tnresult = tnresult + OpenFile(pcommon+'\dspcodes', 'dspcodes', 'share', 'cod')

 DELETE FOR SUBSTR(c_err,3,1)='A' IN rerror
 DELETE FOR SUBSTR(c_err,3,1)='A' IN serror
RETURN .t.

FUNCTION ClBase()
 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('doctor')
  USE IN doctor
 ENDIF 
 IF USED('rerror')
  USE IN rerror
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
 IF USED('osoree')
  USE IN osoree
 ENDIF 
 IF USED('mkb10')
  USE IN mkb10
 ENDIF 
 IF USED('tarif')
  USE IN tarif
 ENDIF 
 IF USED('codwdr')
  USE IN CodWDr
 ENDIF 
 IF USED('ososch')
  USE IN OsoSch
 ENDIF 
 IF USED('profot')
  USE IN ProfOt
 ENDIF 
 IF USED('isv012')
  USE IN isv012
 ENDIF 
 IF USED('rsv009')
  USE IN rsv009
 ENDIF 
 IF USED('kspec')
  USE IN kspec
 ENDIF 
 IF USED('nocodr')
  USE IN NoCodR
 ENDIF 
 IF USED('mesmkb')
  USE IN MesMkb
 ENDIF 
 IF USED('codotd')
  USE IN CodOtd
 ENDIF 
 IF USED('codku')
  USE IN CodKU
 ENDIF 
 IF USED('day_gr')
  USE IN Day_gr
 ENDIF 
 IF USED('e_day')
  USE IN e_day 
 ENDIF 
 IF USED('month_gr')
  USE IN Month_Gr
 ENDIF 
 IF USED('e_month')
  USE IN e_month
 ENDIF 
 IF USED('osoerz')
  USE IN OsoERZ
 ENDIF 
 IF USED('sovmno')
  USE IN sovmno
 ENDIF 
 IF USED('talon_exp')
  USE IN talon_exp
 ENDIF 
 IF USED('gosp')
  USE IN Gosp
 ENDIF 
 IF USED('polic_h')
  USE IN polic_h
 ENDIF 
 IF USED('polic_dp')
  USE IN polic_dp
 ENDIF 
 IF USED('dsdisp')
  USE IN dsdisp
 ENDIF 
 IF USED('kpresl')
  USE IN kpresl
 ENDIF 
 IF USED('movmp')
  USE IN movmp
 ENDIF 
 IF USED('spidd')
  USE IN spidd
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('dspp')
  USE IN dspp
 ENDIF 
 IF USED('dspyear')
  USE IN dspyear
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 IF USED('noth')
  USE IN noth
 ENDIF 
 IF USED('ho')
  USE IN ho
 ENDIF 
RETURN 