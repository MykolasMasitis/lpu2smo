FUNCTION MakeYFilesOne
 PARAMETERS para1
 PRIVATE m.lcpath
  m.lcpath = para1
  hfile   = 'h'+STR(m.lpu_id,4)+'.'+mmy
  dfile   = 'd'+STR(m.lpu_id,4)+'.'+mmy
  nvfile  = 'nv'+STR(m.lpu_id,4)+'.'+mmy
  nsfile  = 'ns'+STR(m.lpu_id,4)+'.'+mmy
  sprfile = 'spr' + STR(m.lpu_id,4) + '.' + mmy
  ryfile  = 'R'+m.qcod+'Y.'+mmy
  syfile  = 'S'+m.qcod+'Y.'+mmy
  hofile  = 'ho'+m.qcod+'.'+mmy

  IF !fso.FileExists(lcpath + '\people.dbf') OR ;
     !fso.FileExists(lcpath + '\talon.dbf') OR ;
     !fso.FileExists(lcpath + '\e'+m.mcod+'.dbf')
   RETURN 
  ENDIF 
  
  fso.CopyFile(pTempl+'\rqqy.mmy', lcpath+'\'+ryfile, .t.)
  fso.CopyFile(pTempl+'\sqqy.mmy', lcpath+'\'+syfile, .t.)
  
  CREATE CURSOR curdbls (sn_pol c(25), k_u n(2))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  =OpenFile(lcpath+'\'+ryfile, "ryfile", "excl")
  SELECT ryfile 
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol IN ryfile
  =OpenFile(lcpath+'\'+syfile, "syfile", "share")
  
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people', 'people', 'share')
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'share')
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'rerror', 'share', 'rrid')
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'serror', 'share', 'rid', 'again')
  
  CREATE CURSOR deftln (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  SELECT people
  SET RELATION TO RecId INTO rerror
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recno = RECNO()
   m.recid = PADL(m.recno,6,'0')
   
   m.tip_p     = m.tipp
   
   IF rerror.c_err='PNA' AND m.qcod='S7' && 'PNA'
    m.er_c      = ''
    m.refreason = ''
    m.osn230    = ''
   ELSE 
    m.er_c      = rerror.c_err
    m.refreason = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.refreason, '')
    m.osn230    = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.osn230, '')
   ENDIF 

   m.et        = 'A'
   m.et230     = 1 && ≈ÒÎË Ã› !
   m.oplata    = IIF(EMPTY(m.er_c), 1, 2) && 1-ÔÓÎÌ‡ˇ ÓÔÎ‡Ú‡, 2-ÔÓÎÌ˚È ÓÚÍ‡Á, 3-˜‡ÒÚË˜ÌÓÂ ÒÌˇÚËÂ
   m.prmcod    = prmcod
   m.lpu_prik  = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)

   IF !SEEK(m.sn_pol, 'ryfile')
    INSERT INTO ryfile FROM MEMVAR 
   ELSE 
    IF !SEEK(m.sn_pol, 'curdbls')
     INSERT INTO curdbls (sn_pol, k_u) VALUES (m.sn_pol, 1)
    ELSE 
     m.o_k_u = curdbls.k_u
     m.n_k_u = m.o_k_u + 1
     UPDATE curdbls SET k_u=m.n_k_u
    ENDIF 
    m.nIsDoubles = m.nIsDoubles + 1
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO rerror
  
  m.sum_st2 = 0
  
  SELECT people
  SET ORDER TO sn_pol
  SELECT talon
  SET RELATION TO recid INTO serror
  SET RELATION TO sn_pol INTO people ADDITIVE 
  SCAN  
   SCATTER MEMVAR 
   m.iotd = m.otd
   m.recno = RECNO()
   m.recid = PADL(m.recno,6,'0')

   m.er_c      = serror.c_err
   m.refreason = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.refreason, '')
   m.osn230    = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.osn230, '')
   m.et        = 'A'
   m.et230     = 1 && ≈ÒÎË Ã› !

   m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod) OR IsEko(m.cod)), .T., .F.)
   
   IF !EMPTY(m.er_c)
    IF !SEEK(m.sn_pol, 'deftln')
     INSERT INTO deftln FROM MEMVAR 
    ENDIF 
   ENDIF 

*   m.s_all = fsumm(m.cod, m.tip, m.k_u, m.IsVed)

   m.sum_st2 = m.sum_st2 + IIF(EMPTY(m.er_c), m.s_all, 0)
  
   m.prmcod = people.prmcod
   m.prik   = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   
   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
   m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
   m.paztip = TipOfPaz(m.mcod, m.prmcod) && 0 (ÌÂ ÔËÍÂÔÎÂÌ),1 (ÔËÍÂÔÎÂÌ ÔÓ ÏÂÒÚÛ Ó·‡˘ÂÌËˇ),2 (Í ÔËÎÓÚÛ),3 (ÌÂ Í ÔËÎÓÚÛ)
   
   DO CASE 
    CASE IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod)
     m.f_type = 'ft' && ÔÓ Ú‡ËÙÛ

    CASE m.IsTpnR OR IsPat(m.cod) OR IsEKO(m.cod) OR INLIST(SUBSTR(m.otd,2,2),'70','73','93') OR m.d_type='s'
     m.f_type = 'fh' && ‰ÓÔÛÒÎÛ„Ë

    CASE (INLIST(SUBSTR(m.otd,2,2),'01','90') AND IsStac(m.mcod)) AND TipOfPr(m.mcod, m.prmcod) = 1
     m.f_type = 'fp' && ËÁ ÒÂ‰ÒÚ‚ ÔÓ‰Û¯Â‚Ó„Ó ÙËÌ‡ÌÒËÓ‚‡ÌËˇ

    CASE (m.ord=7 AND m.lpu_ord=7665) AND TipOfPr(m.mcod, m.prmcod) = 1
     m.f_type = 'fp' && ËÁ ÒÂ‰ÒÚ‚ ÔÓ‰Û¯Â‚Ó„Ó ÙËÌ‡ÌÒËÓ‚‡ÌËˇ
     
    OTHERWISE 

    DO CASE 
     CASE TipOfPr(m.mcod, m.prmcod) = 0 && ÌÂÔËÍÂÔÎÂÌ
      m.f_type = 'fp'
*      IF m.IsLpuTpn = .T.
*       IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*        m.f_type = 'fh'
*       ENDIF 
*      ENDIF 
      CASE TipOfPr(m.mcod, m.prmcod) = 2 && Í ÔËÎÓÚÛ
      IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(SUBSTR(otd,2,2),'08','92')))
       m.f_type = 'vz' && ‚Á‡ËÏÓÁ‡˜ÂÚ˚ (ËÁ ÒÂ‰ÒÚ‚ ËÌÓ„Ó Àœ”)
      ELSE 
       m.f_type = 'fp'
      ENDIF 
*      IF m.IsLpuTpn = .T.
*       m.fil_id = fil_id
*       IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*        m.f_type = 'fh'
*       ENDIF 
*      ENDIF 
      CASE TipOfPr(m.mcod, m.prmcod) = 3 && Ò‚ÓÈ
      m.f_type = 'fp'
*      IF m.IsLpuTpn = .T.
*       m.fil_id = fil_id
*       IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*        m.f_type = 'fh'
*       ENDIF 
*      ENDIF 
      OTHERWISE 
      m.f_type = ''
    ENDCASE 
   ENDCASE 
   
   IF !m.IsPilot
    IF IsKDS(m.cod)
     m.f_type=' '
    ELSE 
     m.f_type='ft'
    ENDIF 
   ENDIF 

   INSERT INTO syfile FROM MEMVAR 
  
  ENDSCAN 
  SET RELATION OFF INTO serror
  SET RELATION OFF INTO people
  USE 

  SELECT people
  SET ORDER TO recid
  SET RELATION TO sn_pol INTO deftln
  SELECT rerror
  SET RELATION TO rid INTO people ADDITIVE 
  SCAN 
   IF EMPTY(deftln.sn_pol)
    DELETE 
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO rerror
  USE 
*  USE IN rerror
  USE IN people 
  USE IN serror
  USE IN syfile
  SELECT ryfile
  SET ORDER TO 
  DELETE TAG ALL 
  SET RELATION TO sn_pol INTO deftln
  SCAN
   IF !EMPTY(er_c)
    IF EMPTY(deftln.sn_pol)
     REPLACE er_c WITH '', refreason WITH '', oplata WITH 1
    ENDIF 
   ENDIF  
  ENDSCAN 
  SET RELATION OFF INTO deftln
  USE 
  USE IN deftln
  
  IF m.sum_st1 != m.sum_st2
   MESSAGEBOX(''+CHR(13)+CHR(10)+;
    '¬Õ»Ã¿Õ»≈!'+CHR(13)+CHR(10)+;
    '—”ÃÃ¿, –¿—◊»“¿ÕÕ¿ﬂ  ¿  –¿«Õ»÷¿ Ã≈∆ƒ” œ–≈ƒ—“¿¬À≈ÕÕŒ… » ‘À , —Œ—“¿¬Àﬂ≈“'+CHR(13)+CHR(10)+;
    TRANSFORM(m.sum_st1,'99 999 999.99')+CHR(13)+CHR(10)+;
    '—”ÃÃ¿ œŒ ƒ¿ÕÕ€Ã œ≈–—ŒÕ»‘»÷»–Œ¬¿ÕÕŒ… Œ“◊≈“ÕŒ—“» —Œ—“¿¬Àﬂ≈“'+CHR(13)+CHR(10)+;
    TRANSFORM(m.sum_st2,'99 999 999.99')+CHR(13)+CHR(10), 0+48, m.mcod)
  ENDIF 
  
  IF m.nIsDoubles>0
   m.stroka = ''
   SELECT curdbls
   SCAN 
    m.sn_pol = sn_pol
    m.stroka = m.stroka + m.sn_pol + CHR(13)+CHR(10)
   ENDSCAN 
  ENDIF 
 
  IF m.nIsDoubles>0
   MESSAGEBOX('Œ¡Õ¿–”∆≈Õ€ ƒ”¡À» ¬ –≈√»—“–≈!'+CHR(13)+CHR(10)+;
   m.stroka+'œ≈–—Œ“◊≈“ Õ≈¬≈–≈Õ!',0+64,m.mcod)
  ENDIF 
  
  USE IN curdbls

  ZipPath = lcPath
  ZipName = 'D'+m.qcod+STR(m.lpu_id,4)+'.zip'
  MmyName = 'D'+m.qcod+STR(m.lpu_id,4)+'.'+mmy

  IF fso.FileExists(lcpath+'\'+ZipName)
   fso.DeleteFile(lcpath+'\'+ZipName)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+MmyName)
   fso.DeleteFile(lcpath+'\'+MmyName)
  ENDIF 

  SET DEFAULT TO (lcpath)
  
  bfile='b'+m.mcod+'.'+mmy
  IF fso.FileExists(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+bfile)
   UnZipOpen(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+bfile)

   UnzipGotoFileByName(dfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(hfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(nvfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(nsfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(sprfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(hofile)
   UnzipFile(lcPath+'\')

   UnZipClose()

  ENDIF 
  
  ZipOpen(MmyName, lcPath+'\')
  IF fso.FileExists(lcpath+'\'+ryfile)
   ZipFile(ryfile, .T.)
   fso.DeleteFile(lcpath+'\'+ryfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+syfile)
   ZipFile(syfile, .T.)
   fso.DeleteFile(lcpath+'\'+syfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+dfile)
   ZipFile(dfile, .T.)
   fso.DeleteFile(lcpath+'\'+dfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+nvfile)
   ZipFile(nvfile, .T.)
   fso.DeleteFile(lcpath+'\'+nvfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+nsfile)
   ZipFile(nsfile, .T.)
   fso.DeleteFile(lcpath+'\'+nsfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+sprfile)
   ZipFile(sprfile, .T.)
   fso.DeleteFile(lcpath+'\'+sprfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+hfile)
   ZipFile(hfile, .T.)
   fso.DeleteFile(lcpath+'\'+hfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+hofile)
   ZipFile(hofile, .T.)
   fso.DeleteFile(lcpath+'\'+hofile)
  ENDIF 

  ZipClose()
RETURN 