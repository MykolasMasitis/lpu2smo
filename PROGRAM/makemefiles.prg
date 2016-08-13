PROCEDURE MakeMEFiles
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
  'ME-файлы?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(13)+'ОТСУТСТВУЕТ ФАЙЛ AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'') 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\mee2mgf', 'mee2mgf', 'shar', 'my_et')>0
  IF USED('mee2mgf')
   USE IN mee2mgf
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pmee+'\svacts\svacts', 'svactsn', 'shar')>0
  IF USED('svactsn')
   USE IN svactsn
  ENDIF 
  USE IN mee2mgf
  USE IN pilot
  USE IN aisoms
  RETURN 
 ENDIF 
 
 SELECT * FROM svactsn INTO CURSOR cursvacts READWRITE 
* INDEX on period+e_period+mcod+STR(codexp,1)+docexp TAG unik
 INDEX on e_period+period+mcod+STR(codexp,1)+docexp TAG unik
 SET ORDER TO unik 
 COPY TO &pmee\cursvacts WITH cdx 
 USE IN svactsn
 
 IF !fso.FolderExists(pOut+'\'+m.gcPeriod)
  fso.CreateFolder(pOut+'\'+m.gcPeriod)
 ENDIF 

 oMailDir        = fso.GetFolder(pOut+'\'+m.gcPeriod)
 MailDirName     = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count

 FOR EACH oFileInMailDir IN oFilesInMailDir
  m.BFullName = oFileInMailDir.Path
  m.bname     = oFileInMailDir.Name
  m.recieved  = oFileInMailDir.DateLastModified
  
  IF LEN(m.bname)!=12
   LOOP 
  ENDIF 
  
  m.part01 = UPPER(LEFT(m.bname,2))
  m.part02 = UPPER(SUBSTR(m.bname,3,2))
  m.part03 = SUBSTR(m.bname,5,4)
  m.ext    = LOWER(RIGHT(m.bname,3))

  IF part01 != 'ME'
   LOOP 
  ENDIF 
  IF part02 != m.qcod
   LOOP 
  ENDIF 
  IF !INLIST(ext, 'dbf', 'zip')
   LOOP 
  ENDIF 
  IF !SEEK(INT(VAL(m.part03)),'aisoms')
   LOOP 
  ENDIF 
  
  fso.DeleteFile(m.BFullName)

 ENDFOR 
 
 USE IN aisoms

 CREATE CURSOR meexp ;
  (period c(6), e_period c(6), lpuid n(6), mcod c(7), et c(1), docexp c(7), sexp n(11,2), stpn n(11,2), s_sank n(11,2), n_akt c(15), d_akt d)

 m.tsexp = 0
 m.tssank = 0
 
 FOR nperiod=18 TO 0 STEP -1 
  expperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+expperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+expperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\aisoms', 'aisomse', 'shar')>0
   IF USED('aisomse')
    USE IN aisomse
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shared', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('aisomse')
    USE IN aisomse
   ENDIF 
   LOOP  
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')>0
   IF USED('sookod')
    USE IN sookod
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('aisomse')
    USE IN aisomse
   ENDIF 
   LOOP  
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   IF USED('sookod')
    USE IN sookod
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('aisomse')
    USE IN aisomse
   ENDIF 
   LOOP  
  ENDIF 
  m.lIsLpuTpn=.f.
  IF fso.FileExists(pbase+'\'+expperiod+'\'+'nsi'+'\lputpn.dbf')
   m.lIsLpuTpn=.t.
  ENDIF 
  IF m.lIsLpuTpn=.t.
   IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\lputpn', 'lputpn', 'shar', 'lpu_id')>0
    IF USED('lputpn')
     USE IN lputpn
    ENDIF 
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
    IF USED('sookod')
     USE IN sookod
    ENDIF 
    IF USED('tarif')
     USE IN tarif
    ENDIF 
    IF USED('aisomse')
     USE IN aisomse
    ENDIF 
    LOOP  
   ENDIF 
  ENDIF 
  
  SELECT aisomse
  WAIT "ОБРАБОТКА ПЕРИОДА "+expperiod+" ..." WINDOW NOWAIT 
  SCAN 
   m.mcod = mcod
   m.IsPilot = IIF(SEEK(m.mcod, 'pilot'), .t., .f.)
   m.lpuid = lpuid
   m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
   m.IsTpn = .f.
   IF m.lIsLpuTpn
    m.IsTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
   ENDIF 

   IF !fso.FolderExists(pbase+'\'+expperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+expperiod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+expperiod+'\'+m.mcod+'\people.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+expperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+expperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+expperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+expperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    LOOP 
   ENDIF 

   SELECT Talon
   SET RELATION TO sn_pol INTO people
   SELECT merror 
   SET RELATION TO recid INTO talon 
   
   m.IsExpMee = .f.

   m.sexp2 = 0
   m.sexp3 = 0
   m.sexp4 = 0
   m.sexp5 = 0
   m.sexp6 = 0
   m.sexp7 = 0

   m.stpn2 = 0
   m.stpn3 = 0
   m.stpn4 = 0
   m.stpn5 = 0
   m.stpn6 = 0
   m.stpn7 = 0

   m.ssank2 = 0
   m.ssank3 = 0
   m.ssank4 = 0
   m.ssank5 = 0
   m.ssank6 = 0
   m.ssank7 = 0

   SCAN 
    IF m.IsExpMee = .f. AND !EMPTY(err_mee) AND e_period = gcperiod
     m.IsExpMee = .t.
    
     MEFile  = 'ME'+UPPER(m.qcod)+STR(m.lpuid,4)
     IF !fso.FolderExists(pOut+'\'+gcPeriod)
      fso.CreateFolder(pOut+'\'+gcPeriod)
     ENDIF 

     MEFilep = pOut+'\'+gcPeriod+'\ME'+UPPER(m.qcod)+STR(m.lpuid,4)
    
     IF !fso.FileExists(MEFilep+'.dbf')
      fso.CopyFile(pTempl+'\MEqqnnnn.dbf', MEFileP+'.dbf')
     ENDIF 
    
     =OpenFile(MEFileP, 'mefile', 'share')
     
     SELECT merror
    ENDIF 

    IF !EMPTY(err_mee) AND e_period = m.gcperiod
     SCATTER MEMVAR 
     
     m.sn_pol = talon.sn_pol
     m.c_i    = talon.c_i
     m.ds     = talon.ds
     m.d_u    = talon.d_u
     m.pcod   = talon.pcod
     m.d_type = talon.d_type

     m.lpu_id   = m.lpuid
     m.fil_id   = talon.fil_id
     m.IsFilTpn = .f.
     IF m.IsTpn
      m.IsFilTpn = IIF(SEEK(m.fil_id, 'lputpn', 'fil_id'), .t., .f.)
     ENDIF 
     m.recid    = PADL(recid,6,'0')
     m.iotd     = talon.otd
     m.period   = m.gcperiod
     m.period_e = expperiod
     m.s_opl    = talon.s_all
     m.er_c     = err_mee
     m.osn230   = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), sookod.osn230, '')
     m.et       = IIF(EMPTY(m.et), '2', m.et)
     m.tip_e    = e_tip
     m.cod_e    = e_cod
     m.k_u_e    = e_ku
     m.ns_all   = 0
     m.delta    = 0
     
     m.s_sank = s_2
     
     m.prmcod = people.prmcod
     m.lpu_prik = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)

     IF m.koeff <= 0 && старый механизм

      IF EMPTY(e_cod) AND EMPTY(e_tip) AND EMPTY(e_ku)
       m.delta = s_all
      ENDIF  && Полное снятие!
      IF (!EMPTY(e_cod) AND cod != e_cod) OR (!EMPTY(e_ku) AND k_u != e_ku) OR (!EMPTY(e_tip) AND e_tip != tip)
       m.ns_all = fsumm(e_cod, e_tip, e_ku, m.IsVed, GOMONTH(tdat1,-nperiod))
       m.delta = s_all - m.ns_all
      ENDIF && Частичное снятие
     
     ELSE 

      m.delta = ROUND(talon.s_all * m.koeff,2)

     ENDIF 

     m.s_opl_e = m.delta
     m.s_opl_e = s_1
     
     IF err_mee='W0'
      m.er_c     = '99'
      m.osn230   = ''

      m.ds_e     = ''
      m.tip_e    = ''
      m.cod_e    = 0
      m.codnom_e = ''
      m.k_u_e    = 0
      m.s_opl_e  = 0
      
*      m.s_sank   = 0
     ENDIF 
    
     m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
     m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
     m.paztip = TipOfPaz(m.mcod, m.prmcod) && 0 (не прикреплен),1 (прикреплен по месту обращения),2 (к пилоту),3 (не к пилоту)

     IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) OR IsEKO(m.cod) OR IsPat(m.cod)
      m.f_type = 'ft'
      IF IsPat(m.cod) OR IsEKO(m.cod)
       m.f_type = 'fh'
      ENDIF 
     ELSE 
      DO CASE 
       CASE TipOfPr(m.mcod, m.prmcod) = 0 && неприкреплен
        m.f_type = 'fp'
        IF m.IsFilTpn = .T.
         IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
          m.f_type = 'fh'
         ENDIF 
        ENDIF 
        CASE TipOfPr(m.mcod, m.prmcod) = 2 && к пилоту
        IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(SUBSTR(talon.otd,2,2),'08','92')))
         m.f_type = 'vz'
        ELSE 
         m.f_type = 'fp'
        ENDIF 
        IF m.IsFilTpn = .T.
         m.fil_id = fil_id
         IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
          m.f_type = 'fh'
         ENDIF 
        ENDIF 
        CASE TipOfPr(m.mcod, m.prmcod) = 3 && свой
        m.f_type = 'fp'
        IF m.IsFilTpn = .T.
         m.fil_id = fil_id
         IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
          m.f_type = 'fh'
         ENDIF 
        ENDIF 
        OTHERWISE 
        m.f_type = ''
      ENDCASE 
     ENDIF 
   
     IF !m.IsPilot
      IF IsKDS(m.cod)
       m.f_type=' '
      ELSE 
       m.f_type='ft'
      ENDIF 
     ENDIF 

     m.sexp2 = m.sexp2 + IIF(m.et='2', m.s_opl_e, 0)
     m.sexp3 = m.sexp3 + IIF(m.et='3', m.s_opl_e, 0)
     m.sexp4 = m.sexp4 + IIF(m.et='4', m.s_opl_e, 0)
     m.sexp5 = m.sexp5 + IIF(m.et='5', m.s_opl_e, 0)
     m.sexp6 = m.sexp6 + IIF(m.et='6', m.s_opl_e, 0)
     m.sexp7 = m.sexp7 + IIF(m.et='7', m.s_opl_e, 0)

     m.stpn2 = m.stpn2 + IIF(m.et='2' and m.IsFilTpn, m.s_opl_e, 0)
     m.stpn3 = m.stpn3 + IIF(m.et='3' and m.IsFilTpn, m.s_opl_e, 0)
     m.stpn4 = m.stpn4 + IIF(m.et='4' and m.IsFilTpn, m.s_opl_e, 0)
     m.stpn5 = m.stpn5 + IIF(m.et='5' and m.IsFilTpn, m.s_opl_e, 0)
     m.stpn6 = m.stpn6 + IIF(m.et='6' and m.IsFilTpn, m.s_opl_e, 0)
     m.stpn7 = m.stpn7 + IIF(m.et='7' and m.IsFilTpn, m.s_opl_e, 0)
     
     m.ssank2 = m.ssank2 + IIF(m.et='2', m.s_sank, 0)
     m.ssank3 = m.ssank3 + IIF(m.et='3', m.s_sank, 0)
     m.ssank4 = m.ssank4 + IIF(m.et='4', m.s_sank, 0)
     m.ssank5 = m.ssank5 + IIF(m.et='5', m.s_sank, 0)
     m.ssank6 = m.ssank6 + IIF(m.et='6', m.s_sank, 0)
     m.ssank7 = m.ssank7 + IIF(m.et='7', m.s_sank, 0)

     m.et_old = m.et
     m.et = IIF(SEEK(m.et, 'mee2mgf'), mee2mgf.mgf_et, m.et)

     m.extip = IIF(INLIST(m.et_old,'2','3','7'),'1','2')
     m.videxp = IIF(INLIST(m.et_old,'2','4','6','7'),'1','2') && 1 - плановая, в т.ч. тематическая.
     m.podtip = IIF(INLIST(m.et_old,'2','3','7'),'0','1') && это неправильно! сделано временно!
*     e_period+period+mcod+STR(codexp,1)+docexp

     m.vvir = m.gcperiod+m.period_e+m.mcod+m.et_old+m.docexp
     IF SEEK(m.vvir, 'cursvacts')
*      MESSAGEBOX(m.gcperiod+m.period_e+m.mcod+m.et_old+"."+m.docexp+".",0+64,'Yes')
     ELSE 
*      MESSAGEBOX(m.gcperiod+m.period_e+m.mcod+m.et_old+"."+m.docexp+".",0+64,'No')
     ENDIF 

     m.actnum = IIF(SEEK(m.vvir, 'cursvacts'), PADL(cursvacts.recid,6,'0'), '0')
     m.actdat = IIF(SEEK(m.vvir, 'cursvacts'), TTOD(cursvacts.actdate), {})
     m.act = m.qcod+STR(lpuid,4)+m.extip+m.videxp+m.podtip+m.actnum
     m.d_a = m.actdat
     
     INSERT INTO mefile FROM MEMVAR 
     
     m.et = m.et_old
    ENDIF 

   ENDSCAN 
   SET RELATION OFF INTO talon
   SELECT talon 
   SET RELATION OFF INTO people
   USE IN talon 
   USE IN people
   USE IN merror
   IF USED('mefile')
    USE IN mefile
   ENDIF 

   IF m.IsExpMee = .t.
    IF fso.FileExists(MEFilep+'.zip')
     fso.DeleteFile(MEFilep+'.zip')
    ENDIF 
 
    ZipOpen(MEFilep+'.zip')
    ZipFile(MEFilep+'.dbf')
    ZipClose()

    IF m.sexp2>0 OR m.ssank2>0
     m.tsexp = m.tsexp + m.sexp2
     m.tssank = m.tssank + m.ssank2
     INSERT INTO meexp (period, e_period, lpuid, mcod, et, sexp, stpn, s_sank) VALUES ;
      (m.gcperiod, m.period_e, m.lpuid, m.mcod, '2', m.sexp2, m.stpn2, m.ssank2)
    ENDIF 
    IF m.sexp3>0 OR m.ssank3>0
     m.tsexp = m.tsexp + m.sexp3
     m.tssank = m.tssank + m.ssank3
     INSERT INTO meexp (period, e_period, lpuid, mcod, et, sexp, stpn, s_sank) VALUES ;
      (m.gcperiod, m.period_e, m.lpuid, m.mcod, '3', m.sexp3, m.stpn3, m.ssank3)
    ENDIF 
    IF m.sexp4>0 OR m.ssank4>0
     m.tsexp = m.tsexp + m.sexp4
     m.tssank = m.tssank + m.ssank4
     INSERT INTO meexp (period, e_period, lpuid, mcod, et, sexp, stpn, s_sank) VALUES ;
      (m.gcperiod, m.period_e, m.lpuid, m.mcod, '4', m.sexp4, m.stpn4, m.ssank4)
    ENDIF 
    IF m.sexp5>0 OR m.ssank5>0
     m.tsexp = m.tsexp + m.sexp5
     m.tssank = m.tssank + m.ssank5
     INSERT INTO meexp (period, e_period, lpuid, mcod, et, sexp, stpn, s_sank) VALUES ;
      (m.gcperiod, m.period_e, m.lpuid, m.mcod, '5', m.sexp5, m.stpn5, m.ssank5)
    ENDIF 
    IF m.sexp6>0 OR m.ssank6>0
     m.tsexp = m.tsexp + m.sexp6
     m.tssank = m.tssank + m.ssank6
     INSERT INTO meexp (period, e_period, lpuid, mcod, et, sexp, stpn, s_sank) VALUES ;
      (m.gcperiod, m.period_e, m.lpuid, m.mcod, '6', m.sexp6, m.stpn6, m.ssank6)
    ENDIF 
    IF m.sexp7>0 OR m.ssank7>0
     m.tsexp = m.tsexp + m.sexp7
     m.tssank = m.tssank + m.ssank7
     INSERT INTO meexp (period, e_period, lpuid, mcod, et, sexp, stpn, s_sank) VALUES ;
      (m.gcperiod, m.period_e, m.lpuid, m.mcod, '7', m.sexp7, m.stpn7, m.ssank7)
    ENDIF 
    
   ENDIF 

   SELECT aisomse
   
  ENDSCAN 
  USE IN aisomse
  USE IN sookod
  USE IN sprlpu
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  USE IN tarif
  WAIT CLEAR 
  
 NEXT 

 SELECT meexp
 
* BROWSE 

 SET RELATION TO period+e_period+mcod+et+docexp INTO cursvacts 
 SCAN 
  m.et = et
  m.extip = IIF(INLIST(m.et,'2','3','7'),'1','2')
  m.videxp = IIF(INLIST(m.et,'2','4','6','7'),'1','2') && 1 - плановая, в т.ч. тематическая.
  m.podtip = IIF(INLIST(m.et,'2','3','7'),'0','1') && это неправильно! сделано временно!
  m.actnum = PADL(cursvacts.recid,6,'0')
  m.actdat = TTOD(cursvacts.actdate)

  m.n_akt = m.qcod+STR(lpuid,4)+m.extip+m.videxp+m.podtip+m.actnum
  
  REPLACE n_akt WITH m.actnum, d_akt WITH m.actdat
  
 ENDSCAN 
 SET RELATION OFF INTO cursvacts

 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "lpu_id") > 0
  USE IN meexp 
  RETURN
 ENDIF 

 PUBLIC oWord as Word.Application
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 
 DotName = pTempl+'\mefiles.dot'
 DocName = pOut+'\'+gcPeriod+'\mefiles'

 oDoc = oWord.Documents.Add(dotname)
 oTable = oDoc.Tables(1)
 
 SET RELATION TO lpuid INTO sprlpu
 npp = 1

 oDoc.Bookmarks('period').Select  
 oWord.Selection.TypeText(NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' года')

 SCAN
  SCATTER MEMVAR 
  m.mcod = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, mcod)
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.name, '')
  

  oTable.Cell(npp+1,1).Select
  oWord.Selection.TypeText(m.mcod)
  oTable.Cell(npp+1,2).Select
  oWord.Selection.TypeText(m.lpuname)
  oTable.Cell(npp+1,3).Select
  oWord.Selection.TypeText(m.e_period)
  oTable.Cell(npp+1,5).Select
  oWord.Selection.TypeText(TRANSFORM(m.sexp, '9999999.99'))
  oTable.Cell(npp+1,6).Select
  oWord.Selection.TypeText(TRANSFORM(m.stpn, '9999999.99'))
  oTable.Cell(npp+1,7).Select
  oWord.Selection.TypeText(TRANSFORM(m.s_sank, '9999999.99'))
  
  oWord.Selection.InsertRowsBelow
  npp = npp + 1
  
 ENDSCAN 
 SET RELATION OFF INTO sprlpu
 USE IN sprlpu
 SELECT meexp
 
 ppath = pout+'\'+gcperiod
 COPY TO &ppath\svmee
 USE
 
 IF OpenFile(ppath+'\svmee', 'svmee', 'excl')>0
  IF USED('svmee')
   USE IN svmee
  ENDIF 
 ENDIF 
 
 SELECT svmee

 INDEX on lpuid TAG lpuid

 USE 
 USE IN pilot
 USE IN mee2mgf
 USE IN cursvacts
 
 oTable.Cell(npp+1,5).Select
 oWord.Selection.TypeText(TRANSFORM(m.tsexp, '99999999.99'))

 oDoc.SaveAs(DocName, 0)
 oWord.Visible = .t.
 
RETURN