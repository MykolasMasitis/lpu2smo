PROCEDURE ExpDSoft
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÝÊÑÏÎÐÒÈÐÎÂÀÒÜ'+CHR(13)+CHR(10)+'ÄÀÍÍÛÅ Â ÄÈÀÑÎÔÒ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 pDSoft = fso.GetParentFolderName(pbin)+'\DSoft'
 IF !fso.FolderExists(pDSoft)
  fso.CreateFolder(pDSoft)
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\osoerzxx', 'OsoERZ', 'Shar', 'ans_r')>0
  IF USED('osoerz')
   USE IN osoerz
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'Shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('osoerz')
   USE IN osoerz
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\pilot', 'pilot', 'Shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('osoerz')
   USE IN osoerz
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 SELECT aisoms
 SCAN 
  IF s_pred<=0
   LOOP 
  ENDIF 
  IF s_pred-sum_flk<=0
   LOOP 
  ENDIF 
  m.mcod = mcod 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  m.mmy = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\b'+m.mcod+'.'+m.mmy)
   LOOP 
  ENDIF 
  m.mcod = mcod 
  m.tpn_mo = IIF(SEEK(m.mcod, 'pilot'), '1', '')
  m.lpuid = lpuid
  IF m.lpuid<=0
   LOOP 
  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 
  IF !fso.FolderExists(pDSoft+'\'+STR(m.lpuid,4))
   fso.CreateFolder(pDSoft+'\'+STR(m.lpuid,4))
  ENDIF 

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\b'+m.mcod+'.'+m.mmy)
   fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\b'+m.mcod+'.'+m.mmy)
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\b'+m.mcod+'.'+m.mmy)
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\b'+m.mcod+'.'+m.mmy, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\b'+m.mcod+'.'+m.mmy, .t.)
  ENDIF 

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\ctrl.dbf')
   fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\ctrl.dbf')
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ctrl'+m.qcod+'.dbf')
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ctrl'+m.qcod+'.dbf', ;
    pDSoft+'\'+STR(m.lpuid,4)+'\ctrl.dbf', .t.)
  ENDIF 

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\erz_ans.dbf')
   fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\erz_ans.dbf')
  ENDIF 
  IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\Answer.dbf')
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\Answer.dbf', ;
    pDSoft+'\'+STR(m.lpuid,4)+'\erz_ans.dbf', .t.)
    IF OpenFile(pDSoft+'\'+STR(m.lpuid,4)+'\erz_ans', 'ans', 'excl')<=0
     SELECT ans
     INDEX on ALLTRIM(recid) TAG rr
     SET ORDER TO 
    ENDIF 
    IF USED('ans')
     USE IN ans
    ENDIF 
  ENDIF 

  ZipName = pDSoft+'\'+STR(m.lpuid,4)+'\b'+m.mcod+'.'+m.mmy
  ZipDir  = pDSoft+'\'+STR(m.lpuid,4)+'\'

  hItem    = 'H'  + STR(m.lpuid,4) + '.' + m.mmy
  dItem    = 'D'  + STR(m.lpuid,4) + '.' + m.mmy
  nvItem   = 'NV' + STR(m.lpuid,4) + '.' + m.mmy
  nsItem   = 'NS' + STR(m.lpuid,4) + '.' + m.mmy
  rItem    = 'R' + m.qcod + '.' + m.mmy
  sItem    = 'S' + m.qcod + '.' + m.mmy
  hoItem   = 'HO' + m.qcod + '.' + m.mmy
  dsItem   = 'D79S' + m.qcod + '.' + m.mmy
  sprItem  = 'SPR' + STR(m.lpuid,4) + '.' + m.mmy
  hosItem  = 'hos6' + STR(m.lpuid,4) + '.' + m.mmy

  dItem2    = 'D'  + STR(m.lpuid,4) + '.dbf'
  hoItem2   = 'HO' + m.qcod + '.dbf'
  hItem2    = 'H'  + STR(m.lpuid,4) + '.dbf'
  nvItem2   = 'NV' + STR(m.lpuid,4) + '.dbf'
  sprItem2  = 'SPR' + STR(m.lpuid,4) + '.dbf'
  rItem2    = 'R' + m.qcod + '.dbf'
  sItem2    = 'S' + m.qcod + '.dbf'

  UnzipOpen(ZipName)

  UnzipGotoFileByName(dItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.dItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.dItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.dItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.dItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.dItem)    
   ENDIF 
  ENDIF 
  UnzipGotoFileByName(hoItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hoItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hoItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hoItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hoItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hoItem)    
   ENDIF 
  ENDIF 
  UnzipGotoFileByName(hItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.hItem)    
   ENDIF 
  ENDIF 
  UnzipGotoFileByName(nvItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.nvItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.nvItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.nvItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.nvItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.nvItem)    
   ENDIF 
  ENDIF 
  UnzipGotoFileByName(rItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem)    
   ENDIF 
  ENDIF 
  UnzipGotoFileByName(sItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem)    
   ENDIF 
  ENDIF 
  UnzipGotoFileByName(sprItem)
  UnzipFile(ZipDir)
  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sprItem)
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sprItem, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sprItem2, .t.)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sprItem2)
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sprItem)    
   ENDIF 
  ENDIF 
  IF UnzipGotoFileByName(hosItem)
   UnzipFile(ZipDir)
  ENDIF 

  UnzipClose()

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem2)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.dbf')
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.dbf')
   ENDIF 
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.rItem2, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\reestr.dbf', .t.)
  ENDIF 

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.dbf')
   oSettings.CodePage(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.dbf', 866, .t.)
   IF OpenFile(pDSoft+'\'+STR(m.lpuid,4)+'\reestr', 'people', 'excl')<=0
    SELECT people
    DELETE TAG ALL 
    INDEX on ALLTRIM(sn_pol) TAG sn_pol
    SET ORDER TO 
    ALTER table people ADD COLUMN file c(20)
    ALTER table people ADD COLUMN c_err c(3)
    ALTER table people ADD COLUMN mcod c(7)
    REPLACE ALL mcod WITH m.mcod
    ALTER table people ADD COLUMN ans_r c(3)
    ALTER table people add COLUMN q c(2)
    ALTER table people add COLUMN sn_poly c(25)
    ALTER table people add COLUMN ans_tip c(1)
    ALTER table people add COLUMN id_erz c(6)
    ALTER table people ADD COLUMN lpu_id_p n(6)
    ALTER table people ADD COLUMN lpu_p09 c(1)
    
    IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers.dbf') AND ;
     fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
     IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answers', 'answer', 'excl')<=0 AND ;
      OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'rerr', 'shar', 'rrid')<=0
      SELECT answer
      m.ansrec = RECCOUNT('answer')
      m.pplrec = RECCOUNT('people')
      IF m.ansrec=m.pplrec
       INDEX on RECNO() TAG recno
       SET ORDER to recno 
       SELECT people
       SET RELATION TO RECNO() INTO answer
       SET RELATION TO RECNO() INTO rerr ADDITIVE 
       SCAN 
        REPLACE ans_r WITH answer.ans_r, q WITH answer.q,;
         sn_poly WITH answer.n_pol, ;
         ans_tip WITH IIF(SEEK(people.ans_r, 'osoerz') AND osoerz.kl == 'y', 'y', 'n'),;
         id_erz WITH answer.recid, lpu_id_p WITH answer.lpu_id, lpu_p09 WITH IIF(answer.lpu_id>0,'1',''),;
         c_err WITH rerr.c_err, file WITH UPPER(m.sItem)
       ENDSCAN 
       SET RELATION OFF INTO answer
       SET RELATION OFF INTO rerr
       SELECT answer
       SET ORDER TO 
       DELETE TAG ALL 
      ENDIF 
     ENDIF 
     IF USED('answer')
      USE IN answer
     ENDIF 
     IF USED('rerr')
      USE IN rerr
     ENDIF 
    ENDIF 
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 

   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.bak')
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.bak')
   ENDIF 

  ENDIF 

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem2)
   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\schet.dbf')
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\schet.dbf')
   ENDIF 
   fso.CopyFile(pDSoft+'\'+STR(m.lpuid,4)+'\'+m.sItem2, ;
    pDSoft+'\'+STR(m.lpuid,4)+'\schet.dbf', .t.)
  ENDIF 

  IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\schet.dbf')
   oSettings.CodePage(pDSoft+'\'+STR(m.lpuid,4)+'\schet.dbf', 866, .t.)
   IF OpenFile(pDSoft+'\'+STR(m.lpuid,4)+'\schet', 'schet', 'excl')<=0
    SELECT schet 
    ALTER TABLE schet ADD COLUMN file c(20)
    ALTER TABLE schet ADD COLUMN s_all n(14,2)
    ALTER TABLE schet ADD COLUMN c_err c(3)
    ALTER TABLE schet ADD COLUMN mcod c(7)
    ALTER TABLE schet ADD COLUMN profil c(3)
    ALTER TABLE schet ADD COLUMN tarif n(10,2)
    ALTER TABLE schet ADD COLUMN tarif1 n(10,2)
    ALTER TABLE schet ADD COLUMN tarif2 n(10,2)
    ALTER TABLE schet ADD COLUMN tarif3 n(10,2)
    ALTER TABLE schet ADD COLUMN s_all1 n(14,2)
    ALTER TABLE schet ADD COLUMN s_all2 n(14,2)
    ALTER TABLE schet ADD COLUMN s_all3 n(14,2)
    ALTER TABLE schet ADD COLUMN ump n(3)
    ALTER TABLE schet ADD COLUMN ump_pf n(3)
    ALTER TABLE schet ADD COLUMN vmp n(1)
    ALTER TABLE schet ADD COLUMN d_norm n(4)
    ALTER TABLE schet ADD COLUMN kd_opl n(4)
    ALTER TABLE schet ADD COLUMN pr_usl c(2)
    ALTER TABLE schet ADD COLUMN date_in d
    ALTER TABLE schet ADD COLUMN pac_tpn c(1)
    ALTER TABLE schet ADD COLUMN tpn c(1)
    ALTER TABLE schet ADD COLUMN tpn_mo c(1)
    
    IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf') AND ;
     fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
     IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'serr', 'shar', 'rid')<=0 AND ;
      OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'excl')<=0 AND ;
      OpenFile(pDSoft+'\'+STR(m.lpuid,4)+'\reestr.dbf', 'reestr', 'shar', 'sn_pol')<=0 AND ;
      OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')<=0
      
      SELECT talon
      INDEX on RECNO() TAG recno
      SET ORDER to recno 
      SELECT schet
      SET RELATION TO RECNO() INTO talon
      SET RELATION TO RECNO() INTO serr ADDITIVE 
      SET RELATION TO cod INTO tarif ADDITIVE 
      SET RELATION TO ALLTRIM(sn_pol) INTO reestr ADDITIVE 
      SET RELATION TO sn_pol INTO people ADDITIVE 
      SCAN 
       m.cod = cod
       m.pr_usl = ''
       DO CASE 
        CASE INLIST(m.cod,101927,101928)
         m.pr_usl = '02'
        CASE INLIST(m.cod,101929,101930,101931,101932)
         m.pr_usl = '04'
        CASE INLIST(m.cod,1900,1901,1902,1903,1904,1905)
         m.pr_usl = '05'
        CASE INLIST(m.cod,15001,115001)
         m.pr_usl = '10'
        OTHERWISE 
       ENDCASE 
       m.kd_opl = IIF(IsMes(m.cod), k_u, 0)
       DO CASE 
        CASE IsMes(m.cod) OR IsVmp(m.cod) OR IsKdS(m.cod) OR IsPat(m.cod)
         m.ump=1
        CASE IsKdP(m.cod)
         m.ump=2
        OTHERWISE 
         m.ump=3
       ENDCASE 
       
       m.pac_tpn = IIF(reestr.lpu_id_p=m.lpuid,'1','')
       REPLACE c_err WITH serr.c_err, file WITH UPPER(m.sItem), s_all WITH talon.s_all, mcod WITH m.mcod,;
        s_all1 WITH talon.s_all, profil WITH talon.profil, tarif WITH tarif.tarif, tarif1 WITH  tarif.tarif,;
        tpn WITH tarif.tpn, vmp WITH tarif.vmp, tpn_mo WITH m.tpn_mo, pac_tpn WITH m.pac_tpn, ;
        date_in WITH people.d_beg, pr_usl WITH m.pr_usl, d_norm WITH tarif.n_kd, kd_opl WITH m.kd_opl,;
        ump WITH m.ump
        
      ENDSCAN 
      
      SET RELATION OFF INTO talon
      SET RELATION OFF INTO serr
      SET RELATION OFF INTO tarif
      SET RELATION OFF INTO reestr 
      SET RELATION OFF INTO people 
      SELECT talon
      SET ORDER TO 
      DELETE TAG recno
     ENDIF 
    ENDIF 
    IF USED('serr')
     USE IN serr
    ENDIF 
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   IF USED('reestr')
    USE IN reestr
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 

   IF fso.FileExists(pDSoft+'\'+STR(m.lpuid,4)+'\schet.bak')
    fso.DeleteFile(pDSoft+'\'+STR(m.lpuid,4)+'\schet.bak')
   ENDIF 

  ENDIF 
  IF USED('schet')
   USE IN schet
  ENDIF 

  WAIT CLEAR 
 ENDSCAN 
 USE IN aisoms
 USE IN osoerz
 USE IN tarif
 USE IN pilot

 MESSAGEBOX('ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!',0+64,'')

RETURN 