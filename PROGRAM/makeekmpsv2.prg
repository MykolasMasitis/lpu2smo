FUNCTION MakeEKMPSv2(parra1, parra2, parra3, parra4, parra5)

 m.lcpath      = parra1
 m.IsVisible   = parra2
 m.IsQuit      = parra3
 m.tipofexp    = parra4
 m.TipOfPeriod = parra5  && 0-ëîêàëüíûé ïåðèîä, 1 - ñâîäíûé!
 
 IF IsUsrDir=.T.
  m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
  IF !fso.FolderExists(m.usrdir)
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(m.usrdir))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SSACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SVACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ELSE 
  IF !fso.FolderExists(pmee)
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ENDIF 

 m.expname = 'Àêò '
 DO CASE 
  CASE m.tipofexp = '4'
   m.expname = m.expname + 'ïëàíîâîé'
  CASE m.tipofexp = '5'
   m.expname = m.expname + 'öåëåâîé'
  CASE m.tipofexp = '6'
   m.expname = m.expname + 'òåìàòè÷åñêîé'
 ENDCASE 
 m.expname = m.expname + ' ýêñïåðòèçû êà÷åñòâà ìåäèöèíñêîé ïîìîùè'
 m.flcod       = goApp.flcod

* IF  m.tipofexp='5'
*  DotName = 'ACT_EKMP_cel.xls'
* ELSE 
  DotName = 'ActEKMPsv.xls'
* ENDIF 

 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË ØÀÁËÎÍ ÎÒ×ÅÒÀ'+CHR(13)+CHR(10)+;
   'ActEKMPsv.xls',0+32,'')
  RETURN 
 ENDIF 
 
 m.mcod       = goApp.mcod 
 m.lpuid      = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.IsVed      = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.lpuaddress = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod+', '+m.lpuaddress, '')
 m.lpudog     = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
 m.lpuddog     = IIF(SEEK(m.lpuid, 'lpudogs'), lpudogs.ddogs, {})
 m.lpudog     = 'â ñîîòâåòñòâèè ñ Äîãîâîðîì '+m.lpudog+' îò '+DTOC(m.lpuddog)
 IF FIELD('SENT')='SENT'
  m.sent       = sent
 ELSE 
  m.sent       = DATETIME()
 ENDIF 
 m.dexp1 = DTOC(m.tdat1)
 m.dexp2 = DTOC(m.tdat2)

 IF m.TipOfPeriod=0
  pPath = pBase+'\'+gcPeriod+'\'+m.mcod
  TFile = 'talon'
  mFile = 'm'+m.mcod
 ELSE 
  m.flcod = aisoms.flcod
  pPath = pBase+'\'+gcPeriod+'\0000000\'+m.mcod
  TFile = 't'+m.flcod
  mFile = 'm'+m.flcod
 ENDIF 

 IF OpenFile(pPath+'\'+TFile, 'Talon', 'SHARED', 'recid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pPath+'\'+mFile, 'merror', 'SHARED')>0
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 

 SELECT merror
 
 COUNT FOR !EMPTY(err_mee) AND et=m.TipOfExp TO m.nIsEkmp
 
 IF m.nIsEkmp<=0
  MESSAGEBOX(CHR(13)+CHR(10)+'ÏÎ ÂÛÁÐÀÍÍÎÌÓ ËÏÓ ÝÊÌÏ ÍÅ ÏÐÎÂÎÄÈËÀÑÜ!'+CHR(13)+CHR(10),0+32,'')
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 SELECT docexp FROM merror WHERE et=m.TipOfExp GROUP BY docexp INTO CURSOR ;
  curexps

 SELECT curexps
 SCAN 
  m.docexp = docexp
  m.docfio = IIF(SEEK(m.docexp, 'explist'), ;
   ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot)+', êîä '+m.docexp, '')
  =OneSvAct(m.docexp)
 ENDSCAN 
 USE IN curexps
 
 IF USED('talon') 
  USE IN talon 
 ENDIF 
 IF USED('merror') 
  USE IN merror
 ENDIF 

 SELECT aisoms

RETURN 

FUNCTION OneSvAct(paraexp)
 PRIVATE m.docexp
 
 m.docexp = m.paraexp

 SELECT recid FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
  AND docexp=m.docexp AND flcod=m.flcod INTO CURSOR rqwest NOCONSOLE  
 m.nfileid = recid
 USE IN rqwest
 IF m.nfileid>0
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO svacts (period,mcod,codexp,docexp,flcod) ;
   VALUES ;
  (m.gcperiod,m.mcod,INT(VAL(m.tipofexp)), m.docexp, m.flcod)
  m.nfileid = GETAUTOINCVALUE()
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
  UPDATE svacts SET actname=PADL(m.nfileid,6,'0')+'.doc', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 IF fso.FileExists(DocName+'.xls')
  oFile = fso.GetFile(DocName+'.xls')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF MESSAGEBOX('ÏÎ ÂÛÁÐÀÍÍÎÌÓ ËÏÓ ÀÊÒ ÓÆÅ ÔÎÐÌÈÐÎÂÀËÑß!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ÄÀÒÀ ÑÎÇÄÀÍÈß ÀÊÒÀ            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ÄÀÒÀ ÏÎÑËÅÄÍÅÃÎ ÎÒÊÐÛÒÈß ÀÊÒÀ : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ÄÀÒÀ ÏÎÑËÅÄÍÅÃÎ ÈÇÌÅÍÅÍÈß ÀÊÒÀ: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ÂÛ ÕÎÒÈÒÅ ÏÅÐÅÔÎÐÌÈÐÎÂÀÒÜ ÀÊÒ?',4+32,'') == 7 
   
   SELECT curexps
   RETURN
  ELSE 
   IF m.TipOfPeriod=0
    UPDATE svacts SET actdate=DATETIME() WHERE recid = m.nfileid
   ENDIF 
  ENDIF 
 ENDIF 
  
 m.checked_tot = 0
 m.checked_amb = 0
 m.checked_dst = 0
 m.checked_st  = 0

 m.nepredst    = 0
 m.checked     = 0
 m.totdefs     = 0
 m.sumneoplata = 0
 m.tot_straf   = 0
 m.kol_straf   = 0
  
 DO CASE 
  CASE m.TipOfExp = '2'
   m.podvid='0'
  CASE m.TipOfExp = '3'
   m.podvid='1'
  CASE m.TipOfExp = '4'
   m.podvid='0'
  CASE m.TipOfExp = '5'
   m.podvid='1'
  CASE m.TipOfExp = '6'
   m.podvid='Ò'
  CASE m.TipOfExp = '7'
   m.podvid='Ò'
  OTHERWISE 
   m.podvid='0'
 ENDCASE 
 IF m.TipOfPeriod=0
*  m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ELSE 
*  m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'
  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ENDIF 

* IF m.TipOfPeriod=0
*  m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
* ELSE 
*  m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'
* ENDIF 

 m.d_akt = DTOC(DATE())
 
 m.nakt = '¹ ' + m.n_akt + ' îò ' + m.d_akt
 m.cpredps = 'Ïðåäïèñàíèå ¹ '+m.n_akt+' îò '+m.d_akt

 m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)

 CREATE CURSOR curpaz (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 CREATE CURSOR qwertamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertst (c_i c(30))
 INDEX on c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR qwertdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertbad (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR workcurs (nrec n(5), sn_pol c(25), c_i c(30), s_all n(11,2), er_c c(2), osn230 c(6), ;
  koeff n(4,2), straf n(4,2))
 INDEX ON sn_pol TAG sn_pol
 INDEX ON c_i TAG c_i
 
 CREATE CURSOR curdata2 (nrec2 n(5), sn_pol c(25), c_i c(30), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
  osn230 c(5), er_c c(3), delta n(11,2), koeff n(3,2), straf n(11,2), cmnt c(50))
 SELECT curdata2
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curdata3 (nrec3 n(5), sn_pol c(25), c_i c(30), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
  osn230 c(5), er_c c(3), delta n(11,2), koeff n(3,2), straf n(11,2), cmnt c(50))
 SELECT curdata3
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT merror
 SET RELATION TO recid INTO talon 
 m.nrec2 = 0 
 m.nrec3 = 0 
 SCAN 

  IF !(et=m.TipOfExp AND docexp=m.docexp)
   LOOP 
  ENDIF 
  
  m.s_1         = s_1
  m.s_2         = s_2
  m.koeff       = koeff
  m.straf       = 0
  m.sn_pol      = talon.sn_pol
  m.c_i         = talon.c_i
  m.er_c        = UPPER(LEFT(err_mee,2))
  m.osn230      = osn230
  m.s_all       = s_all
  m.cod         = cod

  IF m.er_c == 'GG' AND !SEEK(m.sn_pol, 'curdata3') && Íå ïðåäñòàâëåíî ïåðâè÷êè
   m.nepredst = m.nepredst + 1
   INSERT INTO curdata3 FROM MEMVAR 
   m.nrec3 = m.nrec3+1
  ENDIF 

  IF m.s_2>0 AND !SEEK(m.sn_pol, 'curdata2')
   m.straf       = straf
   m.tot_straf = m.tot_straf + (m.straf*m.ynorm)
   INSERT INTO curdata2 FROM MEMVAR 
   m.nrec2 = m.nrec2+1
  ENDIF 

  INSERT INTO workcurs (sn_pol, c_i, s_all, er_c, osn230, koeff, straf) VALUES ;
   (m.sn_pol, m.c_i, ROUND(m.s_all*m.koeff,2), m.er_c, m.osn230, m.koeff, m.straf)

*  IF !SEEK(m.sn_pol, 'curpaz')
*   INSERT INTO curpaz FROM MEMVAR 
*   m.tot_straf = m.tot_straf + (m.straf*m.ynorm)
   m.kol_straf = m.kol_straf + IIF(m.straf*m.ynorm>0,1,0)
   IF m.er_c == 'GG' && Íå ïðåäñòàâëåíî ïåðâè÷êè
*    m.nepredst = m.nepredst + 1
   ELSE 
*    m.checked = m.checked + 1

    IF IsUsl(m.cod) AND !SEEK(m.sn_pol, 'qwertamb')
     INSERT INTO qwertamb (sn_pol) VALUES (m.sn_pol)
     m.checked_amb = m.checked_amb + 1
    ENDIF 
   
    IF IsKD(m.cod) AND !SEEK(m.sn_pol, 'qwertdst')
     INSERT INTO qwertdst (sn_pol) VALUES (m.sn_pol)
     m.checked_dst = m.checked_dst + 1
    ENDIF 
   
    IF (IsMes(m.cod) OR IsVMP(m.cod)) AND !SEEK(m.c_i, 'qwertst')
     INSERT INTO qwertst (c_i) VALUES (m.c_i)
     m.checked_st = m.checked_st + 1
    ENDIF 

    IF (Is02(m.cod)) AND !SEEK(m.c_i, 'qwertst')
     INSERT INTO qwertst (c_i) VALUES (m.c_i)
     m.checked_st = m.checked_st + 1
    ENDIF 

   ENDIF 
   IF m.er_c != 'W0'
    m.totdefs = m.totdefs + 1
   ENDIF 
*  ENDIF 

 ENDSCAN 
 SET RELATION OFF INTO talon 
 
 m.defs = ''
 SELECT workcurs
 SET ORDER TO sn_pol
 GO TOP 
 m.nrec = IIF(m.mcod='0371001', 0, 1)
 m.polis = sn_pol
 SCAN
  m.sumneoplata = m.sumneoplata + s_all
  m.sn_pol = sn_pol
  m.defs = m.defs+osn230+'; '
  IF m.sn_pol!=m.polis OR m.mcod='0371001'
   m.polis = m.sn_pol
   m.nrec = m.nrec + 1
  ENDIF 
  REPLACE nrec WITH m.nrec
 ENDSCAN 
 
 SELECT curdata2
 IF RECCOUNT('curdata2')<=0
  SCATTER MEMVAR 
  INSERT INTO curdata2 FROM MEMVAR 
 ENDIF 
 SELECT curdata3
 IF RECCOUNT('curdata3')<=0
  SCATTER MEMVAR 
  INSERT INTO curdata3 FROM MEMVAR 
 ENDIF 
* SET ORDER TO 
* REPLACE ALL nrec2 WITH RECNO()

 m.checked = m.checked_amb + m.checked_dst + m.checked_st

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'

 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)

 USE IN workcurs
 USE IN qwertamb
 SELECT qwertst
 USE IN qwertdst

 SELECT curexps

RETURN 