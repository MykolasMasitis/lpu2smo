FUNCTION MakeEKMPSv(parra1, parra2, parra3, parra4, parra5)

 m.lcpath      = parra1
 m.IsVisible   = parra2
 m.IsQuit      = parra3
 m.tipofexp    = parra4
 m.TipOfPeriod = parra5  && 0-ëîêàëüíûé ïåðèîä, 1 - ñâîäíûé!
 
 m.expname = ''
 DO CASE 
  CASE m.tipofexp = '4'
   m.expname = 'ïëàíîâîé'
  CASE m.tipofexp = '5'
   m.expname = 'öåëåâîé'
  CASE m.tipofexp = '6'
   m.expname = 'òåìàòè÷åñêîé'
 ENDCASE 

 DotName = 'Àêò_ÝÊÌÏ_ñâîä.dot'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË ØÀÁËÎÍ ÎÒ×ÅÒÀ'+CHR(13)+CHR(10)+;
   'Àêò_ÝÊÌÏ_ñâîä.dot',0+32,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 m.mcod       = mcod 
 m.lpuid      = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.IsVed      = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.lpuaddress = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 m.lpuname    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name)+', '+m.mcod+', '+m.lpuaddress, '')
 IF FIELD('SENT')='SENT'
  m.sent       = sent
 ELSE 
  m.sent       = DATETIME()
 ENDIF 

 IF m.TipOfPeriod=0
  pPath = pBase+'\'+gcPeriod+'\'+m.mcod
  TFile = 'talon'
  mFile = 'm'+m.mcod
 ELSE 
  pPath = pBase+'\'+gcPeriod+'\0000000\'+m.mcod
  TFile = 't'+flcod
  mFile = 'm'+flcod
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

 WAIT "ÇÀÏÓÑÊ WORD..." WINDOW NOWAIT 
 TRY 
  oWord = GETOBJECT(,"Word.Application")
 CATCH 
  oWord = CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 
 
 SELECT curexps
 SCAN 
  m.docexp = docexp
  m.docfio = IIF(SEEK(m.docexp, 'explist'), ;
   ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot)+', êîä '+m.docexp, '')
  =OneSvAct(m.docexp)
 ENDSCAN 
 
 IF USED('talon') 
  USE IN talon 
 ENDIF 
 IF USED('merror') 
  USE IN merror
 ENDIF 
 
 SELECT aisoms

 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

RETURN 

FUNCTION OneSvAct(paraexp)
 PRIVATE m.docexp
 
 m.docexp = m.paraexp

 IF m.TipOfPeriod=0

 SELECT recid FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
  AND docexp=m.docexp INTO CURSOR rqwest NOCONSOLE  
 m.nfileid = recid
 USE IN rqwest

 IF m.nfileid>0
  DocName = pmee+'\svacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO svacts (period,mcod,codexp,docexp) ;
   VALUES ;
  (m.gcperiod,m.mcod,INT(VAL(m.tipofexp)), m.docexp)
  m.nfileid = GETAUTOINCVALUE()
  DocName = pmee+'\svacts\'+PADL(m.nfileid,6,'0')
  UPDATE svacts SET actname=PADL(m.nfileid,6,'0')+'.doc', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
 
 ELSE 

 SELECT aisoms
 DocName = pBase+'\'+gcPeriod+'\0000000\'+m.mcod+'\ekmp'+TipOfExp+'sv'+flcod
 
 ENDIF 
 
 IF fso.FileExists(DocName+'.doc')
  oFile = fso.GetFile(DocName+'.doc')
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
 
 m.IsExpMee    = .f.
 m.checked_tot = 0
 m.bad_kol     = 0
 m.bad_sum     = 0
 m.opl_tot     = 0
 m.vzaim_tot   = 0
 nRow          = 3
 m.dschet      = TTOC(m.sent)+', íîìåð ñ÷åòà '+STR(tYear,4)+PADL(tMonth,2,'0')
 m.nepredst    = 0
 m.checked     = 0
 m.totdefs     = 0
 m.sumneoplata = 0
 m.kol_strafs  = 0
 m.sum_strafs  = 0
 m.fioexp      = m.usrfam+' '+m.usrim+' '+m.usrot
 m.tot_straf   = 0
 m.kol_straf   = 0
 m.tot_sall    = 0
 m.n_akt       = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 m.d_akt       = DTOC(DATE())
  
 CREATE CURSOR workcurs (sn_pol c(25), c_i c(30), s_all n(11,2), er_c c(2), osn230 c(6), ;
  koeff n(4,2), straf n(4,2))
 INDEX ON sn_pol TAG sn_pol
 INDEX ON c_i TAG c_i
 
 SELECT merror
 SET RELATION TO recid INTO talon 
 SCAN 

  IF !(et=m.TipOfExp AND docexp=m.docexp)
   LOOP 
  ENDIF 
  
  m.sn_pol      = talon.sn_pol
  m.c_i         = talon.c_i
  m.er_c        = UPPER(LEFT(err_mee,2))
  m.koeff       = koeff
  m.straf       = straf
  m.sumneoplata = m.sumneoplata + ROUND(s_all*m.koeff,2)
  m.s_all       = s_all

  IF !SEEK(m.sn_pol, 'workcurs', 'sn_pol')
   INSERT INTO workcurs (sn_pol, c_i, s_all, er_c, koeff, straf) VALUES ;
    (m.sn_pol, m.c_i, ROUND(m.s_all*m.koeff,2), m.er_c, m.koeff, m.straf)
   IF m.er_c == 'GG' && Íåïðåäñòàâëåíî ïåðâè÷êè
    m.nepredst = m.nepredst + 1
   ELSE 
    m.checked = m.checked + 1
   ENDIF 
   IF m.er_c != 'W0'
    m.totdefs = m.totdefs + 1
   ENDIF 
  ELSE 
   IF !EMPTY(Talon.Tip)
    IF !SEEK(m.c_i, 'workcurs', 'c_i')
     INSERT INTO workcurs (sn_pol, c_i, s_all, er_c, koeff, straf) VALUES ;
      (m.sn_pol, m.c_i, ROUND(m.s_all*m.koeff,2), m.er_c, m.koeff, m.straf)
     m.checked = m.checked + 1
     IF m.er_c != 'W0'
      m.totdefs = m.totdefs + 1
     ENDIF 
    ELSE 
     IF workcurs.sn_pol = m.sn_pol
      m.new_sum = workcurs.s_all + ROUND(s_all*m.koeff,2)
      REPLACE workcurs.s_all WITH m.new_sum IN workcurs
      IF workcurs.er_c=='W0' AND m.er_c!='W0'
       REPLACE workcurs.er_c WITH m.er_c IN workcurs
      ENDIF 
     ELSE 
      SELECT workcurs
      SET ORDER TO c_i
      m.lIsFound = .F.
      DO WHILE c_i=m.c_i
       SKIP 
       IF sn_pol=m.sn_pol
        m.lIsFound = .T.
        EXIT 
       ENDIF 
      ENDDO 
      IF m.lIsFound = .T.
       m.new_sum = workcurs.s_all + ROUND(s_all*m.koeff,2)
       REPLACE s_all WITH m.new_sum
       IF er_c=='W0' AND m.er_c!='W0'
        REPLACE er_c WITH m.er_c
       ENDIF 
      ELSE 
       MESSAGEBOX('ÍÅ ÓÄÀËÎÑÜ ÊÎÐÐÅÊÒÍÎ ÏÐÎÑÒÀÂÈÒÜ ÑÓÌÌÓ!',0+48,m.sn_pol)
      ENDIF 
      SET ORDER TO 
      SELECT merror
     ENDIF 
    ENDIF
   ELSE 
    m.new_sum = workcurs.s_all + ROUND(s_all*m.koeff,2)
    REPLACE workcurs.s_all WITH m.new_sum IN workcurs
    IF workcurs.er_c=='W0' AND m.er_c!='W0'
     REPLACE workcurs.er_c WITH m.er_c
    ENDIF 
   ENDIF 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO talon 
 
 SELECT workcurs
 SET ORDER TO sn_pol

 oDoc = oWord.Documents.Add(pTempl+'\'+DotName)

 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)
* oDoc.Bookmarks('docname').Select  
* oWord.Selection.TypeText(m.docfio)
 oDoc.Bookmarks('d_exp').Select  
 oWord.Selection.TypeText(m.d_akt)
 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('d_exp1').Select  
 oWord.Selection.TypeText(DTOC(m.tdat1))
 oDoc.Bookmarks('d_exp2').Select  
 oWord.Selection.TypeText(DTOC(m.tdat2))
 
 m.sn_pol = sn_pol
 SCAN 
  m.er_c = er_c
  m.osn230 = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), ;
    sookod.osn230, IIF(LEFT(UPPER(m.er_c),2)!='W0', '!', ''))	
  m.koeff = koeff
  m.straf = straf 

  oDoc.Tables(1).Cell(nRow,1).Select 
  oWord.Selection.InsertRows
  oDoc.Tables(1).Cell(nRow,2).Select 
  oWord.Selection.TypeText(ALLTRIM(sn_pol))
  oDoc.Tables(1).Cell(nRow,3).Select && Êàðòà
  oWord.Selection.TypeText(ALLTRIM(c_i))
  oDoc.Tables(1).Cell(nRow,4).Select && Êîä äåôåêòà ïî Ïð.230
  oWord.Selection.TypeText(m.osn230)
  oDoc.Tables(1).Cell(nRow,5).Select && Ïðîöåíò ñíÿòèé
  oWord.Selection.TypeText(TRANSFORM(m.koeff*100,'999'))
  oDoc.Tables(1).Cell(nRow,6).Select && Îïëà÷åíî çà óñëóãè
  oWord.Selection.TypeText(TRANSFORM(s_all, '9999999.99'))
  oDoc.Tables(1).Cell(nRow,7).Select && Øòðàô
  oWord.Selection.TypeText(TRANSFORM(m.straf*m.ynorm, '9999999.99'))
  nRow = nRow + 1

  m.tot_straf = m.tot_straf + (m.straf*m.ynorm)
  m.kol_straf = m.kol_straf + IIF(m.straf*m.ynorm>0,1,0)
  m.tot_sall = m.tot_sall + s_all

 ENDSCAN 
 
 USE 

 oDoc.Tables(1).Cell(nRow,6).Select && Îïëà÷åíî çà óñëóãè
 oWord.Selection.TypeText(TRANSFORM(m.tot_sall, '9999999.99'))
 oDoc.Tables(1).Cell(nRow,7).Select && Øòðàô
 oWord.Selection.TypeText(TRANSFORM(m.tot_straf, '9999999.99'))

 oDoc.Bookmarks('expname').Select  
 oWord.Selection.TypeText(m.expname)
 oDoc.Bookmarks('neprdest').Select  
 oWord.Selection.TypeText(TRANSFORM(m.nepredst, '999'))
 oDoc.Bookmarks('totchecked').Select  
 oWord.Selection.TypeText(TRANSFORM(m.checked, '999'))
 oDoc.Bookmarks('totdefs').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totdefs, '999'))
 oDoc.Bookmarks('sumneoplata').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumneoplata, '999999.99'))
 oDoc.Bookmarks('kol_strafs').Select  
 oWord.Selection.TypeText(TRANSFORM(m.kol_straf, '9999'))
 oDoc.Bookmarks('sum_strafs').Select  
 oWord.Selection.TypeText(TRANSFORM(m.tot_straf, '9999999.99'))
 
 oDoc.SaveAs(DocName,0)
 IF IsVisible == .F.
  oDoc.Close(0)
 ENDIF 

 SELECT curexps
 RETURN

RETURN 