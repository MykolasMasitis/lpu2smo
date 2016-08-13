FUNCTION MakeMEESv2(lcPath, para3, IsQuit, para4, para5)
 
 m.IsVisible = para3
 m.tipofexp  = para4
 m.TipOfPeriod = para5 && 0-локальный период, 1 - сводный!
 
 IF IsUsrDir=.T.
  m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
  IF !fso.FolderExists(m.usrdir)
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SSACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
  IF !fso.FolderExists(m.usrdir+'\SVACTS')
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(m.usrdir+'\SSACTS'))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ELSE 
  IF !fso.FolderExists(pmee)
   MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ENDIF 

 DotName = 'ActMEEsv.xls'
 IF !fso.FileExists(pTempl+'\'+DotName)
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ ШАБЛОН ОТЧЕТА'+CHR(13)+CHR(10)+;
   'ActMEEsv.xls',0+32,'')
  RETURN 
 ENDIF 
 
 m.mcod  = SUBSTR(lcpath,RAT('\',lcpath)+1)
 IF m.TipOfPeriod=0
  m.flcod = ''
  pPath = pBase+'\'+gcPeriod+'\'+m.mcod
  TFile = 'talon'
  mFile = 'm'+m.mcod
 ELSE 
  m.flcod = flcod
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
 
 oal = ALIAS()
 SELECT merror
 m.nexps=0
 COUNT FOR !EMPTY(err_mee) AND et=m.tipofexp TO m.nexps
 IF m.nexps=0
  MESSAGEBOX(CHR(13)+CHR(10)+'ПО ВЫБРАННОМУ ЛПУ ЭКСПЕРТИЗА'+CHR(13)+CHR(10)+;
   'ЗАДАННОГО ТИПА НЕ ПРОВОДИЛАСЬ!',0+64,'')
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT &oal
  RETURN 
 ENDIF 

 SELECT docexp FROM merror WHERE et=m.TipOfExp GROUP BY docexp INTO CURSOR ;
  curexps
 SELECT curexps
 SCAN 
  m.docexp = docexp
  m.docfio = IIF(SEEK(m.docexp, 'explist'), ;
   ALLTRIM(explist.fam)+' '+ALLTRIM(explist.im)+' '+ALLTRIM(explist.ot)+', код '+m.docexp, '')
  =OneSvAct(m.docexp)
 ENDSCAN 
 USE IN curexps
 
 IF USED('talon') 
  USE IN talon 
 ENDIF 
 IF USED('merror') 
  USE IN merror
 ENDIF 

 SELECT &oal

RETURN 

FUNCTION OneSvAct(m.paraexp)

 PRIVATE m.docexp
 
 m.docexp = m.paraexp

 IF m.TipOfPeriod=0 && локальный период
  m.exp_dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
  m.exp_dat2 = DTOC(GOMONTH(CTOD(m.exp_dat1),1)-1)
 ELSE  && произвольный период
  m.exp_dat1 = DTOC(flmindate(m.flcod))
  m.exp_dat2 = DTOC(flmaxdate(m.flcod))
 ENDIF 

 IF !EMPTY(goApp.d_exp)
  m.edat1    = DTOC(goApp.d_exp)
  m.edat2    = m.edat1  
 ELSE 
  m.edat1    = DTOC(DATE())
  m.edat2    = m.edat1  
 ENDIF 

 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
 m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname)+', '+m.mcod, '')
 
 m.fioexp  = ''
 IF !EMPTY(goApp.smoexp)
  IF USED('users')
   IF SEEK(ALLTRIM(goApp.smoexp), 'users', 'name')
     m.fioexp  = ALLTRIM(users.fam)+' '+ALLTRIM(users.im)+' '+ALLTRIM(users.ot)
   ENDIF 
  ENDIF 
 ELSE 
 m.fioexp  = m.usrfam+' '+m.usrim+' '+m.usrot
 ENDIF 

 m.lpudog  = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
 m.lpuddog = IIF(SEEK(m.lpuid, 'lpudogs'), lpudogs.ddogs, {})
 m.lpucdog = m.lpudog+' от '+DTOC(m.lpuddog)

 ooal = ALIAS()

 SELECT recid FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.TipOfExp)) ;
   AND docexp=m.docexp AND flcod=m.flcod INTO CURSOR rqwest NOCONSOLE  
 m.nfileid = recid
 USE 
 SELECT (ooal)
 IF m.nfileid>0
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
 ELSE 
  INSERT INTO svacts (period,mcod,codexp,docexp,flcod) ;
   VALUES ;
  (m.gcperiod,m.mcod,INT(VAL(m.tipofexp)),m.docexp,m.flcod)
  m.nfileid = GETAUTOINCVALUE()
  DocName = IIF(!IsUsrDir, m.pmee, m.usrdir)+'\svacts\'+PADL(m.nfileid,6,'0')
  UPDATE svacts SET actname=PADL(m.nfileid,6,'0')+'.xls', actdate=DATETIME() WHERE recid = m.nfileid
 ENDIF 
  
 IF fso.FileExists(DocName+'.xls')
  oFile = fso.GetFile(DocName+'.xls')
  DateCreated      = TTOC(oFile.DateCreated)
  DateLastAccessed = TTOC(oFile.DateLastAccessed)
  DateLastModified = TTOC(oFile.DateLastModified)
  RELEASE oFile
  
  IF MESSAGEBOX('ПО ВЫБРАННОМУ ЛПУ АКТ УЖЕ ФОРМИРОВАЛСЯ!'+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА СОЗДАНИЯ АКТА            : '+m.DateCreated+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ОТКРЫТИЯ АКТА : '+m.DateLastAccessed+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ДАТА ПОСЛЕДНЕГО ИЗМЕНЕНИЯ АКТА: '+m.DateLastModified+CHR(13)+CHR(10)+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ ПЕРЕФОРМИРОВАТЬ АКТ?',4+32,'') == 7 
   
   USE IN talon 
   USE IN merror
   SELECT aisoms
   RETURN
  ELSE
   IF m.TipOfPeriod=0
    UPDATE svacts SET actdate=DATETIME() WHERE recid = m.nfileid
   ENDIF 
  ENDIF 
 ENDIF 

 m.IsExpMee = .f.

 m.checked_tot = 0
 m.checked_amb = 0
 m.checked_dst = 0
 m.checked_st  = 0

 m.bad_kol   = 0
 m.bad_sum   = 0
 m.opl_tot   = 0
 m.vzaim_tot = 0
 m.tot_straf = 0

 m.aktname=''
 DO CASE 
  CASE m.TipOfExp = '2'
   m.podvid='0'
   m.aktname='Акт плановой медико-экономичекой экспертизы №'
  CASE m.TipOfExp = '3'
   m.podvid='1'
   m.aktname='Акт целевой медико-экономичекой экспертизы №'
  CASE m.TipOfExp = '4'
   m.podvid='0'
   m.aktname='Сводный акт плановой ЭКМП №'
  CASE m.TipOfExp = '5'
   m.podvid='1'
   m.aktname='Сводный акт целевой ЭКМП №'
  CASE m.TipOfExp = '6'
   m.podvid='Т'
   m.aktname='Сводный акт тематической ЭКМП №'
  CASE m.TipOfExp = '7'
   m.podvid='Т'
   m.aktname='Акт тематической медико-экономичекой экспертизы №'
  CASE m.TipOfExp = '8'
   m.podvid='Т'
   m.aktname='Акт медико-экономичекой экспертизы по жалобе №'
  CASE m.TipOfExp = '9'
   m.podvid='Т'
   m.aktname='Сводный акт ЭКМП по жалобе №'
  OTHERWISE 
   m.podvid='0'
 ENDCASE 
 IF m.TipOfPeriod=0
*  m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'+ALLTRIM(STR(m.nfileid))
  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','8'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ELSE 
*  m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)+'/'
  m.n_akt = m.qcod+STR(m.lpuid,4)+IIF(INLIST(m.TipOfExp,'2','3','7'),'1','2')+;
   IIF(INLIST(m.TipOfExp,'2','4','6','7','8'),'1','2')+m.podvid+ALLTRIM(STR(m.nfileid))
 ENDIF 
 
 m.IsMee  = IIF(INLIST(m.TipOfExp,'2','3','7','8'),.T.,.F.)
 m.IsEkmp = !m.IsMee

 m.d_akt = DTOC(DATE())

 m.dschet    = IIF(FIELD('processed', 'aisoms')='PROCESSED', TTOC(aisoms.processed), '')+', номер счета '+STR(tYear,4)+PADL(tMonth,2,'0')
  
 m.cpredps = 'Предписание № '+m.n_akt+' от '+m.d_akt

 CREATE CURSOR qwert (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR qwertamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertst (c_i c(25))
 INDEX on c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR qwertdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR qwertbad (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curdata (nrec n(5), sn_pol c(25), c_i c(25), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
  osn230 c(5), er_c c(3), delta n(11,2), straf n(11,2), cmnt c(50))

 CREATE CURSOR curdata2 (nrec2 n(5), sn_pol c(25), c_i c(25), d_beg d, d_end d, ds c(6), cod n(6), s_all n(11,2), ;
  osn230 c(5), er_c c(3), delta n(11,2), koeff n(3,2), straf n(11,2), cmnt c(50))
 SELECT curdata2
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT merror
 SET RELATION TO recid INTO talon 
 
 m.nrec = 1
 m.nrec2 = 1
 SCAN 
  IF !(et=m.TipOfExp AND docexp=m.docexp)
   LOOP 
  ENDIF 

  m.sn_pol = talon.sn_pol
  m.cod    = cod
  m.c_i    = talon.c_i

  IF !SEEK(m.sn_pol, 'qwert')
   INSERT INTO qwert (sn_pol) VALUES (m.sn_pol)
   m.checked_tot = m.checked_tot + 1
  ENDIF 
   
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

  m.er_c   = err_mee
  m.osn230 = osn230
  m.d_beg  = IIF(!IsMes(m.cod) and !IsVMP(m.cod), talon.d_u,talon.d_u-talon.k_u+1)
  m.d_end  = talon.d_u

  m.ds     = talon.ds   
  m.ns_all = 0
  m.delta  = 0
*  m.s_all  = 0
  m.s_all  = s_all
  m.straf  = 0

  m.opl_tot   = m.opl_tot + s_all && Изменено 12.09.12 по замечанию СОГАЗ - неверная сумма в сводном акте "к оплате"
  
  m.s_1       = s_1
  m.s_2       = s_2

  IF LEFT(UPPER(err_mee),2) != 'W0'
   IF !SEEK(m.sn_pol, 'qwertbad')
    INSERT INTO qwertbad (sn_pol) VALUES (m.sn_pol)
    m.bad_kol = m.bad_kol+1
   ENDIF 

   m.bad_sum   = m.bad_sum + m.s_1
   m.delta     = IIF(m.qcod='S6', m.s_all - m.s_1, m.s_1)
   m.vzaim_tot = m.vzaim_tot +  m.delta

  ELSE 

   m.delta     = IIF(m.qcod='S6', m.s_all - m.s_1, m.s_1)
   m.vzaim_tot = m.vzaim_tot + m.delta

  ENDIF 

  IF m.s_2>0 AND !SEEK(m.sn_pol, 'curdata2')
   m.straf     = m.s_2
   m.tot_straf = m.tot_straf + m.straf
   INSERT INTO curdata2 FROM MEMVAR 
   m.nrec2 = m.nrec2+1
  ENDIF 

  m.s_all = s_all
  m.cmnt = IIF(LEFT(m.er_c,2)='W0','Замечаний нет','')
  INSERT INTO curdata FROM MEMVAR 

  m.nrec = m.nrec+1

 ENDSCAN 
 SET RELATION OFF INTO talon 

* USE IN talon 
* USE IN merror
 USE IN qwert
 USE IN qwertamb
 USE IN qwertst
 USE IN qwertdst
 
 SELECT curdata2
 IF RECCOUNT('curdata2')<=0
  SCATTER MEMVAR 
  INSERT INTO curdata2 FROM MEMVAR 
 ENDIF 
 SET ORDER TO 
 REPLACE ALL nrec2 WITH RECNO()
* BROWSE 
  
 SELECT (ooal)

 m.vzaim_tot = ROUND(m.vzaim_tot,2)
 m.checked_tot = m.checked_amb + m.checked_st + m.checked_dst
 
 m.saystraf = cpr(INT(m.tot_straf))+' '+PADL(INT((m.tot_straf-INT(m.tot_straf))*100),2,'0')+' КОП.'

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl+'\'+m.dotname
 m.lcRepName = m.docname+'.xls'
 
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)

 USE IN curdata 
 USE IN curdata2

 SELECT aisoms

RETURN 