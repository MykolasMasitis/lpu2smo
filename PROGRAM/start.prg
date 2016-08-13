# DEFINE DEBUGMODE .F.
*On Error Do ErrorHandler With Error( ), ;
							Message( ), ;
							Message(1), ;
							Program( ), ;
							Lineno(1)

ON SHUTDOWN DO ExitProg

RELEASE ALL EXTENDED
CLEAR ALL
SET TALK OFF
SET HOURS TO 24
SET DATE TO GERMAN
SET CENTURY ON 
SET CONSOLE OFF
SET RESOURCE OFF 
SET safety OFF 
SET REPROCESS TO 3 SECONDS 
SET DELETED ON 
SET ESCAPE OFF  

DECLARE INTEGER GetPrivateProfileString IN Win32API  AS GetPrivStr ;
	STRING cSection, STRING cKey, STRING cDefault, STRING @cBuffer, ;
	INTEGER nBufferSize, STRING cINIFile
DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
	STRING cSection, STRING cKey, STRING cValue, STRING cINIFile
DECLARE INTEGER GetSysColor IN User32.DLL INTEGER

DECLARE ScreenSize In Tools32 ;
	Integer @nW, ;  && Øèðèíà
	Integer @nH     && Âûñîòà

PUBLIC fso AS SCRIPTING.FileSystemObject, wshell AS Shell.Application

fso      = CREATEOBJECT('Scripting.FileSystemObject')
WShell   = CREATEOBJECT('Shell.Application')
WSHShell = CREATEOBJECT('Wscript.Shell')

SET PROCEDURE TO Utils

PUBLIC nWidth, nHeight, nDifSize, IsNotePad
m.nWidth    = 0
m.nHeight   = 0
m.nDifSize  = 800-768
m.IsNotePad = .F.
=ScreenSize(@nWidth, @nHeight)

m.IsNotePad = IIF(m.nHeight<800, .T., .F.)
*m.IsNotePad = .T.

WITH _SCREEN
 .Width      = 1024
 .Height     = (800-m.nDifSize)-100
 .BackColor  = RGB(192,192,192)
 .AutoCenter = .f.
 .Picture    = 'lpu2smo.jpg'
 .Visible    = .t.
ENDWITH 

lcPathSystem = sys(5) + sys(2003)
Set DEFAULT TO (lcPathSystem)
lcPathMain = sys(5) + sys(2003)
lcPathSystem = lcPathMain+;
	','+(lcPathMain+'\BITMAPS')+;
	','+(lcPathMain+'\FORMS')+;
	','+(lcPathMain+'\GRAPHICS')+;
	','+(lcPathMain+'\INCLUDE')+;
	','+(lcPathMain+'\LIBS')+;
	','+(lcPathMain+'\MENU')+;
	','+(lcPathMain+'\PROGRAM')

SET PATH TO (lcPathSystem)

PUBLIC paisoms, parc, pbase, pbin, pcommon, pout, ptempl, ptrash, pdouble, pmee, pexpimp, DaemonDir, SpamDir, ;
 tyear, tmonth, tdat1, tdat2, curlpu, qcod, qmail, qobjid, UserERZ, qname, oMenu, gcPeriod, gcUser, gcFormat,;
 usrfam, usrim, usrot, usrfio, pattppl, m.ynorm

m.ynorm = 0

PUBLIC ARRAY mes_text[12], mes_main[12]
mes_text[1]="ÿíâàðÿ"
mes_text[2]="ôåâðàëÿ"
mes_text[3]="ìàðòà"
mes_text[4]="àïðåëÿ"
mes_text[5]="ìàÿ"
mes_text[6]="èþíÿ"
mes_text[7]="èþëÿ"
mes_text[8]="àâãóñòà"
mes_text[9]="ñåíòÿáðÿ"
mes_text[10]="îêòÿáðÿ"
mes_text[11]="íîÿáðÿ"
mes_text[12]="äåêàáðÿ"

mes_main[1]="ßíâàðü"
mes_main[2]="Ôåâðàëü"
mes_main[3]="Ìàðò"
mes_main[4]="Àïðåëü"
mes_main[5]="Ìàé"
mes_main[6]="Èþíü"
mes_main[7]="Èþëü"
mes_main[8]="Àâãóñò"
mes_main[9]="Ñåíòÿáðü"
mes_main[10]="Îêòÿáðü"
mes_main[11]="Íîÿáðü"
mes_main[12]="Äåêàáðü"

numlib = adir(alib,lcPathMain+'\LIBS\*.vcx')
for i = 1 to numlib
	lcSetLibrary = lcPathMain+'\LIBS\' + alib(i,1)
	set classlib to (lcSetLibrary) additive
endfor

*lcSetLibrary = lcPathMain+'\LIBS\vfpcompression.fll'
*SET LIBRARY TO &lcPathMain\LIBS\vfpcompression
SET LIBRARY TO &lcPathMain\vfpzip
SET LIBRARY TO &lcPathMain\vfpexmapi ADDITIVE

*lcSetLibrary = lcPathMain+'\LIBS\vfpexmapi.fll'
*SET LIBRARY TO (lcSetLibrary) ADDITIVE 

IF CfgBase() = -1
 =ExitProg()
ENDIF 

IF !fso.FolderExists(pcommon)
 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÑÒÂÓÅÒ ÈËÈ ÍÅÄÎÑÒÓÏÍÀ'+CHR(13)+CHR(10)+'ÄÈÐÅÊÒÎÐÈß '+pcommon,0+16,'')
 =ExitProg()
ENDIF  

m.tdat1 = CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0'))
m.tdat2 = GOMONTH(CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0')),1)-1
m.gcPeriod = STR(tYear,4)+PADL(tMonth,2,'0')

DO CASE 
 CASE m.qcod == 'P2'
  m.qname = 'ÎÀÎ "ÌÅÄÈÖÈÍÑÊÀß ÑÒÐÀÕÎÂÀß ÊÎÌÏÀÍÈß "ÓÐÀËÑÈÁ"'
  m.qmail = 'skpomed.msk.oms'
  m.qobjid = 3386
 CASE m.qcod == 'P3'
  m.qname = 'ÎÎÎ ÌÑÎ "ÏÀÍÀÖÅß" ÌÎÑÊÎÂÑÊÈÉ ÔÈËÈÀË'
  m.qmail = 'panacea.msk.oms'
  m.qobjid = 5387
 CASE m.qcod == 'I3'
  m.qname = 'ÎÎÎ ÑÊ "ÈÍÃÎÑÑÒÐÀÕ-Ì"'
  m.qmail = 'ingos.msk.oms'
  m.qobjid = 5398
 CASE m.qcod == 'I1'
  m.qname = 'ÎÎÎ ÌÑÊ "ÈÊÀÐ"'
  m.qmail = 'ikar.msk.oms'
  m.qobjid = 110
 CASE m.qcod == 'R4'
  m.qname = 'ÎÎÎ "ÑÒÐÀÕÎÂÀß ÌÅÄÈÖÈÍÑÊÀß ÊÎÌÏÀÍÈß ÐÅÑÎ-ÌÅÄ" ÌÎÑÊÎÂÑÊÈÉ ÔÈËÈÀË'
  m.qmail = 'reso.msk.oms'
  m.qobjid = 3415
 CASE m.qcod == 'S7'
  m.qname = 'ÎÀÎ ÑÊ "ÑÎÃÀÇ-Ìåä"'
  m.qmail = 'sogaz.msk.oms'
  m.qobjid = 5400
 OTHERWISE 
  m.qname = 'ÎÀÎ "ÌÅÄÈÖÈÍÑÊÀß ÑÒÐÀÕÎÂÀß ÊÎÌÏÀÍÈß "ÓÐÀËÑÈÁ"'
  m.qmail = 'skpomed.msk.oms'
  m.qobjid = 3386
ENDCASE 

public goApp
goApp = NEWOBJECT('_goapp','main')
ADDPROPERTY(goApp, "mcod", "")

*goApp.Show()
goApp.Begin_process()

=chkbase()

IF !fso.FileExists(pCommon+'\Users.cdx')
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT Users 
  INDEX ON name TAG name 
  USE 
 ENDIF 
ENDIF 

IF !fso.FileExists(pCommon+'\pnyear.dbf')
 CREATE TABLE &pCommon\pnyear (period c(6), pnorm n(13,2))
 INDEX on period TAG period 
 SET ORDER TO period
 INSERT INTO pnyear (period, pnorm) VALUES ('2013', 9594.08) 
 INSERT INTO pnyear (period, pnorm) VALUES ('2014', 11321.33) 
 INSERT INTO pnyear (period, pnorm) VALUES ('2015', 13191.01) 
 USE 
ELSE 
 IF OpenFile(pCommon+'\pnyear', 'pnyear', 'shar', 'period')>0
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ELSE 
  SELECT pnyear
  IF SEEK(STR(tYear,4), 'pnyear')
   m.ynorm = pnyear.pnorm
  ENDIF 
  IF USED('pnyear')
   USE IN pnyear
  ENDIF 
 ENDIF 
ENDIF 

=OpenFile(pcommon+'\Users', 'users', 'shar')
SELECT users
IF VARTYPE(fam) != 'C'
 USE 
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT users
  ALTER TABLE Users ADD COLUMN fam c(25)
  USE 
 ENDIF 
ENDIF 
USE 

=OpenFile(pcommon+'\Users', 'users', 'shar')
SELECT users
IF VARTYPE(im) != 'C'
 USE 
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT users
  ALTER TABLE Users ADD COLUMN im c(25)
  USE 
 ENDIF 
ENDIF 
USE 

=OpenFile(pcommon+'\Users', 'users', 'shar')
SELECT users
IF VARTYPE(ot) != 'C'
 USE 
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT users
  ALTER TABLE Users ADD COLUMN ot c(25)
  USE 
 ENDIF 
ENDIF 
USE 

=OpenFile(pcommon+'\Users', 'users', 'shar')
SELECT users
IF VARTYPE(fio) != 'C'
 USE 
 IF OpenFile(pcommon+'\Users', 'users', 'excl') <= 0
  SELECT users
  ALTER TABLE Users ADD COLUMN fio c(40)
  INDEX ON name TAG name 
  USE 
 ENDIF 
ENDIF 
USE 

=OpenFile(pCommon+'\Users', 'Users', 'shar', 'name')
SELECT Users
IF !SEEK(m.gcUser, 'Users')
 USE 
 MESSAGEBOX('ÈÌß '+ALLTRIM(m.gcUser)+' ÎÒÑÓÒÑÒÂÓÅÒ Â ÑÏÐÀÂÎ×ÍÈÊÅ!', 0+16, '')
 =ExitProg()
ELSE 
 IF !RLOCK()
  USE 
  MESSAGEBOX('ÏÎËÜÇÎÂÀÒÅËÜ '+ALLTRIM(m.gcUser)+' ÓÆÅ ÐÀÁÎÒÀÅÒ Â ÑÈÑÒÅÌÅ!', 0+16, '')
  =ExitProg()
 ELSE 
  m.usrfam = ALLTRIM(Fam)
  m.usrim = ALLTRIM(im)
  m.usrot = ALLTRIM(ot)
  m.usrfio = ALLTRIM(fio)
 ENDIF 
ENDIF 


=OpenFile(pCommon+'\smo', 'smo', 'shar')
SELECT smo
IF VARTYPE(chieftip) != 'C'
 USE 
 IF OpenFile(pCommon+'\smo', 'smo', 'excl')<=0
  SELECT smo
  ALTER TABLE smo ADD COLUMN chieftip c(100)
  USE 
 ENDIF 
 =OpenFile(pCommon+'\smo', 'smo', 'shar')
 SELECT smo
ENDIF 
IF VARTYPE(chiefname) != 'C'
 USE 
 IF OpenFile(pCommon+'\smo', 'smo', 'excl')<=0
  SELECT smo
  ALTER TABLE smo ADD COLUMN chiefname c(100)
  USE 
 ENDIF 
 =OpenFile(pCommon+'\smo', 'smo', 'shar')
 SELECT smo
ENDIF 
IF VARTYPE(buhtip) != 'C'
 USE 
 IF OpenFile(pCommon+'\smo', 'smo', 'excl')<=0
  SELECT smo
  ALTER TABLE smo ADD COLUMN buhtip c(100)
  USE 
 ENDIF 
 =OpenFile(pCommon+'\smo', 'smo', 'shar')
 SELECT smo
ENDIF 
IF VARTYPE(buhname) != 'C'
 USE 
 IF OpenFile(pCommon+'\smo', 'smo', 'excl')<=0
  SELECT smo
  ALTER TABLE smo ADD COLUMN buhname c(100)
  USE 
 ENDIF 
 =OpenFile(pCommon+'\smo', 'smo', 'shar')
 SELECT smo
ENDIF 
COUNT FOR v != .f. TO kol_q
PUBLIC smo(kol_q, 2)
COPY FOR v != .f. FIELDS code, name TO ARRAY smo
USE 

SET SYSMENU TO
SET SYSMENU ON
SET STATUS BAR ON 
WITH _SCREEN
* .Width      = 1024
* .Height     = 768-100
* .BackColor  = RGB(192,192,192)
 .Icon = 'cross.ico'
* .AutoCenter = .f.
 .Caption = m.qname+', ÏÎËÜÇÎÂÀÒÅËÜ: '+ALLTRIM(m.gcUser)+', ÏÅÐÈÎÄ: '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' ÃÎÄÀ'+;
  ' (ñ '+DTOC(tdat1)+' ïî '+DTOC(tdat2)+')'
ENDWITH 

*m.mmy=PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
*pcommonmmy = pcommon+m.mmy
*WAIT "ÑÎÇÄÀÍÈÅ ÄÈÐÅÊÒÎÐÈÈ "+pcommonmmy WINDOW NOWAIT 
*IF !fso.FolderExists (pcommonmmy)
* fso.CopyFolder(pcommon, pcommonmmy)
* pcommon = pcommonmmy
* IF OpenFile(pbin+'\lpu2smo.cfg', 'llpu', 'shar')==0
*  REPLACE pcommon WITH pcommonmmy
*  USE 
* ENDIF 
*ENDIF 
*WAIT CLEAR 

PUBLIC IsAdmin
m.IsAdmin=.f.
IF fso.FileExists(pbin+'\admin')
 ffile = fso.GetFile(pbin+'\admin')
  IF ffile.size == 4
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(4)
   fhandl.Close
*   MESSAGEBOX(lcHead,0+64,'')
   IF lcHead == 'ruby'
    m.IsAdmin = .t.
   ENDIF 
  ENDIF 
ENDIF 

*DO DelSpareFiles

DO m_menu
*DO demomenu1
 
READ EVENTS

=ExitProg()

*Clear all
*RELEASE ALL EXTENDED

FUNCTION ExitProg
 IF USED('Users')
  USE IN Users
 ENDIF 
 RELEASE m.oError && ???
 #IF DEBUGMODE
*  _SCREEN.Caption = oApp.cOldWindCaption
  SET SYSMENU TO DEFAULT
  _SCREEN.TitleBar = 1
  _SCREEN.WindowState = 2
  _SCREEN.LockScreen = .F.
  _SCREEN.Picture = ''
  _SCREEN.BackColor = RGB(255,255,255)
*  oApp.ShowToolBars()
  SET SYSMENU ON
 #ELSE
 _SCREEN.Caption = ""
 #ENDIF
* oApp.CloseAllTable()
* RELEASE m.oApp
 RELEASE m.goApp
 #IF !DEBUGMODE
  ON SHUTDOWN
  QUIT
 #ELSE
  ON SHUTDOWN
  _SCREEN.Icon =""
  _SCREEN.FirstStart = .T.
  *SET HELP TO
  CLEAR ALL
  CLOSE ALL
  CLEAR PROGRAM
  SET SYSMENU NOSAVE
  SET SYSMENU TO DEFAULT
  SET SYSMENU ON
 #ENDIF
RETURN 
