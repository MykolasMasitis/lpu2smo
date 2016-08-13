PROCEDURE cfgbase
LOCAL m.HomePath
m.HomePath = SUBSTR(SYS(5)+SYS(2003), 1, rat('\',SYS(5)+SYS(2003)))

IF FILE('lpu2smo.cfg')
 USE lpu2smo.cfg IN 0 ALIAS cnfg EXCLUSIVE 
 SELECT cnfg

 IF UPPER(FIELD('tYear')) !=  'TYEAR'
  ALTER TABLE cnfg ADD COLUMN tyear n(4) 
  REPLACE tyear WITH YEAR(DATE())
 ENDIF 
 IF UPPER(FIELD('tMonth')) !=  'TMONTH'
  ALTER TABLE cnfg ADD COLUMN tmonth n(2)
  REPLACE tmonth WITH MONTH(DATE())
 ENDIF 
 IF UPPER(FIELD('tDat1')) !=  'TDAT1'
  ALTER TABLE cnfg ADD COLUMN tdat1 d
  REPLACE tdat1 WITH CTOD('01.'+PADL(MONTH(DATE()),2,'0')+'.'+PADL(YEAR(DATE()),4,'0'))
 ENDIF 
 IF UPPER(FIELD('tDat2')) !=  'TDAT2'
  ALTER TABLE cnfg ADD COLUMN tdat2 d
  REPLACE tdat2 WITH GOMONTH(CTOD('01.'+PADL(MONTH(DATE()),2,'0')+'.'+PADL(YEAR(DATE()),4,'0')),1)-1
 ENDIF 
 IF UPPER(FIELD('User')) !=  'USER'
  ALTER TABLE cnfg ADD COLUMN User c(6)
  REPLACE User WITH 'OMS'
 ENDIF 
 IF UPPER(FIELD('Frmt')) !=  'FRMT'
  ALTER TABLE cnfg ADD COLUMN Frmt c(3)
  REPLACE Frmt WITH 'DOC'
 ENDIF 
 IF UPPER(FIELD('PMEE')) !=  'PMEE'
  ALTER TABLE cnfg ADD COLUMN pmee c(100)
 ENDIF 
 IF UPPER(FIELD('PATTPPL')) !=  'PATTPPL'
  ALTER TABLE cnfg ADD COLUMN pattppl c(100)
  ATTPPLDir = LEFT(pBin,2)+'\ATTPPL\BIN'
  IF fso.FolderExists(ATTPPLDir)
   REPLACE pattppl WITH ATTPPLDir
  ELSE 
   pattfile = GETFILE('cfg','','',0,'Укажите на файл attppl.cfg')
   IF !EMPTY(pattfile)
    ATTPPLDir = LEFT(pattfile,RAT('\',pattfile)-1)
    IF fso.FolderExists(ATTPPLDir)
     REPLACE pattppl WITH ATTPPLDir
    ENDIF 
   ELSE 
   ENDIF 
  ENDIF 
 ENDIF 
 IF UPPER(FIELD('IsUsrDir')) !=  'ISUSRDIR'
  ALTER TABLE cnfg ADD COLUMN IsUsrDir L
  REPLACE IsUsrDir WITH .F.
 ENDIF 
 
ELSE 
IF MESSAGEBOX('ОТСУТСТВУЕТ КОНФИГУРАЦИОННЫЙ ФАЙЛ LPU2SMO.CFG!' + CHR(13) + 'СОЗДАТЬ?',4+32,'ВНИМАНИЕ!') = 6
	CREATE TABLE lpu2smo ;
		(qcod c(2), tyear n(4), tmonth n(4), tdat1 d, tdat2 d, User c(6), Frmt c(3), ;
		 paisoms c(100), parc c(100), pbase c(100), pbin c(100), pcommon c(100), ;
		 pout c(100), plocal c(100), pexpimp c(100), pmail c(100), ptempl c(100), ptrash c(100), ;
		 pdouble c(100), pmee c(100), pattppl c(100), IsUsrDir L)
	SELECT lpu2smo
	USE 
	RENAME  lpu2smo.dbf TO lpu2smo.cfg
	INSERT INTO lpu2smo.cfg (qcod, tyear, tmonth, tdat1, tdat2, user, Frmt, paisoms, parc, pbase, pbin, pcommon, ;
	                         pout, plocal, pexpimp, pmail, ptempl, ptrash, pdouble, pmee, pattppl, IsUsrDir) ;
		VALUES ('P2', YEAR(DATE()), MONTH(DATE()), CTOD('01.'+PADL(MONTH(DATE()),2,'0')+'.'+PADL(YEAR(DATE()),4,'0')), ;
		GOMONTH(CTOD('01.'+PADL(MONTH(DATE()),2,'0')+'.'+PADL(YEAR(DATE()),4,'0')),1)-1, 'OMS', 'DOC', ;
		 m.HomePath+'AISOMS', m.HomePath+'ARC', m.HomePath+'BASE', m.HomePath+'BIN', m.HomePath+'COMMON', ;
		 m.HomePath+'OUT', m.HomePath+'LOCAL', m.HomePath+'EXCHANGE.DIR',  m.HomePath+'MAIL', m.HomePath+'TEMPLATES', ;
		 m.HomePath+'TRASH', m.HomePath+'DOUBLES', m.HomePath+'MEE', '', .F.)
	SELECT lpu2smo
ELSE
	MESSAGEBOX('ПРОДОЛЖЕНИЕ РАБОТЫ НЕВОЗМОЖНО!',0+16,'ВНИМАНИЕ!')
	RETURN -1
ENDIF 
ENDIF 


m.qcod      = ALLTRIM(qcod)
m.pAisOms   = ALLTRIM(paisoms)
m.parc      = ALLTRIM(parc)
m.pbin      = ALLTRIM(pbin)
m.pbase     = ALLTRIM(pbase)
m.pcommon   = ALLTRIM(pcommon)
m.pout      = ALLTRIM(pout)
m.plocal    = ALLTRIM(plocal)
m.pexpimp   = ALLTRIM(pexpimp)
m.pmail     = ALLTRIM(pmail)
m.ptempl    = ALLTRIM(ptempl)
m.ptrash    = ALLTRIM(ptrash)
m.pdouble   = ALLTRIM(pdouble)
m.pmee      = ALLTRIM(pmee)
m.pattppl   = ALLTRIM(pattppl)
m.tyear     = tyear
m.tmonth    = tmonth
m.tdat1     = tdat1
m.tdat2     = tdat2
m.gcUser    = ALLTRIM(User)
m.gcFormat  = UPPER(Frmt)
m.IsUsrDir  = IsUsrDir

IF m.HomePath+'BIN' != m.pBin
	MESSAGEBOX("Директория запуска программы"+CHR(13)+;
	"отличается от директории по умолчанию."+CHR(13)+;
	"Сделать директорию запуска новой директорией по умолчанию?",4+32,"Внимание!")
	
	m.paisoms   = m.HomePath + 'AISOMS'
	m.parc      = m.HomePath + 'ARC'
	m.pbin      = m.HomePath + 'BIN'
	m.pbase     = m.HomePath + 'BASE'
	m.pcommon   = m.HomePath + 'COMMON'
	m.pout      = m.HomePath + 'OUT'
	m.plocal    = m.HomePath + 'LOCAL'
	m.pexpimp   = m.HomePath + 'EXCHANGE.DIR'
	m.pmail     = m.HomePath + 'MAIL'
	m.ptempl    = m.HomePath + 'TEMPLATES'
	m.ptrash    = m.HomePath + 'TRASH'
	m.pdouble   = m.HomePath + 'DOUBLES'
	m.pmee      = m.HomePath + 'MEE'
	m.pattppl   = ''
    m.IsUsrDir  = .F.
	REPLACE paisoms WITH m.paisoms, parc WITH m.parc, pbin WITH m.pbin ,;
		pbase WITH m.pbase, pcommon WITH m.pcommon, pout WITH m.pout ,;
		plocal WITH m.plocal, pexpimp WITH m.pexpimp, pmail WITH m.pmail, ;
		ptempl WITH m.ptempl, ptrash WITH m.ptrash, pdouble WITH m.pdouble, ;
		pmee WITH m.pmee, pattppl WITH m.pattppl, IsUsrDir WITH m.IsUsrDir
	
ENDIF

IF IsUsrDir=.T. AND UPPER(LEFT(m.gcUser,3))='EXP'
 m.usrdir = fso.GetParentFolderName(pbin) + '\'+UPPER(m.gcuser)
 IF !fso.FolderExists(m.usrdir)
  fso.CreateFolder(m.usrdir)
 ENDIF 
 IF !fso.FolderExists(m.usrdir+'\SVACTS')
  fso.CreateFolder(m.usrdir+'\SVACTS')
 ENDIF 
 IF !fso.FolderExists(m.usrdir+'\SSACTS')
  fso.CreateFolder(m.usrdir+'\SSACTS')
 ENDIF 
ENDIF 

USE 

DaemonDir = fso.GetParentFolderName(pbase) + '\DAEMON'
SpamDir   = fso.GetParentFolderName(pbase) + '\SPAM'

IF !fso.FolderExists(DaemonDir)
 fso.CreateFolder(DaemonDir)
ENDIF 

IF !fso.FolderExists(SpamDir)
 fso.CreateFolder(SpamDir)
ENDIF 

=CrLicsDir()

IF !fso.FileExists('usrdata.dbf')
 CREATE TABLE usrdata (recid i AUTOINC, mname c(50))
 INDEX on recid TAG recid 
 SET ORDER TO recid
 m.mname = SYS(0)
 INSERT INTO usrdata (mname) VALUES (m.mname)
 USE 

 #Define MAPI_ORIG 0
 #Define MAPI_TO 1
 #Define MAPI_CC 2
 #Define MAPI_BCC 3

 #Define IMPORTANCE_LOW 0
 #Define IMPORTANCE_NORMAL 1
 #Define IMPORTANCE_HIGH 2

 #DEFINE CRCR CHR(13)

 LOCAL lcMessage, llSuccess
	
 m.lcSubject = m.qcod
 m.lcMessage = "Обновление успешно установлено"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + m.mname

 m.lcAddress    = '9950825@mail.ru'
 
 IF EMCreateMessage(m.lcSubject, m.lcMessage, IMPORTANCE_HIGH)
  IF EMAddRecipient(m.lcAddress, MAPI_TO)
   IF EMSend(.T.)
    m.llSuccess = .T.
   ENDIF
  ENDIF
 ENDIF
	
 IF m.llSuccess
  MESSAGEBOX("Уведомление об обновлении успешно отправлено!", 0+48, "9950825@mail.ru") 
 ELSE
  MESSAGEBOX("Не удалось отправить уведомление!", 64, "9950825@mail.ru")
 ENDIF 
ENDIF 

RETURN 0