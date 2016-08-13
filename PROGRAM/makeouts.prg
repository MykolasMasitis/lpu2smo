PROCEDURE MakeOuts
 IF MESSAGEBOX(CHR(13)+CHR(10)+'напюанрюрэ мнлепмхй?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 SET DEFAULT TO (pUpdDir)
 csprfile = ''
 csprfile=GETFILE('dbf')
 IF EMPTY(csprfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'бш мхвецн ме бшапюкх!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 ospr = fso.GetFile(csprfile)
 IF LOWER(LEFT(ospr.name,4)) != 'nomp'
  MESSAGEBOX(CHR(13)+CHR(10)+'щрн ме мнлепмхй!'+CHR(13)+CHR(10),0+16,'nompmmyy')
  RELEASE ospr 
  RETURN 
 ENDIF 
 
 m.mm = SUBSTR(ospr.name,5,2)
 m.yy = SUBSTR(ospr.name,7,2)
 
 RELEASE ospr 

 IF !BETWEEN(INT(VAL(m.mm)),1,12)
  MESSAGEBOX('медносярхлши леяъж!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !INLIST(m.yy,'14','15','16')
  MESSAGEBOX('медносярхлши цнд!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 oSettings.CodePage(csprfile, 866, .t.)

 IF OpenFile(csprfile, 'nomp', 'excl')>0
  RETURN 
 ENDIF 
 
 SELECT nomp
 
 m.lcperiod = '20'+m.yy+m.mm
 COPY STRUCTURE TO &pbase\&lcperiod\nsi\outs
 
 IF OpenFile(pbase+'\'+m.lcperiod+'\nsi\outs', 'outs', 'excl')>0
  IF USED('outs')
   USE IN outs
  ENDIF 
  USE IN nomp 
  RELEASE ospr 
  RETURN 
 ENDIF 
 
 SELECT outs
 INDEX ON s_card + ' ' + PADL(n_card,10,'0') TAG kms
 INDEX ON vsn tag vsn 
 INDEX ON enp TAG enp 

 WAIT "напюанрйю мнлепмхйю..." WINDOW NOWAIT 
 SELECT nomp 
 m.nRecs = 0 
 SCAN 
  IF q!=m.qcod
   LOOP 
  ENDIF 
  m.nRecs = m.nRecs + 1
  SCATTER MEMVAR 
  INSERT INTO outs FROM MEMVAR 
 ENDSCAN 
 WAIT CLEAR 

 USE IN nomp
 USE IN outs

 MESSAGEBOX(CHR(13)+CHR(10)+'напюанрйю мнлепмхйю гюйнмвемю!'+CHR(13)+CHR(10), 0+64, '')

RETURN 