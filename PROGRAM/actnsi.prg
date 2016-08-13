PROCEDURE ActNSI
 IF MESSAGEBOX(CHR(13)+CHR(10)+'юйрсюкхгхпнбюрэ мях?'+CHR(13)+CHR(10),4+32,'')=7
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
 IF LOWER(LEFT(ospr.name,6)) != 'sprspr'
  MESSAGEBOX(CHR(13)+CHR(10)+'щрн ме яопюбнвмхй мях!'+CHR(13)+CHR(10),0+16,'sprsprxx')
  RELEASE ospr 
  RETURN 
 ENDIF 
 
 oSettings.CodePage(csprfile, 866, .t.)
 
 IF OpenFile(csprfile, 'spr', 'excl')>0
  =Exit()
  RETURN 
 ENDIF 
 
 SELECT spr
 
 nfields = FCOUNT('spr')
 IF nfields != 6
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'йнкхвеярбн онкеи ('+STR(nfields,1)+') ме пюбмн 6!',0+16,'')
  =Exit()
  RETURN 
 ENDIF 
 
 IF !checkspr()
  RETURN 
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'тюик опюбхкэмши!'+CHR(13)+CHR(10),0+64,'')
 
 odir = ospr.ParentFolder.Path
 
 CREATE CURSOR nfiles (fname c(8))
 INDEX ON fname TAG fname
 SET ORDER TO fname
 
 INSERT INTO nfiles (fname) VALUES ('admokr')
 INSERT INTO nfiles (fname) VALUES ('codku_')
 INSERT INTO nfiles (fname) VALUES ('codotd')
 INSERT INTO nfiles (fname) VALUES ('codwdr')
 INSERT INTO nfiles (fname) VALUES ('hopff_')
 INSERT INTO nfiles (fname) VALUES ('isv012')
 INSERT INTO nfiles (fname) VALUES ('kdolg')
 INSERT INTO nfiles (fname) VALUES ('kpresl')
* INSERT INTO nfiles (fname) VALUES ('kspec')
 INSERT INTO nfiles (fname) VALUES ('mkb10_')
 INSERT INTO nfiles (fname) VALUES ('modpac')
 INSERT INTO nfiles (fname) VALUES ('ms_mkb')
 INSERT INTO nfiles (fname) VALUES ('nocodr')
 INSERT INTO nfiles (fname) VALUES ('osoerz')
 INSERT INTO nfiles (fname) VALUES ('osoree')
 INSERT INTO nfiles (fname) VALUES ('ososch')
 INSERT INTO nfiles (fname) VALUES ('profot')
 INSERT INTO nfiles (fname) VALUES ('profus')
 INSERT INTO nfiles (fname) VALUES ('promed	')
 INSERT INTO nfiles (fname) VALUES ('prv002')
 INSERT INTO nfiles (fname) VALUES ('reesms')
 INSERT INTO nfiles (fname) VALUES ('reesus')
 INSERT INTO nfiles (fname) VALUES ('reesvp')
 INSERT INTO nfiles (fname) VALUES ('rsv009')
 INSERT INTO nfiles (fname) VALUES ('sookod')
 INSERT INTO nfiles (fname) VALUES ('sovmno')
 INSERT INTO nfiles (fname) VALUES ('spr_ul')
 INSERT INTO nfiles (fname) VALUES ('sprabo')
 INSERT INTO nfiles (fname) VALUES ('sprlpu')
 INSERT INTO nfiles (fname) VALUES ('spv015')
 INSERT INTO nfiles (fname) VALUES ('tarimu')
 INSERT INTO nfiles (fname) VALUES ('tersmo')
 INSERT INTO nfiles (fname) VALUES ('tipgrp')
 INSERT INTO nfiles (fname) VALUES ('tipno_')
 INSERT INTO nfiles (fname) VALUES ('z_dsno')
 INSERT INTO nfiles (fname) VALUES ('usvmp')
 INSERT INTO nfiles (fname) VALUES ('vidvp_')
 INSERT INTO nfiles (fname) VALUES ('z_cod_')
 INSERT INTO nfiles (fname) VALUES ('z_dsno')
 INSERT INTO nfiles (fname) VALUES ('territ')
 
 SELECT spr
 SCAN 
  m.name_eta = ALLTRIM(name_eta)
  m.fname = PADR(LOWER(LEFT(m.name_eta, LEN(m.name_eta)-2)),8)
  IF SEEK(m.fname, 'nfiles')
   m.etname    = ALLTRIM(nfiles.fname)
   m.intr_data = intr_data
   IF fso.FileExists(odir+'\'+m.name_eta+'.dbf')
    IF m.intr_data<=tdat1
     MESSAGEBOX(m.name_eta+'.dbf'+' -> '+m.etname+'xx.dbf' ,0+64,'')
     fso.CopyFile(odir+'\'+m.name_eta+'.dbf', pcommon+'\'+m.etname+'xx.dbf', .t.)
    ENDIF 
   ENDIF 
  ENDIF 
 ENDSCAN 
 
 =Exit() 
RETURN 

FUNCTION checkspr
 F1Name = FIELD(1)
 IF F1Name != 'SCOD'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе оепбнцн онкъ ('+F1Name+') днкфмн ашрэ SCOD!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F1Type = VARTYPE(&F1Name)
 IF F1Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо оепбнцн онкъ '+F1Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F1Size = FSIZE(F1Name)
 IF F1Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F1Name+' днкфмю ашрэ 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F2Name = FIELD(2)
 IF F2Name != 'CUR_VER'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе брнпнцн онкъ ('+F2Name+') днкфмн ашрэ CUR_VER!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F2Type = VARTYPE(&F2Name)
 IF F2Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F2Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F2Size = FSIZE(F2Name)
 IF F2Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F2Name+' днкфмю ашрэ 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F3Name = FIELD(3)
 IF F3Name != 'FULL_NAME'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе рперэецн онкъ ('+F3Name+') днкфмн ашрэ FULL_NAME!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F3Type = VARTYPE(&F3Name)
 IF F3Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F3Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F3Size = FSIZE(F3Name)
 IF F3Size != 120
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F3Name+' днкфмю ашрэ 120!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F4Name = FIELD(4)
 IF F4Name != 'INTR_DATA'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе вербепрнцн онкъ ('+F4Name+') днкфмн ашрэ INTR_DATA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F4Type = VARTYPE(&F4Name)
 IF F4Type!='D'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F4Name+' днкфем ашрэ DATE!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F5Name = FIELD(5)
 IF F5Name != 'NAME_ETA'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе оърнцн онкъ ('+F5Name+') днкфмн ашрэ NAME_ETA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F5Type = VARTYPE(&F5Name)
 IF F5Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F5Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F5Size = FSIZE(F5Name)
 IF F5Size != 8
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F5Name+' днкфмю ашрэ 8!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

 F6Name = FIELD(6)
 IF F6Name != 'CRC_ETA'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'мюхлемнбюмхе ьеярнцн онкъ ('+F6Name+') днкфмн ашрэ CRC_ETA!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F6Type = VARTYPE(&F6Name)
 IF F6Type!='C'
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'рхо онкъ '+F6Name+' днкфем ашрэ CHAR!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 
 F6Size = FSIZE(F6Name)
 IF F6Size != 10
  MESSAGEBOX(CHR(13)+CHR(10)+'мебепмюъ ярпсйрспю тюикю '+UPPER(ospr.name)+CHR(13)+CHR(10)+;
   'пюглепмнярэ онкъ '+F6Name+' днкфмю ашрэ 10!',0+16,'')
  =Exit()
  RETURN .f.
 ENDIF 

RETURN .t.

FUNCTION exit 
 IF USED('spr')
  USE IN spr 
 ENDIF 
 RELEASE ospr  
RETURN 
