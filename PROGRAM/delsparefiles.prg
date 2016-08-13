PROCEDURE DelSpareFiles

 pComDel= fso.GetParentFolderName(pcommon) + '\COMMON.DEL'

 CREATE CURSOR curFiles (flname c(12))
 SELECT curFiles
 INDEX ON flname TAG flname 
 SET ORDER TO flname 
 
 INSERT INTO curFiles (flname) VALUES ('admokrxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('codku_xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('codotdxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('codwdrxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('dsdisp.dbf')
 INSERT INTO curFiles (flname) VALUES ('dsdisp.cdx')
 INSERT INTO curFiles (flname) VALUES ('emails.dbf')
 INSERT INTO curFiles (flname) VALUES ('emails.cdx')
 INSERT INTO curFiles (flname) VALUES ('isv012xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('kdolgxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('kpreslxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('kspecxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('loggfile.dbf')
 INSERT INTO curFiles (flname) VALUES ('lpudogs.dbf')
 INSERT INTO curFiles (flname) VALUES ('lpudogs.cdx')
 INSERT INTO curFiles (flname) VALUES ('mkb10_xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('ms_mkbxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('nocodrxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('osoerzxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('osoreexx.dbf')
 INSERT INTO curFiles (flname) VALUES ('ososchxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('pnorm.dbf')
 INSERT INTO curFiles (flname) VALUES ('pnorm.cdx')
 INSERT INTO curFiles (flname) VALUES ('polic_dp.dbf')
 INSERT INTO curFiles (flname) VALUES ('polic_h.dbf')
 INSERT INTO curFiles (flname) VALUES ('profotxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('profusxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('prv002xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('prv002xx.cdx')
 INSERT INTO curFiles (flname) VALUES ('readme.tarif')
 INSERT INTO curFiles (flname) VALUES ('rsv009xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('smo.dbf')
 INSERT INTO curFiles (flname) VALUES ('smo.cdx')
 INSERT INTO curFiles (flname) VALUES ('sookodxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('sovmnoxx.dbf')
* INSERT INTO curFiles (flname) VALUES ('spi_cz.dbf')
* INSERT INTO curFiles (flname) VALUES ('spi_cz_ch.dbf')
 INSERT INTO curFiles (flname) VALUES ('spr_ulxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('spraboxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('sprlpuxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('sprsprxx.dbf')
 INSERT INTO curFiles (flname) VALUES ('tarifn.dbf')
 INSERT INTO curFiles (flname) VALUES ('tipno_xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('users.dbf')
 INSERT INTO curFiles (flname) VALUES ('users.cdx')
 INSERT INTO curFiles (flname) VALUES ('usrlpu.dbf')
 INSERT INTO curFiles (flname) VALUES ('usrlpu.cdx')
 INSERT INTO curFiles (flname) VALUES ('volumes.dbf')
 INSERT INTO curFiles (flname) VALUES ('volumes.cdx')
 INSERT INTO curFiles (flname) VALUES ('z_cod_xx.dbf')
 INSERT INTO curFiles (flname) VALUES ('z_dsnoxx.dbf')
 
 oDir        = fso.GetFolder(pCommon)
 DirName     = oDir.Path
 oFilesInDir = oDir.Files
 nFilesInDir = oFilesInDir.Count
 
 FOR EACH oFileInDir IN oFilesInDir
  m.bname = LOWER(ALLTRIM(oFileInDir.Name))
  
  IF !SEEK(PADR(m.bname,12), 'curFiles')
   IF !fso.FolderExists(pComDel)
    fso.CreateFolder(pComDel)
   ENDIF 
   fso.CopyFile(pcommon+'\'+m.bname, pcomdel+'\'+m.bname, .t.)
   fso.DeleteFile(pcommon+'\'+m.bname)
  ENDIF 
  
 NEXT 
 
 USE IN curFiles
 
RETURN 