PROCEDURE ActsJournal

 IF MESSAGEBOX(CHR(13)+CHR(10)+'’Œ“»“≈ —Œ«ƒ¿“‹  ¿“¿ÀŒ√ ¿ “Œ¬?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ Ã››!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF  
 
 MailDir  = fso.GetFolder(pmee)
 oPeriods = MailDir.SubFolders
 nPeriods = oPeriods.Count
  
 IF nPeriods <= 0
  RETURN 
 ENDIF 
 
 IF fso.FileExists(pmee+'\acts.dbf')
  fso.DeleteFile(pmee+'\acts.dbf')
 ENDIF 
 IF fso.FileExists(pmee+'\_acts.dbf')
  fso.DeleteFile(pmee+'\_acts.dbf')
 ENDIF 
 IF fso.FileExists(pmee+'\__acts.dbf')
  fso.DeleteFile(pmee+'\__acts.dbf')
 ENDIF 
 
 CREATE TABLE &pmee\_acts ;
  (period c(6), mcod c(7), codexp n(1), tipacc n(1), isok l, issv l, sn_pol c(25), fam c(25), im c(20), ot c(20),;
   actname c(250), actdate t)
 USE 
 =OpenFile(pmee+'\_acts', 'acts', 'excl')
 
 FOR EACH oPeriod IN oPeriods
  m.period       = oPeriod.Name
  oMcodInPeriods = oPeriod.SubFolders
  
  FOR EACH oMcodInPeriod IN oMcodInPeriods
   cmcod = oMcodInPeriod.name
   oFilesInMcod = oMcodInPeriod.Files
   nFilesInMcod = oFilesInMcod.Count

   m.mcod = oMcodInPeriod.Name
   
   m.lIsFExists = .f.
   IF fso.FileExists(pbase+'\'+m.period+'\'+m.mcod+'\people.dbf')
    IF OpenFile(pbase+'\'+m.period+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')==0
     m.lIsFExists = .t.
    ENDIF 
   ENDIF 
    
   
   FOR EACH oFileInMcod IN oFilesInMcod
    m.actname = oFileInMcod.name
    m.filepath = oFileInMcod.Path
    m.actdate = oFileInMcod.DateLastModified
    m.fileext  = UPPER(RIGHT(ALLTRIM(oFileInMcod.name),4))
    
    IF UPPER(LEFT(m.actname,3))!='ACT' OR fileext != '.DOC'    
     LOOP 
    ENDIF 

    m.isok = IIF(OCCURS('OK',UPPER(m.actname))==1, .t., .f.)

    DO CASE 
     CASE INLIST(UPPER(m.actname), 'ACTSV', 'ACTSSOK_', 'ACTSSPLOK_','ACTSSPL_','ACTSS_')
      m.codexp = 2
     CASE INLIST(UPPER(m.actname), 'ACTSSTGOK_', 'ACTSSTG_')
      m.codexp = 3
     CASE INLIST(UPPER(m.actname), 'ACTEKMPSV', 'ACTEKMPPLAN', 'ACTEKMP_','ACTEKMPOK_')
      m.codexp = 4
     CASE LEFT(UPPER(m.actname),11) == 'ACTEKMPTARG'
      m.codexp = 5
     CASE LEFT(UPPER(m.actname),11) == 'ACTEKMPTHEM'
      m.codexp = 6
     OTHERWISE 
      m.codexp = 0
    ENDCASE 
    
    DO CASE 
     CASE OCCURS('Ò‚',m.actname)==1
      m.tipacc = 0
     CASE OCCURS('‡Ï·',m.actname)==1
      m.tipacc = 1
     CASE OCCURS('‰ÒÚ',m.actname)==1
      m.tipacc = 2
     CASE OCCURS('ÒÚ',m.actname)==1 AND OCCURS('‰ÒÚ',m.actname)==0
      m.tipacc = 3
     OTHERWISE 
      m.tipacc = 0
    ENDCASE 

    m.IsSv = IIF(OCCURS('SV',UPPER(m.actname))==1, .t., .f.)

    IF !m.IsSv AND OCCURS('_',m.actname)>=3
     m.sn_pol = SUBSTR(m.actname, AT('_', m.actname,2)+1,AT('_', m.actname,3)-AT('_', m.actname,2)-1)
    ELSE 
     m.sn_pol = ''
    ENDIF 
    
    m.fam = ''
    m.im  = ''
    m.ot  = ''
    IF !EMPTY(m.sn_pol) AND  m.lIsFExists
     IF SEEK(PADR(ALLTRIM(m.sn_pol),25), 'people')
      m.fam = people.fam
      m.im  = people.im
      m.ot  = people.ot
     ELSE 
      IF SEEK(PADR(LEFT(ALLTRIM(m.sn_pol),6)+' '+SUBSTR(ALLTRIM(m.sn_pol),7),25), 'people')
       m.fam = people.fam
       m.im  = people.im
       m.ot  = people.ot
      ENDIF 
     ENDIF 
    ENDIF 

    INSERT INTO acts FROM MEMVAR 
    
   ENDFOR 

   IF m.lIsFExists = .t.
    USE IN people
   ENDIF 

  ENDFOR 

 ENDFOR 

 SELECT acts 
 INDEX ON actdate TAG actdate 
 SET ORDER TO actdate
 COPY TO &pmee\__acts
 SET ORDER TO 
 DELETE TAG ALL 
 USE 

 fso.DeleteFile(pmee+'\_acts.dbf')
 
 CREATE TABLE &pmee\ssacts\ssacts ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1, period c(6), mcod c(7), codexp n(1), tipacc n(1), isok l, sn_pol c(25), ;
   fam c(25), im c(20), ot c(20), actname c(250), actdate t)
 INDEX ON recid TAG recid 
 INDEX ON period TAG period
 INDEX ON mcod TAG mcod 
 INDEX ON sn_pol TAG sn_pol
 INDEX ON actdate TAG actdate
 INDEX ON PADR(ALLTRIM(fam)+' '+LEFT(im,1)+LEFT(ot,1),28) TAG fio 
 USE 
 =OpenFile(pmee+'\ssacts\ssacts', 'ssacts', 'shar')
 
 CREATE TABLE &pmee\svacts\svacts ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1, period c(6), mcod c(7), codexp n(1), actname c(250), actdate t)
 INDEX ON recid TAG recid 
 INDEX ON period TAG period
 INDEX ON mcod TAG mcod 
 INDEX ON actdate TAG actdate
 USE 
 =OpenFile(pmee+'\svacts\svacts', 'svacts', 'shar')

 =OpenFile(pmee+'\__acts', 'acts', 'shar')
 SELECT acts 
 SCAN 
  SCATTER MEMVAR 
  IF issv
   INSERT INTO svacts FROM MEMVAR 
  ELSE 
   INSERT INTO ssacts FROM MEMVAR 
  ENDIF 
 ENDSCAN 
 USE 
 fso.DeleteFile(pmee+'\__acts.dbf')

 SELECT ssacts
 WAIT " Œœ»–Œ¬¿Õ»≈ ¿ “Œ¬..." WINDOW NOWAIT 
 SCAN 
  WAIT " Œœ»–Œ¬¿Õ»≈ ¿ “¿ "+ALLTRIM(actname)+'...' WINDOW NOWAIT 
  m.name1 = pmee+'\'+period+'\'+mcod+'\'+ALLTRIM(actname)
  m.name2 = pmee+'\ssacts\'+PADL(recid,6,'0')+'.doc'
  fso.CopyFile(m.name1, m.name2)
  WAIT CLEAR 
 ENDSCAN 
 USE 
 
 SELECT svacts
 WAIT " Œœ»–Œ¬¿Õ»≈ ¿ “Œ¬..." WINDOW NOWAIT 
 SCAN 
  WAIT " Œœ»–Œ¬¿Õ»≈ ¿ “¿ "+ALLTRIM(actname)+'...' WINDOW NOWAIT 
  m.name1 = pmee+'\'+period+'\'+mcod+'\'+ALLTRIM(actname)
  m.name2 = pmee+'\svacts\'+PADL(recid,6,'0')+'.doc'
  fso.CopyFile(m.name1, m.name2)
  WAIT CLEAR 
 ENDSCAN 
 USE 

 MESSAGEBOX(CHR(13)+CHR(10)+'√Œ“Œ¬Œ!'+CHR(13)+CHR(10),0+64,'')

RETURN 