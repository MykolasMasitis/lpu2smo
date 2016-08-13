FUNCTION ChMcod(para1, para2)
 
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÇÀÌÅÍÈÒÜ MCOD?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.oldmcod = m.para1
 m.newmcod = m.para1

 DO FORM NewMcod
 
 IF m.oldmcod = m.newmcod
  RETURN 
 ENDIF 
 
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÇÀÌÅÍÈÒÜ MCOD C '+m.oldmcod+CHR(13)+CHR(10)+;
  'ÍÀ '+m.newmcod+'?',4+32,'')=7
  RETURN 
 ENDIF 
 
 orec = RECNO('aisoms')
 IF SEEK(m.newmcod, 'aisoms', 'mcod')
  GO (orec)
  MESSAGEBOX('ÒÀÊÎÉ MCOD ÓÆÅ ÑÓÙÅÑÒÂÓÅÒ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 GO (orec)
 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.oldmcod)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+pbase+'\'+m.gcperiod+'\'+m.oldmcod+'!'+;
   'ÄÀËÜÍÅÉØÀß ÎÁÐÀÁÎÒÊÀ ÍÅÂÎÇÌÎÆÍÀ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.newmcod)
  MESSAGEBOX('ÄÈÐÅÊÒÎÐÈß '+pbase+'\'+m.gcperiod+'\'+m.newmcod+' ÓÆÅ ÑÓÙÅÑÒÂÓÅÒ!'+;
   'ÄÀËÜÍÅÉØÀß ÎÁÐÀÁÎÒÊÀ ÍÅÂÎÇÌÎÆÍÀ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.CopyFolder(pbase+'\'+m.gcperiod+'\'+m.oldmcod, pbase+'\'+m.gcperiod+'\'+m.newmcod)
  MESSAGEBOX('ÍÅ ÓÄÀËÎÑÜ ÑÊÎÏÈÐÎÂÀÒÜ ÄÈÐÅÊÒÎÐÈÞ '+CHR(13)+CHR(10)+;
   pbase+'\'+m.gcperiod+'\'+m.oldmcod+CHR(13)+CHR(10)+;
   'Â '+pbase+'\'+m.gcperiod+'\'+m.oldmcod++'!'+CHR(13)+CHR(10)+;
   'ÄÀËÜÍÅÉØÀß ÎÁÐÀÁÎÒÊÀ ÍÅÂÎÇÌÎÆÍÀ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 m.mmy = PADL(m.tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\B'+m.oldmcod+'.'+m.mmy)
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\B'+m.newmcod+'.'+m.mmy)
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\B'+m.oldmcod+'.'+m.mmy, ;
    pbase+'\'+m.gcperiod+'\'+m.newmcod+'\B'+m.newmcod+'.'+m.mmy)
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\B'+m.oldmcod+'.'+m.mmy)
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.oldmcod+'.dbf')
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.newmcod+'.dbf')
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.oldmcod+'.dbf', ;
    pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.newmcod+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.oldmcod+'.dbf')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.oldmcod+'.cdx')
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.newmcod+'.cdx')
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.oldmcod+'.cdx', ;
    pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.newmcod+'.cdx')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\e'+m.oldmcod+'.cdx')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.oldmcod+'.dbf')
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.newmcod+'.dbf')
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.oldmcod+'.dbf', ;
    pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.newmcod+'.dbf')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.oldmcod+'.dbf')
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.oldmcod+'.cdx')
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.newmcod+'.cdx')
   fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.oldmcod+'.cdx', ;
    pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.newmcod+'.cdx')
   fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\m'+m.oldmcod+'.cdx')
  ENDIF 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\people.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT people
   SCAN 
    IF mcod!=m.newmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
    IF prmcod=m.oldmcod
     REPLACE prmcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN people
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\talon.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.newmcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT talon
   SCAN 
    IF mcod!=m.newmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN talon
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\people.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT people
   SCAN 
    IF prmcod=m.oldmcod
     REPLACE prmcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN people
   SELECT aisoms
  ENDIF 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\talon.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT talon
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN talon
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\otdel.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\otdel', 'otdel', 'shar')>0
   IF USED('otdel')
    USE IN otdel
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT otdel
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN otdel
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\doctor.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\doctor', 'doctor', 'shar')>0
   IF USED('doctor')
    USE IN doctor
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT doctor
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN doctor
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\dsp.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\dsp', 'dsp', 'shar')>0
   IF USED('dsp')
    USE IN dsp
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT dsp
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN dsp
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\pr4.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar')>0
   IF USED('pr4')
    USE IN pr4
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT pr4
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN pr4
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\ud'+m.qcod+m.gcperiod+'.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\ud'+m.qcod+m.gcperiod, 'udfile', 'shar')>0
   IF USED('udfile')
    USE IN udfile
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT udfile
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN udfile
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\e'+m.gcperiod+'.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\e'+m.gcperiod, 'err', 'shar')>0
   IF USED('err')
    USE IN err
   ENDIF 
   SELECT aisoms
  ELSE 
   SELECT err
   SCAN 
    IF mcod=m.oldmcod
     REPLACE mcod WITH m.newmcod
    ENDIF 
   ENDSCAN 
   USE IN err
   SELECT aisoms
  ENDIF 
 ENDIF 
 
 SELECT aisoms
 SCAN
  m.mcod = mcod
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  WAIT m.mcod+'...' WINDOW NOWAIT 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  SELECT people
  SCAN 
   IF prmcod=m.oldmcod
    REPLACE mcod WITH m.newmcod
   ENDIF 
  ENDSCAN 
  USE IN people 
  WAIT CLEAR 
  SELECT aisoms
 ENDSCAN  
 
 UPDATE aisoms SET mcod=m.newmcod WHERE mcod=m.oldmcod
 
 IF fso.FileExists(pcommon+'\dsdisp.dbf')
  IF OpenFile(pcommon+'\dsdisp', 'dsdisp', 'shar')>0
  ELSE 
   UPDATE dsdisp SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN dsdisp
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pcommon+'\emails.dbf')
  IF USED('emails')
   UPDATE emails SET mcod=m.newmcod WHERE mcod=m.oldmcod
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pcommon+'\lpudogs.dbf')
  IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar')>0
  ELSE 
   UPDATE lpudogs SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN lpudogs
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\pilot.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar')>0
  ELSE 
   UPDATE pilot SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN pilot
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pcommon+'\spi_lpu_dd_2014.dbf')
  IF OpenFile(pcommon+'\spi_lpu_dd_2014', 'spi', 'shar')>0
  ELSE 
   UPDATE spi SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN spi
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pcommon+'\sprlpuxx.dbf')
  IF OpenFile(pcommon+'\sprlpuxx', 'spi', 'shar')>0
  ELSE 
   UPDATE spi SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN spi
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pcommon+'\usrlpu.dbf')
  IF OpenFile(pcommon+'\usrlpu', 'spi', 'shar')>0
  ELSE 
   UPDATE spi SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN spi
   SELECT aisoms
  ENDIF 
 ENDIF 
 
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\mo_vmp.dbf') 
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\mo_vmp', 'movmp', 'shar')>0
   IF USED('movmp')
    USE IN movmp
   ENDIF 
  ELSE 
   UPDATE movmp SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN movmp
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\spi_lpu_dd.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\spi_lpu_dd', 'spi', 'shar')>0
  ELSE 
   UPDATE spi SET mcod=m.newmcod WHERE mcod=m.oldmcod
   USE IN spi
   SELECT aisoms
  ENDIF 
 ENDIF 

 IF USED('sprlpu')
  UPDATE sprlpu SET mcod=m.newmcod WHERE mcod=m.oldmcod
  SELECT aisoms
 ENDIF 

 MESSAGEBOX('OK!',0+64,'')

* REPLACE mcod WITH m.newmcod

RETURN 