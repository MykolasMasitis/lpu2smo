PROCEDURE LoadMeFiles
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÇÀÃÐÓÇÈÒÜ ÄÀÍÍÛÅ ÏÎ ÑÍßÒÈßÌ Â ÀÏÑÔ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(13)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+pbase+'\'+gcperiod+'!'+CHR(13)+CHR(10),0+16,'') 
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(13)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'') 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 m.totmee  = 0
 m.totekmp = 0
 SELECT aisoms
 SCAN 
  m.totmee  = m.totmee  + e_mee
  m.totekmp = m.totekmp + e_ekmp
 ENDSCAN 
 
 IF m.totmee>0 OR m.totekmp>0
  IF MESSAGEBOX(CHR(13)+CHR(10)+'Â ÁÀÇÅ ÓÆÅ ÑÎÄÅÐÆÀÒÑß ÄÀÍÍÛÅ ÏÎ ÑÍßÒÈßÌ.'+CHR(13)+CHR(10)+;
  'ÂÛ ÓÂÅÐÅÍÛ Â ÍÅÎÁÕÎÄÈÌÎÑÒÈ ÏÎÂÒÎÐÍÎÉ ÇÀÃÐÓÇÊÈ?',4+32,'')=7
   USE IN aisoms
   RETURN 
  ENDIF 
 ENDIF 

 IF !fso.FolderExists(pout+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(13)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+pout+'\'+gcperiod+'!'+CHR(13)+CHR(10),0+16,'') 
  RETURN 
 ENDIF 
 
 SCAN 
*  REPLACE e_mee WITH 0, e_ekmp WITH 0,;
   ambmeesum WITH 0, stmeesum WITH 0, dstmeesum WITH 0,;
   ambbadmee WITH 0, stbadmee WITH 0, dstbadmee WITH 0
  REPLACE e_mee WITH 0, e_ekmp WITH 0

  m.lpuid = lpuid
  m.mcod  = mcod
  mefile = 'me'+UPPER(m.qcod)+STR(lpuid,4)
  IF !fso.FileExists(pout+'\'+gcperiod+'\'+mefile+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pout+'\'+gcperiod+'\'+mefile, 'mefls', 'shar')>0
   IF USED('mefls')
    USE IN mefls
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 

  MESSAGEBOX(CHR(13)+CHR(10)+UPPER(mefile)+'.DBF'+CHR(13)+CHR(10),0+64,'')

  m.ambmeesum = 0
  m.stmeesum  = 0
  m.dstmeesum = 0
  m.emee  = 0
  m.eekmp = 0

  CREATE CURSOR curamb (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR curst (c_i c(25))
  INDEX ON c_i TAG c_i
  SET ORDER TO c_i
  CREATE CURSOR curdst (sn_pol c(25))
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol

  SELECT mefls
  SCAN 
   m.prd = period
   IF m.prd != m.gcperiod
    LOOP 
   ENDIF 
   m.eperiod = period_e
   m.c_i     = c_i
   m.sn_pol  = sn_pol
   m.cod     = cod
   
   DO CASE 
    CASE IsUsl(m.cod)
     IF !SEEK(m.sn_pol, 'curamb')
      INSERT INTO curamb (sn_pol) VALUES (m.sn_pol)
     ENDIF 
    CASE IsKd(m.cod)
     IF !SEEK(m.sn_pol, 'curdst')
      INSERT INTO curdst (sn_pol) VALUES (m.sn_pol)
     ENDIF 
    CASE IsMes(m.cod) OR IsVMP(m.cod)
     IF !SEEK(m.c_i, 'curst')
      INSERT INTO curst (c_i) VALUES (m.c_i)
     ENDIF 
   ENDCASE 

   m.et     = et
   m.sexp   = s_opl_e

   m.emee   = m.emee  + IIF(INLIST(m.et,'2','3'), m.sexp, 0)
   m.eekmp  = m.eekmp + IIF(INLIST(m.et,'4','5','6'), m.sexp, 0)

   m.ambmeesum = m.ambmeesum + IIF(IsUsl(m.cod), m.sexp, 0)
   m.stmeesum  = m.stmeesum  + IIF(IsMes(m.cod) or IsVMP(m.cod), m.sexp, 0)
   m.dstmeesum = m.dstmeesum + IIF(IsKd(m.cod), m.sexp, 0)
  ENDSCAN 
  USE IN mefls
  
  m.ambbadmee = RECCOUNT('curamb')
  m.dstbadmee = RECCOUNT('curdst')
  m.stbadmee  = RECCOUNT('curst')
  
  USE IN curamb
  USE IN curst
  USE IN curdst
  
  m.ppath = pbase+'\'+m.eperiod+'\'+m.mcod
  m.mmfile = 'm'+m.mcod
  IF fso.FileExists(pbase+'\'+m.eperiod+'\aisoms.dbf')
   IF OpenFile(pbase+'\'+m.eperiod+'\aisoms', 'asoms', 'shar', 'lpuid')>0
    IF USED('asoms')
     USE IN asoms
    ENDIF 
   ELSE 
    SELECT asoms
    IF SEEK(m.mcod, 'asoms')
     REPLACE ambmeesum WITH m.ambmeesum, stmeesum WITH m.stmeesum, dstmeesum WITH m.dstmeesum,;
      ambbadmee WITH m.ambbadmee, stbadmee WITH m.stbadmee, dstbadmee WITH m.dstbadmee
    ENDIF 
    USE IN asoms
   ENDIF 
  ENDIF 
  IF fso.FileExists(m.ppath+'\'+m.mmfile+'.dbf')
   IF OpenFile(m.ppath+'\'+m.mmfile, 'mmfile', 'shar')>0
    IF USED('mmfile')
     USE IN mmfile
    ENDIF 
   ENDIF 
   SELECT mmfile
   SCAN 
    IF et!=m.et
     LOOP 
    ENDIF 
    REPLACE e_period WITH m.eperiod
   ENDSCAN 
   USE IN mmfile
   MESSAGEBOX(CHR(13)+CHR(10)+UPPER(m.mmfile)+'.DBF'+CHR(13)+CHR(10),0+64,'')
  ENDIF 

  SELECT aisoms
*  REPLACE e_mee WITH m.emee, e_ekmp WITH m.eekmp,;
   ambmeesum WITH m.ambmeesum, stmeesum WITH m.stmeesum, dstmeesum WITH m.dstmeesum,;
   ambbadmee WITH m.ambbadmee, stbadmee WITH m.stbadmee, dstbadmee WITH m.dstbadmee
  REPLACE e_mee WITH m.emee, e_ekmp WITH m.eekmp
   
 ENDSCAN 

 USE IN aisoms

RETURN 