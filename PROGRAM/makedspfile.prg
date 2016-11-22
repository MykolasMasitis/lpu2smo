PROCEDURE MakeDspFile
 IF tdat1<{01.01.2014}
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТЧЕТ МОЖЕТ ФОРМИРОВАТЬСЯ НАЧИНАЯ С ЯНВАРЯ 2014!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ ФАЙЛ DSP-ФАЙЛ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.prperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 IF tdat1>{01.01.2014}
  IF !fso.FileExists(pbase+'\'+m.prperiod+'\dsp.dbf')
   MESSAGEBOX(CHR(13)+CHR(10)+'НЕ СФОРМИРОВАН DSP-ФАЙЛ ЗА ПРЕДЫДУЩИЙ ПЕРИОД!'+CHR(13)+CHR(10),0+16,'')
   RETURN 
  ENDIF 
 ENDIF 
 
 lcpath = pbase+'\'+m.gcperiod
 m.period = m.gcperiod
 
 IF !fso.FileExists(lcpath+'\talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'НЕ СФОРМИРОВАН СВОДНЫЙ СЧЕТ ЗА ПЕРИОД!'+CHR(13)+CHR(10),0+16,'talon.dbf')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(lcpath+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'НЕ СФОРМИРОВАН СВОДНЫЙ СЧЕТ ЗА ПЕРИОД!'+CHR(13)+CHR(10),0+16,'people.dbf')
  RETURN 
 ENDIF 

 IF !fso.FileExists(lcpath+'\e'+m.gcperiod+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'НЕ СФОРМИРОВАН СВОДНЫЙ СЧЕТ ЗА ПЕРИОД!'+CHR(13)+CHR(10),0+16,'\e'+m.gcperiod+'.dbf')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(lcpath+'\talon', 'talon', 'shar')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(lcpath+'\people', 'people', 'shar', 'sn_pol')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(lcpath+'\e'+m.gcperiod, 'errsv', 'shar', 'rid')>0
  IF USED('errsv')
   USE IN errsv
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 

 IF RECCOUNT('talon')<=0
  IF USED('talon')
   USE IN talon
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'СВОДНЫЙ ФАЙЛ TALON.DBF ПУСТ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF RECCOUNT('people')<=0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'СВОДНЫЙ ФАЙЛ PEOPLE.DBF ПУСТ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 
 
 dspfile = 'dsp'
 IF fso.FileExists(lcpath+'\'+dspfile+'.dbf')
  fso.DeleteFile(lcpath+'\'+dspfile+'.dbf')
 ENDIF  
 IF !fso.FileExists(lcpath+'\'+dspfile+'.dbf')
  IF INLIST(m.gcperiod,'201401','201501','201601')
  CREATE TABLE &lcpath\&dspfile (recid i, period c(6), mcod c(7), sn_pol c(17), c_i c(30), ;
   fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
    d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3))
  INDEX ON period+mcod+PADL(recid,6,'0') TAG uniqq
  USE 
  ELSE 
   fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.dbf', pbase+'\'+m.gcperiod+'\dsp.dbf')
   fso.CopyFile(pbase+'\'+m.prperiod+'\dsp.cdx', pbase+'\'+m.gcperiod+'\dsp.cdx')
  ENDIF
 ENDIF 

 IF OpenFile(lcpath+'\'+dspfile, 'dsp', 'shar', 'uniqq')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  RETURN 
 ENDIF 

 WAIT "ОБРАБОТКА..." WINDOW NOWAIT 
 SELECT talon
 SET RELATION TO sn_pol INTO people
 SET RELATION TO recid INTO errsv ADDITIVE 
 m.nusl = 0
 SCAN 
  m.cod = cod
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  
  m.tipofcod = dspcodes.tip
  m.rslt     = rslt
  
  DO CASE 
   CASE m.tipofcod = 1 && Диспасеризация взрослых, с июля
    IF !BETWEEN(m.rslt,316,319) AND !BETWEEN(m.rslt,352,358)
     LOOP 
    ENDIF 
   CASE m.tipofcod = 2 && Профосмотры взрослых, с сентября
*    IF tdat1<{01.09.2013}
*     LOOP 
*    ENDIF 
    IF !BETWEEN(m.rslt,343,345)
     LOOP 
    ENDIF 
   CASE m.tipofcod = 3 && Диспасеризация детей-сирот, с июля
    IF !BETWEEN(m.rslt,321,325) AND m.rslt!=320 AND m.rslt!=390 AND !BETWEEN(m.rslt,347,351) && Это - усыновленные сироты!
     LOOP 
    ENDIF 
   CASE m.tipofcod = 4 && Профосмотры несовершеннолетних, с сентября
*    IF tdat1<{01.09.2013}
*     LOOP 
*    ENDIF 
    IF !BETWEEN(m.rslt,332,336) AND m.rslt!=326
     LOOP 
    ENDIF 
   CASE m.tipofcod = 5 && Предварительные профосмотры несовершеннолетних, с сентября
*    IF tdat1<{01.09.2013}
*     LOOP 
*    ENDIF 
*    IF !BETWEEN(m.rslt,337,341) AND m.rslt!=326
    IF !BETWEEN(m.rslt,337,341) AND !INLIST(m.rslt,326,396)
     LOOP 
    ENDIF 
   CASE m.tipofcod = 6 && Периодические профосмотры несовершеннолетних, с сентября
*    IF tdat1<{01.09.2013}
*     LOOP 
*    ENDIF 
    IF m.rslt!=342  AND m.rslt!=326
     LOOP 
    ENDIF 
   CASE m.tipofcod = 7 && профосмотры, с сентября
   OTHERWISE 
    LOOP 
  ENDCASE 

  m.nusl = m.nusl + 1

  m.mcod  = mcod
  m.recid = recid
  m.key   = m.period + m.mcod + PADL(m.recid,6,'0')
  IF SEEK(m.key, 'dsp')
   LOOP 
  ENDIF 

  m.d_u   = d_u
  m.k_u   = k_u
  m.s_all = s_all
  m.er    = ''

  m.id_smo = recid
  m.sn_pol = sn_pol
  m.fam    = people.fam
  m.im     = people.im
  m.ot     = people.ot
  m.w      = people.w
  m.dr     = people.dr
  m.er     = errsv.c_err
   
  m.ages   = YEAR(tdat1) - YEAR(m.dr)
  
  m.c_i    = c_i 

  INSERT INTO dsp FROM MEMVAR 
   
 ENDSCAN 
 SET ORDER TO sn_pol
* SET RELATION OFF INTO people
* SET RELATION OFF INTO errsv
 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'ОБНАРУЖЕНО '+TRANSFORM(m.nusl, '99999')+' УСЛУГ!',0+64,'')
 
 WAIT "РАСЧЕТ ВТОРОГО ЭТАПА..." WINDOW NOWAIT 
 SELECT dsp
 SET ORDER TO 
 SCAN 
  m.rslt = rslt 
*  IF !INLIST(m.rslt,316,320,326,396)
*   LOOP 
*  ENDIF 
  IF !INLIST(m.rslt,316,320,326,346,352,353,354,357,358,390,396)
   LOOP 
  ENDIF 
  m.k_u2 = k_u2
  IF m.k_u2>0
   LOOP 
  ENDIF 
  
  m.sn_pol   = sn_pol
  m.d_u      = d_u
  m.cod      = cod
  m.k_u2     = 0
  m.s_all2   = 0
  m.k_u2ok   = 0
  m.s_all2ok = 0
 
  WAIT "РАСЧЕТ ВТОРОГО ЭТАПА: " + m.sn_pol WINDOW NOWAIT 

  SELECT talon 
  =SEEK(m.sn_pol)
*  m.rslt = rslt

  DO WHILE sn_pol = m.sn_pol
   m.rslt = rslt
   IF d_u < m.d_u 
    SKIP 
    LOOP 
   ENDIF 
   IF cod = m.cod 
    SKIP 
    LOOP 
   ENDIF 
   IF !BETWEEN(m.rslt,317,319) AND !BETWEEN(m.rslt,321,325) AND !BETWEEN(m.rslt,347,351) AND ;
    m.rslt!=346 AND m.rslt!=320 AND m.rslt!=390
    SKIP 
    LOOP 
   ENDIF 
    
   m.k_u2   = m.k_u2 + k_u
   m.s_all2 = m.s_all2 + s_all
    
   IF EMPTY(errsv.rid)
    m.k_u2ok   = m.k_u2ok + k_u
    m.s_all2ok = m.s_all2ok + s_all
   ENDIF 
    
   SKIP 

  ENDDO && talon 

  SELECT dsp 
  REPLACE k_u2 WITH m.k_u2, s_all2 WITH m.s_all2, k_u2ok WITH m.k_u2ok, s_all2ok WITH m.s_all2ok
 
  WAIT CLEAR 
 
 ENDSCAN && dsp
 WAIT CLEAR 

 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('errsv')
  USE IN errsv
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 

RETURN 
