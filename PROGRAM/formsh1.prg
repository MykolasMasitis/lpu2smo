PROCEDURE FormSh1
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ятнплхпнбюрэ нрвер он ябндмни напюыюелнярх?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 ppath = pbase+'\'+m.gcperiod
 IF !fso.FileExists(ppath+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик AISOMS.DBF'+CHR(13)+CHR(10),0+16,m.gcperiod)
  RETURN 
 ENDIF 
 
 IF OpenFile(ppath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
 ENDIF 
 
 SELECT lpuid, mcod, SPACE(120) as lpuname, paz, krank, paz_dst, paz_st FROM aisoms INTO CURSOR svstat READWRITE 
 SELECT svstat
 INDEX on lpuid TAG lpuid
 SET ORDER TO lpuid
 
 USE IN aisoms
 
 FOR nmonth=1 TO tmonth-1
  m.lcperiod = STR(tYear,4)+PADL(tmonth-nmonth,2,'0')
  m.lcpath = pbase+'\'+m.lcperiod
  IF !fso.FileExists(m.lcpath+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.lcpath+'\nsi\sprlpuxx.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.lcpath+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(m.lcpath+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'fil_id')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT aisoms
  SCAN 
   SCATTER MEMVAR 
   IF SEEK(m.lpuid, 'svstat')
    m.o_paz     = svstat.paz
    m.o_paz_amb = svstat.krank
    m.o_paz_dst = svstat.paz_dst
    m.o_paz_st  = svstat.paz_st
    
    m.n_paz     = m.o_paz     + m.paz
    m.n_paz_amb = m.o_paz_amb + m.krank
    m.n_paz_dst = m.o_paz_dst + m.paz_dst
    m.n_paz_st  = m.o_paz_st  + m.paz_st

    UPDATE svstat SET paz=m.n_paz, krank=m.n_paz_amb, paz_dst=m.n_paz_dst, ;
    paz_st=m.n_paz_st WHERE lpuid = m.lpuid
   ELSE 
    IF SEEK(m.lpuid, 'sprlpu')
     m.lpuid = sprlpu.lpu_id
     IF SEEK(m.lpuid, 'svstat')
      m.o_paz     = svstat.paz
      m.o_paz_amb = svstat.krank
      m.o_paz_dst = svstat.paz_dst
      m.o_paz_st  = svstat.paz_st

      m.n_paz     = m.o_paz     + m.paz
      m.n_paz_amb = m.o_paz_amb + m.krank
      m.n_paz_dst = m.o_paz_dst + m.paz_dst
      m.n_paz_st  = m.o_paz_st  + m.paz_st

      UPDATE svstat SET paz=m.n_paz, krank=m.n_paz_amb, paz_dst=m.n_paz_dst, ;
       paz_st=m.n_paz_st WHERE lpuid = m.lpuid
     ENDIF 
    ENDIF 
   ENDIF 
   
  ENDSCAN 
  USE IN aisoms
  USE IN sprlpu
  
 ENDFOR 
 
 
 SELECT svstat
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx.dbf')
  IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
  ELSE 
   SELECT svstat 
   SET RELATION TO lpuid INTO sprlpu
   SCAN
    IF EMPTY(sprlpu.fullname)
     LOOP 
    ENDIF 

    REPLACE lpuname WITH sprlpu.fullname

   ENDSCAN 
   SET RELATION OFF INTO sprlpu
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
  ENDIF 
 ENDIF 
 
 COPY TO &pmee\repssh
 USE 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'нрвер ятнплхпнбюм!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 