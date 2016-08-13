PROCEDURE CorrDsp
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÏÐÎÂÅÑÒÈ ÊÎÐÐÅÊÒÈÐÎÂÊÓ DSP.DBF?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\dsp.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÔÀÉË DSP ÍÅ ÑÂÔÎÐÌÈÐÎÂÀÍ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\dsp', 'dsp', 'shar')>0
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curmcod (period c(6), lpuid n(4), mcod1 c(7), mcod2 c(7))
 INDEX on period+mcod1 TAG vir
 SET ORDER TO vir
 
 SELECT dsp
 SET RELATION TO mcod INTO sprlpu
 
 SCAN 
  IF !EMPTY(sprlpu.mcod)
   LOOP 
  ENDIF 
  m.period = period
  m.mcod    = mcod
  m.vir = m.period + m.mcod
  IF m.period=m.gcperiod
   LOOP 
  ENDIF 

  IF !SEEK(m.vir, 'curmcod')
   IF fso.FolderExists(pbase+'\'+m.period+'\nsi')
    IF fso.FileExists(pbase+'\'+m.period+'\nsi\sprlpuxx.dbf')
     IF OpenFile(pbase+'\'+m.period+'\nsi\sprlpuxx', 'sprlpp', 'shar', 'mcod')<=0
      m.lpuid = IIF(SEEK(m.mcod, 'sprlpp'), sprlpp.lpu_id, 0)
      m.mcod2 = IIF(SEEK(m.lpuid, 'sprlpu', 'lpu_id'), sprlpu.mcod, '')
      INSERT INTO curmcod (period, lpuid, mcod1, mcod2) VALUES (m.period, m.lpuid, m.mcod, m.mcod2)
      IF USED('sprlpp')
       USE IN sprlpp
      ENDIF 
     ENDIF 
     IF USED('sprlpp')
      USE IN sprlpp
     ENDIF 
     SELECT dsp 
     IF !EMPTY(m.mcod2)
      REPLACE mcod WITH m.mcod2
     ENDIF 
    ENDIF 
   ENDIF 
  ELSE 
   m.mcod = curmcod.mcod2
   REPLACE mcod WITH m.mcod2
  ENDIF 
  
 ENDSCAN 
 SET RELATION OFF INTO sprlpu
 
 USE 
 USE IN sprlpu
 
 SELECT curmcod
 BROWSE 
 USE 

RETURN 