PROCEDURE FormSh8
 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ?',4+32,'')=7
  RETURN 
 ENDIF 
 
 CREATE CURSOR curstat (er_c c(2), osn230 c(5), k_u n(6), s_all n(11,2), s_1 n(11,2), s_2 n(11,2))
 INDEX on osn230 TAG osn230
 SET ORDER TO osn230
 
 FOR nmon=1 TO m.tMonth
  m.lcperiod = STR(m.tYear,4)+PADL(m.nmon,2,'0')
  IF fso.FolderExists(pBase+'\'+m.lcperiod)
   WAIT m.lcperiod WINDOW NOWAIT 
   =FormSh8One(m.lcperiod)
   WAIT CLEAR 
  ENDIF 
 ENDFOR 
 SELECT curstat
 COPY TO &pMee\Sh8
 BROWSE 
 USE 
 MESSAGEBOX('OK', 0+64, '')
 
RETURN 

FUNCTION FormSh8One(para1)
 PRIVATE m.lcPeriod
 m.lcPeriod = para1
 IF !fso.FileExists(pBase+'\'+m.lcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
   RETURN 
  ENDIF 
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(pBase+'\'+m.lcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  
  =FormSh8OneLpu(pBase+'\'+m.lcPeriod+'\'+m.mcod)

 ENDSCAN 
 USE IN aisoms
 
RETURN 

FUNCTION FormSh8OneLpu(para1)
 PRIVATE m.ppath
 m.ppath = para1
 IF OpenFile(m.ppath+'\m'+m.mcod, 'merror', 'shar')>0
  IF USED('merror')
   USE IN merror
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(m.ppath+'\talon', 'talon', 'shar', 'recid')>0
  USE IN merror
  IF USED('talon')
   USE IN talon
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 SELECT merror
 SET RELATION TO recid INTO talon 
 SCAN 
  m.et = et
  IF !INLIST(m.et,'4','5','6')
   LOOP 
  ENDIF 
  IF !(talon.tip='5' OR talon.d_type='5' OR talon.cod=1561)
   LOOP 
  ENDIF 
  
  m.er_c   = err_mee
  m.osn230 = osn230
  m.s_all  = s_all
  m.s_1    = s_1
  m.s_2    = s_2
  m.k_u    = 1
  
  IF !SEEK(m.osn230, 'curstat')
   INSERT INTO curstat FROM MEMVAR 
  ELSE 
   m.os_all = curstat.s_all
   m.os_1   = curstat.s_1
   m.os_2   = curstat.s_2
   m.ok_u   = curstat.k_u
   
   m.ns_all = m.os_all + m.s_all
   m.ns_1   = m.os_1   + m.s_all
   m.ns_2   = m.os_2   + m.s_all
   m.nk_u   = m.ok_u   + 1

   UPDATE curstat SET s_all=m.ns_all, s_1=m.ns_1, s_2=m.ns_2, k_u=m.nk_u ;
    WHERE osn230=m.osn230
   
  ENDIF 
  
 ENDSCAN 
 SET RELATION OFF INTO talon 
 USE IN talon 
 USE IN merror
 SELECT aisoms
 
RETURN 