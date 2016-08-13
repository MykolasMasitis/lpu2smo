PROCEDURE FFOMS18022016mek 
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ'+CHR(13)+CHR(10)+;
  'ÏÎ ÇÀÏÐÎÑÓ ÔÔÎÌÑ ÎÒ 18.02.2016 ÄËß ÑÌÎ?',4+32,'ÌÝÊ')=7
  RETURN 
 ENDIF 
 
 m.ddat0 = {01.01.2015}
 m.ddat  = {01.01.2015}
 CREATE CURSOR def2015 (osn230 c(5), k_u n(5), s_all n(11,2), s_1 n(11,2), s_2 n(11,2))
 INDEX on osn230 TAG osn230
 SET ORDER TO osn230
 m.defcur = 'def2015'

 FOR m.nm = 0 TO 12
  m.ddat     = GOMONTH(m.ddat0,m.nm)
  m.lcperiod = LEFT(DTOS(m.ddat),6)
  IF fso.FolderExists(pBase+'\'+m.lcperiod)
   =OnePeriod(m.lcperiod)
  ENDIF 
 ENDFOR  

 m.ddat0 = {01.01.2016}
 m.ddat  = {01.01.2016}
 CREATE CURSOR def2016 (osn230 c(5), k_u n(5), s_all n(11,2), s_1 n(11,2), s_2 n(11,2))
 INDEX on osn230 TAG osn230
 SET ORDER TO osn230
 m.defcur = 'def2016'

 FOR m.nm = 0 TO 3
  m.ddat     = GOMONTH(m.ddat0,m.nm)
  m.lcperiod = LEFT(DTOS(m.ddat),6)
  IF fso.FolderExists(pBase+'\'+m.lcperiod)
   =OnePeriod(m.lcperiod)
  ENDIF 
 ENDFOR  
 
 SELECT def2015
 COPY TO &pOut\mek2015
 USE
 SELECT def2016
 COPY TO &pOut\mek2016
 USE 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

FUNCTION OnePeriod(m.lcperiod) 
 IF !fso.FileExists(pBase+'\'+m.lcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(pBase+'\'+m.lcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  =OneLpu(m.mcod)
 ENDSCAN 
 USE IN aisoms 

RETURN 

FUNCTION OneLpu(m.mcod)
 IF OpenFile(pBase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'merror', 'shar', 'rid')>0
  IF USED('merror')
   USE IN merror 
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
  USE IN merror
  IF USED('talon')
   USE IN talon 
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.lcperiod+'\nsi\sookodxx', 'sookod', 'shar', 'er_c')>0
  USE IN merror
  USE IN talon 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 CREATE CURSOR unik (rid i)
 INDEX on rid TAG rid 
 SET ORDER TO rid 

 SELECT merror
 SET RELATION TO rid INTO talon 
 SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 
 SCAN 
  m.f = f
  IF m.f!='S'
   LOOP 
  ENDIF 
  m.rid = rid 
  m.osn230 = LEFT(sookod.osn230,5)
  m.k_u    = 1
  IF !SEEK(m.rid, 'unik')
   INSERT INTO unik FROM MEMVAR 
   m.s_all  = talon.s_all
   m.s_1    = talon.s_all
   m.s_2   = 0
  ELSE 
   m.s_all = 0
   m.s_1   = 0
   m.s_2   = 0
  ENDIF 
  IF !SEEK(m.osn230, '&defcur')
   INSERT INTO &defcur FROM MEMVAR 
  ELSE 
   m.ok_u   = &defcur..k_u
   m.os_all = &defcur..s_all
   m.os_1   = &defcur..s_1
   m.os_2   = &defcur..s_2
   
   m.nk_u   = m.ok_u   + m.k_u
   m.ns_all = m.os_all + m.s_all
   m.ns_1   = m.os_1   + m.s_1
   m.ns_2   = m.os_2   + m.s_2

   UPDATE &defcur SET k_u=m.nk_u, s_all=m.ns_all, s_1=m.ns_1, s_2=m.ns_2 WHERE osn230=m.osn230
   
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO sookod 
 SET RELATION OFF INTO talon 

 USE IN merror 
 USE IN talon 
 USE IN sookod
 
 USE IN unik 

 SELECT aisoms 
RETURN 