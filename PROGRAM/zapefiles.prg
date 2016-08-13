FUNCTION ZapEFiles

IF MESSAGEBOX('ÝÒÀ ÏÐÎÖÅÄÓÐÀ ÔÈÇÈ×ÅÑÊÈ ÓÄÀËßÅÒ'+CHR(13)+CHR(10)+;
 'ÂÑÅ ÇÀÏÈÑÈ Â ÔÀÉËÀÕ ÎØÈÁÎÊ.'+CHR(13)+CHR(10)+'ÏÐÎÄÎËÆÈÒÜ?',4+32, '')==7
 RETURN 
ENDIF 

IF MESSAGEBOX(''+CHR(13)+CHR(10)+;
 'ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?'+CHR(13)+CHR(10)+;
 ''+CHR(13)+CHR(10),4+32, '')==7
 RETURN 
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\errsv', 'errsv', 'excl')>0
* RETURN 
ELSE 
 SELECT errsv
 ZAP
 USE 
ENDIF 

IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar') > 0
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\nsi\UsrLpu', "UsrLpu", "shar", "mcod") > 0
 USE IN aisoms
 RETURN
ENDIF 

SELECT AisOms
SCAN
 m.mcod = mcod
* m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
* IF m.usr != m.gcUser AND m.gcUser!='OMS'
*  LOOP 
* ENDIF 

 WAIT m.mcod WINDOW NOWAIT 

 lcPath = pbase+'\'+m.gcperiod+'\'+mcod
 IF fso.FileExists(lcPath+'\e'+mcod+'.dbf')
  =ZapEFile(lcPath)
 ENDIF 
 IF fso.FileExists(lcPath+'\ctrl'+m.qcod+'.dbf')
  =ZapCtrlFile(lcPath)
 ENDIF 
 SELECT AisOms
 REPLACE sum_flk WITH 0
 WAIT CLEAR 
ENDSCAN 
USE IN AisOms
USE IN UsrLpu

RETURN 


FUNCTION ZapEFile(lcPath)

 tc_mcod = SUBSTR(lcPath, RAT('\', lcPath)+1)
 
 IF OpenFile(lcPath+'\e'+m.tc_mcod, "Error", "excl") == 0
  SELECT Error
  ZAP 
  USE IN error
 ENDIF

RETURN 

FUNCTION ZapCtrlFile(lcPath)

 IF OpenFile(lcPath+'\ctrl'+m.qcod, "Error", "excl") == 0
  SELECT Error
  ZAP 
  USE IN error
 ENDIF

RETURN 
