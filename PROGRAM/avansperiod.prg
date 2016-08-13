PROCEDURE AvansPeriod
 IF MESSAGEBOX('лндскэ гюцпсфюер юбюмянбше окюрефх, '+CHR(13)+CHR(10)+;
		    'мю рейсыхи оепхнд.'+CHR(13)+CHR(10)+;
            '(гюонкмъеряъ онке S_AVANS)'+CHR(13)+CHR(10)+;
            'щрн рн, врн бш деиярбхрекэмн унрхре?'+CHR(13)+CHR(10)+;
            ''+CHR(13)+CHR(10)+;
            '',4+32, '') == 7
   RETURN 
  ENDIF 

 oal = SYS(5)+SYS(2003)
* SET DEFAULT TO (OutDirPeriod)
 AvansFile = GETFILE('dbf','','',0,'сЙЮФХРЕ МЮ ТЮИК!')
 SET DEFAULT TO (oal)
 
 IF EMPTY(AvansFile)
  MESSAGEBOX('бш мхвецн ме бшапюкх!',0+64,'')
  RETURN 
 ENDIF 
 
 tnresult = 0
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shared', 'mcod')
 tnresult = tnresult + OpenFile(AvansFile, 'AFile', 'excl')
 
 IF tnresult>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('AFile')
   USE IN AFile
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT afile
 IF VARTYPE(mcod)!='C'
  USE 
  USE IN aisoms
  MESSAGEBOX('б тюике '+AvansFile+CHR(13)+CHR(10)+;
   'нрярсрярбсер менаундхлне онке MCOD!'+CHR(13)+CHR(10)+;
   'опнднкфемхе нпюанрш мебнглнфмн!',0+16, '')
  RETURN 
 ENDIF 
 IF VARTYPE(s_avans)!='N'
  USE 
  USE IN aisoms
  MESSAGEBOX('б тюике '+AvansFile+CHR(13)+CHR(10)+;
   'нрярсрярбсер менаундхлне онке S_AVANS!'+CHR(13)+CHR(10)+;
   'опнднкфемхе нпюанрш мебнглнфмн!',0+16, '')
  RETURN 
 ENDIF 
 
 m.recsafile = RECCOUNT('afile')
 m.recsffile = RECCOUNT('aisoms')
 
 IF m.recsafile != m.recsffile
  IF MESSAGEBOX('йнкхвеярбн гюохяеи б тюике '+AvansFile +'- '+PADL(m.recsafile,3,'0')+CHR(13)+CHR(10)+;
  'ме пюбмн йнкхвеярбс гюохяеи б тюике AISOMS.DBF!'+'- '+PADL(m.recsffile,3,'0')+'!'+CHR(13)+CHR(10)+;
  'опнднкфхрэ?'+CHR(13)+CHR(10)+;
  ''+CHR(13)+CHR(10),4+32, 'гЮДСЛЮИРЕЯЭ!') == 7
   MESSAGEBOX('напюанрйю опепбюмю!',0+16,'')
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   IF USED('afile')
    USE IN afile
   ENDIF 
   RETURN 
  ENDIF 
 ENDIF 

 SELECT afile 
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod

 SELECT aisoms
 SET RELATION TO mcod INTO afile
 m.notnullavanc = 0
 REPLACE ALL s_avans WITH afile.s_avans
 SET RELATION OFF INTO afile
 USE

 SELECT afile
 SET ORDER TO 
 DELETE TAG all
 USE 
* USE IN sprlpu

 MESSAGEBOX('опнярюбкемн мемскебшу юбюмянб: '+PADL(m.notnullavanc,3,'0'), 0+64, '')

RETURN 