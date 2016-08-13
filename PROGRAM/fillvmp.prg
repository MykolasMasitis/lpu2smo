PROCEDURE FillVmp
 IF MESSAGEBOX('гюонкмхрэ онке VMP тюикю TARIFN'+CHR(13)+CHR(10)+'гмювемхел онкъ VMP146 тюикю USVMPXX?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\usvmpxx.dbf')
  MESSAGEBOX('нрясрярбсер тюик USVMPXX.DBF!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\nsi\tarifn.dbf')
  MESSAGEBOX('нрясрярбсер тюик TARIFN.DBF!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\tarifn', 'tarif', 'shar')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\usvmpxx', 'usvmp', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('usvmp')
   USE IN usvmp
  ENDIF 
  RETURN 
 ENDIF 
 
 m.recsintar = RECCOUNT('tarif')
 m.recsinvmp = RECCOUNT('usvmp')
 
 IF m.recsintar!=m.recsinvmp
  IF MESSAGEBOX('йнк-бн гюохяеи б тюике TARIFN ('+TRANSFORM(m.recsintar,'9999')+')'+CHR(13)+CHR(10)+;
   'ме яннрберярбсер '+CHR(13)+CHR(10)+;
   'йнк-бс гюохяеи б тюике USVMPXX ('+TRANSFORM(m.recsintar,'9999')+')'+CHR(13)+CHR(10)+;
   'опнднкфхрэ?',4+32,'')=7
   USE IN usvmp
   USE IN tarif
   RETURN 
  ENDIF 
 ENDIF 
 
 SELECT tarif
 SET RELATION TO cod INTO usvmp
 SCAN 
  IF EMPTY(usvmp.cod)
   LOOP 
  ENDIF 
  REPLACE vmp WITH usvmp.vmp146
 ENDSCAN 
 SET RELATION OFF INTO usvmp

 USE IN usvmp
 USE IN tarif
 MESSAGEBOX('напюанрйю гюйнмвемю!',0+64,'')
RETURN 
