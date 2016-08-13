PROCEDURE AppPolisH
 IF MESSAGEBOX(''+CHR(13)+CHR(10)+'бш унрхре намнбхрэ '+;
  'яопюбнвмхй POLIC_H?',4+32,'')==7
  RETURN 
 ENDIF 
 
 oal = SYS(5)+SYS(2003)
 SET DEFAULT TO (pbase+'\'+gcperiod+'\'+'nsi')
 hnew = GETFILE('dbf','','',0,'сЙЮФХРЕ МЮ ТЮИК!')
 SET DEFAULT TO (oal)
 
 IF EMPTY(hnew)
  MESSAGEBOX('бш мхвецн ме бшапюкх!',0+64,'')
  RETURN 
 ENDIF 
 
 tnresult = 0
 tnresult = tnresult + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\polic_h', 'h_svod', 'shared', 'sn_pol')
 tnresult = tnresult + OpenFile(hnew, 'h_new', 'shar')
 
 IF tnresult>0
  IF USED('h_svod')
   USE IN h_svod
  ENDIF 
  IF USED('h_new')
   USE IN h_new
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT h_new
 IF VARTYPE(sn_pol)!='C'
  USE
  USE IN dp_svod
  MESSAGEBOX('б тюике '+hnew+CHR(13)+CHR(10)+;
   'нрярсрярбсер менаундхлне онке SN_POL!'+CHR(13)+CHR(10)+;
   'опнднкфемхе пюанрш мебнглнфмн!',0+16, '')
  RETURN 
 ENDIF 
 IF VARTYPE(D_U)!='D'
  USE 
  USE IN h_svod
  MESSAGEBOX('б тюике '+hnew+CHR(13)+CHR(10)+;
   'нрярсрярбсер менаундхлне онке D_U!'+CHR(13)+CHR(10)+;
   'опнднкфемхе пюанрш мебнглнфмн!',0+16, '')
  RETURN 
 ENDIF 
 
 SELECT h_new
 m.totrecs = RECCOUNT()
 m.addrecs = 0
 PRIVATE lpu_id, sn_pol, qq, d_u, tms, year
 SCAN 
  SCATTER MEMVAR 
*  m.sn_pol = sn_pol
*  m.d_u = d_u
  IF !SEEK(m.sn_pol, 'h_svod')
*   INSERT INTO h_svod (sn_pol, d_u) VALUES (m.sn_pol, m.d_u)
   INSERT INTO h_svod FROM MEMVAR 
   m.addrecs = m.addrecs + 1
  ELSE 
   DELETE IN h_svod
   INSERT INTO h_svod (sn_pol, d_u) VALUES (m.sn_pol, m.d_u)
   m.addrecs = m.addrecs + 1
  ENDIF 
 ENDSCAN 
 USE 
 USE IN h_svod
 
 MESSAGEBOX('днаюбкемн '+ALLTRIM(STR(m.addrecs))+ ' гюохяеи б тюик POLIC_H'+CHR(13)+CHR(10)+;
  'хг '+ALLTRIM(STR(m.totrecs))+ ' б мел опхясрярбсчыху!',0+64, '')

RETURN 