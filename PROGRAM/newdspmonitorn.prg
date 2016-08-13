PROCEDURE NewDspMonitorN
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ХОТИТЕ СФОРМИРОВАТЬ ОТЧЕТ'+CHR(13)+CHR(10)+;
 'ПО МОНИТОРНИГУ ДИСПАНСЕРИЗАЦИИ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF MESSAGEBOX('ОТЧЕТ С МЭЭ?',4+16,'')=7
  m.lIsMee = .f.
 ELSE 
  m.lIsMee = .t.
 ENDIF 

 IF !fso.FolderExists(pbase+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ ПЕРИОДА!'+CHR(13)+CHR(10),0+16,gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\dsp.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ DSP.DBF!'+CHR(13)+CHR(10),0+16,gcperiod)
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\dsp', 'dsp', 'shar')>0
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 

 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))

 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 DotName = 'dispmonitor.xlt'
 DocName = 'dspmon'+m.qcod+m.mmyy
 IF !fso.FileExists(ptempl+'\'+dotname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ШАБЛОН ОТЧЕТА DISPMONITOR.XLT'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 WAIT "ЗАПУСК EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 
 m.IsVisible = .t. 
 m.IsQuit    = .f.

 oDoc = oExcel.WorkBooks.Add(pTempl+'\'+DotName)
 
 =MakePage1()
 =MakePage2Old()
 =MakePage3()
 =MakePage4()
 =MakePage5()

 IF m.lIsMee = .t.
 =MakePage6()
 =MakePage7()
 =MakePage8()
 =MakePage9()
 =MakePage10()
 ENDIF 

 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 

 IF fso.FileExists(pbase+'\'+gcperiod+'\'+DocName+'.xls')
  fso.DeleteFile(pbase+'\'+gcperiod+'\'+DocName+'.xls')
 ENDIF 
 oDoc.SaveAs(pbase+'\'+gcperiod+'\'+DocName,18)
 WAIT CLEAR 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 

RETURN 

FUNCTION MakePage1
 WAIT "Формирование листа 1" WINDOW NOWAIT 
 
 CREATE CURSOR curdsp (sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 CREATE CURSOR curlpu1 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 CREATE CURSOR curlpu2 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 

 CREATE CURSOR curlpu3 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 CREATE CURSOR curpaz1 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz2 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz3 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz4 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz5 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz6 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz7 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz8 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz9 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz10 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz11 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 CREATE CURSOR curpaz12 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 SELECT dsp
 
 DIMENSION dimdsp(3,26)
 dimdsp = 0

 SCAN 
  m.cod    = cod
  m.sn_pol = sn_pol
  m.mcod = mcod

  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF !INLIST(m.tipofcod,4,5,6)
   LOOP 
  ENDIF 

  DO CASE 

   CASE m.tipofcod =  4
    IF !SEEK(m.mcod, 'curlpu1')
     INSERT INTO curlpu1 (mcod) VALUES (m.mcod)
    ENDIF 
    IF rslt=326
     dimdsp(1,21) = dimdsp(1,21) + 1
     LOOP 
    ENDIF 
    IF !INLIST(rslt,332,333,334,335,336)
     LOOP 
    ENDIF 

    dimdsp(1,2) = dimdsp(1,2) + 1
    dimdsp(1,3) = dimdsp(1,3) + s_all
    IF EMPTY(k_u2)
     dimdsp(1,6) = dimdsp(1,6) + 1
     IF !SEEK(m.sn_pol, 'curpaz1')
      INSERT INTO curpaz1 FROM MEMVAR 
      dimdsp(1,7) = dimdsp(1,7) + 1
     ENDIF 
     dimdsp(1,8) = dimdsp(1,8) + s_all
    ELSE 
     dimdsp(1,9) = dimdsp(1,9) + 1
     IF !SEEK(m.sn_pol, 'curpaz2')
      INSERT INTO curpaz2 FROM MEMVAR 
      dimdsp(1,10) = dimdsp(1,10) + 1
     ENDIF 
     dimdsp(1,11) = dimdsp(1,11) + s_all2
    ENDIF 
    IF EMPTY(er)
     dimdsp(1,13) = dimdsp(1,13) + 1
     dimdsp(1,14) = dimdsp(1,14) + s_all
     IF EMPTY(k_u2)
      dimdsp(1,15) = dimdsp(1,15) + 1
      IF !SEEK(m.sn_pol, 'curpaz3')
       INSERT INTO curpaz3 FROM MEMVAR 
       dimdsp(1,16) = dimdsp(1,16) + 1
      ENDIF 
      dimdsp(1,17) = dimdsp(1,17) + s_all
     ELSE 
      dimdsp(1,18) = dimdsp(1,18) + 1
      IF !SEEK(m.sn_pol, 'curpaz4')
       INSERT INTO curpaz4 FROM MEMVAR 
       dimdsp(1,19) = dimdsp(1,19) + 1
      ENDIF 
*      dimdsp(1,19) = dimdsp(1,19) + k_u2
      dimdsp(1,20) = dimdsp(1,20) + s_all2
     ENDIF 
     DO CASE 
      CASE rslt=326
       dimdsp(1,21) = dimdsp(1,21) + 1
      CASE rslt=332
       dimdsp(1,22) = dimdsp(1,22) + 1
      CASE rslt=333
       dimdsp(1,23) = dimdsp(1,23) + 1
      CASE rslt=334
       dimdsp(1,24) = dimdsp(1,24) + 1
      CASE rslt=335
       dimdsp(1,25) = dimdsp(1,25) + 1
      CASE rslt=336
       dimdsp(1,26) = dimdsp(1,26) + 1
      OTHERWISE 
       dimdsp(1,22) = dimdsp(1,22) + 1
     ENDCASE 
    ENDIF 
    
   CASE m.tipofcod = 5
    IF !SEEK(m.mcod, 'curlpu2')
     INSERT INTO curlpu2 (mcod) VALUES (m.mcod)
    ENDIF 
    IF rslt=396
     dimdsp(2,21) = dimdsp(2,21) + 1
    LOOP 
    ENDIF 
    IF !INLIST(rslt,337,338,339,340,341)
     LOOP 
    ENDIF 

    dimdsp(2,2) = dimdsp(2,2) + 1
    dimdsp(2,3) = dimdsp(2,3) + s_all
    IF EMPTY(k_u2)
     dimdsp(2,6) = dimdsp(2,6) + 1
     IF !SEEK(m.sn_pol, 'curpaz5')
      INSERT INTO curpaz5 FROM MEMVAR 
      dimdsp(2,7) = dimdsp(2,7) + 1
     ENDIF 
     dimdsp(2,8) = dimdsp(2,8) + s_all
    ELSE 
     dimdsp(2,9) = dimdsp(2,9) + 1
     IF !SEEK(m.sn_pol, 'curpaz6')
      INSERT INTO curpaz6 FROM MEMVAR 
      dimdsp(2,10) = dimdsp(2,10) + 1
     ENDIF 
     dimdsp(2,11) = dimdsp(2,11) + s_all2
    ENDIF 
    IF EMPTY(er)
     dimdsp(2,13) = dimdsp(2,13) + 1
     dimdsp(2,14) = dimdsp(2,14) + s_all
     IF EMPTY(k_u2)
      dimdsp(2,15) = dimdsp(2,15) + 1
      IF !SEEK(m.sn_pol, 'curpaz7')
       INSERT INTO curpaz7 FROM MEMVAR 
       dimdsp(2,16) = dimdsp(2,16) + 1
      ENDIF 
      dimdsp(2,17) = dimdsp(2,17) + s_all
     ELSE 
      dimdsp(2,18) = dimdsp(2,18) + 1
      IF !SEEK(m.sn_pol, 'curpaz8')
       INSERT INTO curpaz8 FROM MEMVAR 
       dimdsp(2,19) = dimdsp(2,19) + 1
      ENDIF 
*      dimdsp(2,19) = dimdsp(2,19) + k_u2
      dimdsp(2,20) = dimdsp(2,20) + s_all2
     ENDIF 
     DO CASE 
      CASE rslt=396
       dimdsp(2,21) = dimdsp(2,21) + 1
      CASE rslt=337
       dimdsp(2,22) = dimdsp(2,22) + 1
      CASE rslt=338
       dimdsp(2,23) = dimdsp(2,23) + 1
      CASE rslt=339
       dimdsp(2,24) = dimdsp(2,24) + 1
      CASE rslt=340
       dimdsp(2,25) = dimdsp(2,25) + 1
      CASE rslt=341
       dimdsp(2,26) = dimdsp(2,26) + 1
      OTHERWISE 
       dimdsp(2,22) = dimdsp(2,22) + 1
     ENDCASE 
    ENDIF 

   CASE m.tipofcod = 6
    IF !SEEK(m.mcod, 'curlpu3')
     INSERT INTO curlpu3 (mcod) VALUES (m.mcod)
    ENDIF 
    IF !INLIST(rslt,327,328,329,330,331,342)
     LOOP 
    ENDIF 

    dimdsp(3,2) = dimdsp(3,2) + 1
    dimdsp(3,3) = dimdsp(3,3) + s_all
    IF EMPTY(k_u2)
     dimdsp(3,6) = dimdsp(3,6) + 1
     IF !SEEK(m.sn_pol, 'curpaz9')
      INSERT INTO curpaz9 FROM MEMVAR 
      dimdsp(3,7) = dimdsp(3,7) + 1
     ENDIF 
     dimdsp(3,8) = dimdsp(3,8) + s_all
    ELSE 
     dimdsp(3,9) = dimdsp(3,9) + 1
     IF !SEEK(m.sn_pol, 'curpaz10')
      INSERT INTO curpaz10 FROM MEMVAR 
      dimdsp(3,10) = dimdsp(3,10) + 1
     ENDIF 
     dimdsp(3,11) = dimdsp(3,11) + s_all2
    ENDIF 
    IF EMPTY(er)
     dimdsp(3,13) = dimdsp(3,13) + 1
     dimdsp(3,14) = dimdsp(3,14) + s_all
     IF EMPTY(k_u2)
      dimdsp(3,15) = dimdsp(3,15) + 1
      IF !SEEK(m.sn_pol, 'curpaz11')
       INSERT INTO curpaz11 FROM MEMVAR 
       dimdsp(3,16) = dimdsp(3,16) + 1
      ENDIF 
      dimdsp(3,17) = dimdsp(3,17) + s_all
     ELSE 
      dimdsp(3,18) = dimdsp(3,18) + 1
      IF !SEEK(m.sn_pol, 'curpaz12')
       INSERT INTO curpaz12 FROM MEMVAR 
       dimdsp(3,19) = dimdsp(3,19) + 1
      ENDIF 
*      dimdsp(3,19) = dimdsp(3,19) + k_u2
      dimdsp(3,20) = dimdsp(3,20) + s_all2
     ENDIF 
     DO CASE && Неправильно! Только один результат - 342, так как группы здоровью не устанавливаются!
      CASE rslt=327
       dimdsp(3,22) = dimdsp(3,22) + 1
      CASE rslt=328
       dimdsp(3,23) = dimdsp(3,23) + 1
      CASE rslt=329
       dimdsp(3,24) = dimdsp(3,24) + 1
      CASE rslt=330
       dimdsp(3,25) = dimdsp(3,25) + 1
      CASE rslt=331
       dimdsp(3,26) = dimdsp(3,26) + 1
      OTHERWISE 
       dimdsp(3,22) = dimdsp(3,22) + 1
     ENDCASE 
    ENDIF 

   OTHERWISE 

  ENDCASE 
  
  IF !SEEK(m.sn_pol, 'curdsp')
  	INSERT INTO curdsp (sn_pol) VALUES (m.sn_pol)
  ENDIF 
  
 ENDSCAN 

 oExcel.Sheets(1).Select
 WITH oExcel
  .Cells(2,1).Value   = m.qname
  .Cells(5,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'

  .Cells(12,2).Value  = TRANSFORM(RECCOUNT('curlpu1'), '99999')
  .Cells(12,3).Value  = TRANSFORM(dimdsp(1,6)+dimdsp(1,9), '999999')
  .Cells(12,4).Value  = TRANSFORM(dimdsp(1,7)+dimdsp(1,10), '999999')
  .Cells(12,5).Value  = TRANSFORM(dimdsp(1,8)+dimdsp(1,11), '99999999.99')
  .Cells(12,6).Value  = TRANSFORM(dimdsp(1,6), '99999')
  .Cells(12,7).Value  = TRANSFORM(dimdsp(1,7), '99999')
  .Cells(12,8).Value  = TRANSFORM(dimdsp(1,8), '99999999.99')
  .Cells(12,9).Value  = TRANSFORM(dimdsp(1,9), '99999')
  .Cells(12,10).Value = TRANSFORM(dimdsp(1,10), '99999')
  .Cells(12,11).Value = TRANSFORM(dimdsp(1,11), '99999999.99')
  .Cells(12,12).Value = TRANSFORM(dimdsp(1,15)+dimdsp(1,18), '99999')
  .Cells(12,13).Value = TRANSFORM(dimdsp(1,16)+dimdsp(1,19), '99999')
  .Cells(12,14).Value = TRANSFORM(dimdsp(1,17)+dimdsp(1,20), '99999999.99')
  .Cells(12,15).Value = TRANSFORM(dimdsp(1,15), '99999')
  .Cells(12,16).Value  = TRANSFORM(dimdsp(1,16), '99999')
  .Cells(12,17).Value = TRANSFORM(dimdsp(1,17), '99999999.99')
  .Cells(12,18).Value = TRANSFORM(dimdsp(1,18), '99999')
  .Cells(12,19).Value = TRANSFORM(dimdsp(1,19), '99999')
  .Cells(12,20).Value = TRANSFORM(dimdsp(1,20), '99999999.99')
  .Cells(12,21).Value = TRANSFORM(dimdsp(1,21), '99999')
  .Cells(12,22).Value = TRANSFORM(dimdsp(1,22), '99999')
  .Cells(12,23).Value = TRANSFORM(dimdsp(1,23), '99999')
  .Cells(12,24).Value = TRANSFORM(dimdsp(1,24), '99999')
  .Cells(12,25).Value = TRANSFORM(dimdsp(1,25), '99999')
  .Cells(12,26).Value = TRANSFORM(dimdsp(1,26), '99999')

  .Cells(13,2).Value  = TRANSFORM(RECCOUNT('curlpu2'), '99999')
  .Cells(13,3).Value  = TRANSFORM(dimdsp(2,6)+dimdsp(2,9), '999999')
  .Cells(13,4).Value  = TRANSFORM(dimdsp(2,7)+dimdsp(2,10), '999999')
  .Cells(13,5).Value  = TRANSFORM(dimdsp(2,8)+dimdsp(2,11), '99999999.99')
  .Cells(13,6).Value  = TRANSFORM(dimdsp(2,6), '99999')
  .Cells(13,7).Value  = TRANSFORM(dimdsp(2,7), '99999')
  .Cells(13,8).Value  = TRANSFORM(dimdsp(2,8), '99999999.99')
  .Cells(13,9).Value  = TRANSFORM(dimdsp(2,9), '99999')
  .Cells(13,10).Value  = TRANSFORM(dimdsp(2,10), '99999')
  .Cells(13,11).Value  = TRANSFORM(dimdsp(2,11), '99999999.99')
  .Cells(13,12).Value  = TRANSFORM(dimdsp(2,15)+dimdsp(2,18), '99999')
  .Cells(13,13).Value  = TRANSFORM(dimdsp(2,16)+dimdsp(2,19), '99999')
  .Cells(13,14).Value = TRANSFORM(dimdsp(2,17)+dimdsp(2,20), '99999999.99')
  .Cells(13,15).Value = TRANSFORM(dimdsp(2,15), '99999')
  .Cells(13,16).Value = TRANSFORM(dimdsp(2,16), '99999')
  .Cells(13,17).Value = TRANSFORM(dimdsp(2,17), '99999999.99')
  .Cells(13,18).Value = TRANSFORM(dimdsp(2,18), '99999')
  .Cells(13,19).Value = TRANSFORM(dimdsp(2,19), '99999')
  .Cells(13,20).Value = TRANSFORM(dimdsp(2,20), '99999999.99')
  .Cells(13,21).Value = TRANSFORM(dimdsp(2,21), '99999')
  .Cells(13,22).Value = TRANSFORM(dimdsp(2,22), '99999')
  .Cells(13,23).Value = TRANSFORM(dimdsp(2,23), '99999')
  .Cells(13,24).Value = TRANSFORM(dimdsp(2,24), '99999')
  .Cells(13,25).Value = TRANSFORM(dimdsp(2,25), '99999')
  .Cells(13,26).Value = TRANSFORM(dimdsp(2,26), '99999')

  .Cells(14,2).Value  = TRANSFORM(RECCOUNT('curlpu3'), '99999')
  .Cells(14,3).Value  = TRANSFORM(dimdsp(3,6)+dimdsp(3,9), '999999')
  .Cells(14,4).Value  = TRANSFORM(dimdsp(3,7)+dimdsp(3,10), '999999')
  .Cells(14,5).Value  = TRANSFORM(dimdsp(3,8)+dimdsp(3,11), '99999999.99')
  .Cells(14,6).Value  = TRANSFORM(dimdsp(3,6), '99999')
  .Cells(14,7).Value  = TRANSFORM(dimdsp(3,7), '99999')
  .Cells(14,8).Value  = TRANSFORM(dimdsp(3,8), '99999999.99')
  .Cells(14,9).Value  = TRANSFORM(dimdsp(3,9), '99999')
  .Cells(14,10).Value  = TRANSFORM(dimdsp(3,10), '99999')
  .Cells(14,11).Value  = TRANSFORM(dimdsp(3,11), '99999999.99')
  .Cells(14,12).Value  = TRANSFORM(dimdsp(3,15)+dimdsp(3,18), '99999')
  .Cells(14,13).Value  = TRANSFORM(dimdsp(3,16)+dimdsp(3,19), '99999')
  .Cells(14,14).Value = TRANSFORM(dimdsp(3,17)+dimdsp(3,20), '99999999.99')
  .Cells(14,15).Value = TRANSFORM(dimdsp(3,15), '99999')
  .Cells(14,16).Value = TRANSFORM(dimdsp(3,16), '99999')
  .Cells(14,17).Value = TRANSFORM(dimdsp(3,17), '99999999.99')
  .Cells(14,18).Value = TRANSFORM(dimdsp(3,18), '99999')
  .Cells(14,19).Value = TRANSFORM(dimdsp(3,19), '99999')
  .Cells(14,20).Value = TRANSFORM(dimdsp(3,20), '99999999.99')
  .Cells(14,22).Value = TRANSFORM(dimdsp(3,22), '99999')
  .Cells(14,23).Value = TRANSFORM(dimdsp(3,23), '99999')
  .Cells(14,24).Value = TRANSFORM(dimdsp(3,24), '99999')
  .Cells(14,25).Value = TRANSFORM(dimdsp(3,25), '99999')
  .Cells(14,26).Value = TRANSFORM(dimdsp(3,26), '99999')
 ENDWITH 
 
 USE IN curdsp
 USE IN curlpu1
 USE IN curlpu2
 USE IN curlpu3
 
 USE IN curpaz1
 USE IN curpaz2
 USE IN curpaz3
 USE IN curpaz4
 USE IN curpaz5
 USE IN curpaz6
 USE IN curpaz7
 USE IN curpaz8
 USE IN curpaz9
 USE IN curpaz10
 USE IN curpaz11
 USE IN curpaz12

 WAIT CLEAR 
RETURN 


FUNCTION MakePage2Old
 WAIT "Формирование листа 2" WINDOW NOWAIT 
 DIMENSION dimtb4(9,20)
 dimtb4 = 0
 
 SELECT dsp
 
 SCAN 

  m.cod = cod 
  m.rslt = rslt
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=2
   LOOP 
  ENDIF 
*  IF !INLIST(m.cod,15001,1017,1018,1027)
*   LOOP 
*  ENDIF 
  IF !INLIST(m.rslt,343,344,345)
*   LOOP 
  ENDIF 

  m.vozr = ROUND((m.tdat1 - dr)/365.25,2)

  IF m.vozr>=18 AND w=1
   dimtb4(2,1) = dimtb4(2,1) + 1
   dimtb4(2,2) = dimtb4(2,2) + s_all
   IF EMPTY(er)
    dimtb4(2,3) = dimtb4(2,3) + 1
    dimtb4(2,4) = dimtb4(2,4) + s_all
   ENDIF
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(2,5) = dimtb4(2,5) + 1
     dimtb4(2,6) = dimtb4(2,6) + s_all
     IF EMPTY(er)
      dimtb4(2,7) = dimtb4(2,7) + 1
      dimtb4(2,8) = dimtb4(2,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(2,9) = dimtb4(2,9) + 1
     dimtb4(2,10) = dimtb4(2,10) + s_all
     IF EMPTY(er)
      dimtb4(2,11) = dimtb4(2,11) + 1
      dimtb4(2,12) = dimtb4(2,12) + s_all
     ENDIF
    CASE m.rslt = 345
     dimtb4(2,13) = dimtb4(2,13) + 1
     dimtb4(2,14) = dimtb4(2,14) + s_all
     IF EMPTY(er)
      dimtb4(2,15) = dimtb4(2,15) + 1
      dimtb4(2,16) = dimtb4(2,16) + s_all
     ENDIF
   ENDCASE 
  ENDIF 

  IF m.vozr>=18 and w=2
   dimtb4(3,1) = dimtb4(3,1) + 1
   dimtb4(3,2) = dimtb4(3,2) + s_all
   IF EMPTY(er)
    dimtb4(3,3) = dimtb4(3,3) + 1
    dimtb4(3,4) = dimtb4(3,4) + s_all
   ENDIF 
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(3,5) = dimtb4(3,5) + 1
     dimtb4(3,6) = dimtb4(3,6) + s_all
     IF EMPTY(er)
      dimtb4(3,7) = dimtb4(3,7) + 1
      dimtb4(3,8) = dimtb4(3,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(3,9) = dimtb4(3,9) + 1
     dimtb4(3,10) = dimtb4(3,10) + s_all
     IF EMPTY(er)
      dimtb4(3,11) = dimtb4(3,11) + 1
      dimtb4(3,12) = dimtb4(3,12) + s_all
     ENDIF
    CASE m.rslt = 345
    dimtb4(3,13) = dimtb4(3,13) + 1
    dimtb4(3,14) = dimtb4(3,14) + s_all
    IF EMPTY(er)
     dimtb4(3,15) = dimtb4(3,15) + 1
     dimtb4(3,16) = dimtb4(3,16) + s_all
    ENDIF
  ENDCASE 
 ENDIF 

 IF BETWEEN(m.vozr,18,60) AND w=1 
  dimtb4(5,1) = dimtb4(5,1) + 1
  dimtb4(5,2) = dimtb4(5,2) + s_all
  IF EMPTY(er)
   dimtb4(5,3) = dimtb4(5,3) + 1
   dimtb4(5,4) = dimtb4(5,4) + s_all
  ENDIF 
  DO CASE 
   CASE m.rslt = 343 
    dimtb4(5,5) = dimtb4(5,5) + 1
    dimtb4(5,6) = dimtb4(5,6) + s_all
    IF EMPTY(er)
     dimtb4(5,7) = dimtb4(5,7) + 1
     dimtb4(5,8) = dimtb4(5,8) + s_all
    ENDIF
   CASE m.rslt = 344
    dimtb4(5,9) = dimtb4(5,9) + 1
    dimtb4(5,10) = dimtb4(5,10) + s_all
    IF EMPTY(er)
     dimtb4(5,11) = dimtb4(5,11) + 1
     dimtb4(5,12) = dimtb4(5,12) + s_all
    ENDIF
   CASE m.rslt = 345
    dimtb4(5,13) = dimtb4(5,13) + 1
    dimtb4(5,14) = dimtb4(5,14) + s_all
    IF EMPTY(er)
     dimtb4(5,15) = dimtb4(5,15) + 1
     dimtb4(5,16) = dimtb4(5,16) + s_all
    ENDIF
  ENDCASE 
 ENDIF 

 IF BETWEEN(m.vozr,18,55) AND w=2
  dimtb4(6,1) = dimtb4(6,1) + 1
  dimtb4(6,2) = dimtb4(6,2) + s_all
  IF EMPTY(er)
   dimtb4(6,3) = dimtb4(6,3) + 1
   dimtb4(6,4) = dimtb4(6,4) + s_all
  ENDIF 
  DO CASE 
   CASE m.rslt = 343 
    dimtb4(6,5) = dimtb4(6,5) + 1
    dimtb4(6,6) = dimtb4(6,6) + s_all
    IF EMPTY(er)
     dimtb4(6,7) = dimtb4(6,7) + 1
     dimtb4(6,8) = dimtb4(6,8) + s_all
    ENDIF
   CASE m.rslt = 344
    dimtb4(6,9) = dimtb4(6,9) + 1
    dimtb4(6,10) = dimtb4(6,10) + s_all
    IF EMPTY(er)
     dimtb4(6,11) = dimtb4(6,11) + 1
     dimtb4(6,12) = dimtb4(6,12) + s_all
    ENDIF
   CASE m.rslt = 345
    dimtb4(6,13) = dimtb4(6,13) + 1
    dimtb4(6,14) = dimtb4(6,14) + s_all
    IF EMPTY(er)
     dimtb4(6,15) = dimtb4(6,15) + 1
     dimtb4(6,16) = dimtb4(6,16) + s_all
    ENDIF
   ENDCASE 
  ENDIF 

     IF m.vozr>60 AND w=1
      dimtb4(8,1) = dimtb4(8,1) + 1
      dimtb4(8,2) = dimtb4(8,2) + s_all
      IF EMPTY(er)
       dimtb4(8,3) = dimtb4(8,3) + 1
       dimtb4(8,4) = dimtb4(8,4) + s_all
      ENDIF 
      DO CASE 
       CASE m.rslt = 343 
        dimtb4(8,5) = dimtb4(8,5) + 1
        dimtb4(8,6) = dimtb4(8,6) + s_all
        IF EMPTY(er)
         dimtb4(8,7) = dimtb4(8,7) + 1
         dimtb4(8,8) = dimtb4(8,8) + s_all
        ENDIF
       CASE m.rslt = 344
        dimtb4(8,9) = dimtb4(8,9) + 1
        dimtb4(8,10) = dimtb4(8,10) + s_all
        IF EMPTY(er)
         dimtb4(8,11) = dimtb4(8,11) + 1
         dimtb4(8,12) = dimtb4(8,12) + s_all
        ENDIF
       CASE m.rslt = 345
        dimtb4(8,13) = dimtb4(8,13) + 1
        dimtb4(8,14) = dimtb4(8,14) + s_all
        IF EMPTY(er)
         dimtb4(8,15) = dimtb4(8,15) + 1
         dimtb4(8,16) = dimtb4(8,16) + s_all
        ENDIF
      ENDCASE 
     ENDIF 

     IF m.vozr>55 and w=2
      dimtb4(9,1) = dimtb4(9,1) + 1
      dimtb4(9,2) = dimtb4(9,2) + s_all
      IF EMPTY(er)
       dimtb4(9,3) = dimtb4(9,3) + 1
       dimtb4(9,4) = dimtb4(9,4) + s_all
      ENDIF 
      DO CASE 
       CASE m.rslt = 343 
        dimtb4(9,5) = dimtb4(9,5) + 1
        dimtb4(9,6) = dimtb4(9,6) + s_all
        IF EMPTY(er)
         dimtb4(9,7) = dimtb4(9,7) + 1
         dimtb4(9,8) = dimtb4(9,8) + s_all
        ENDIF
       CASE m.rslt = 344
        dimtb4(9,9) = dimtb4(9,9) + 1
        dimtb4(9,10) = dimtb4(9,10) + s_all
        IF EMPTY(er)
         dimtb4(9,11) = dimtb4(9,11) + 1
         dimtb4(9,12) = dimtb4(9,12) + s_all
        ENDIF
       CASE m.rslt = 345
        dimtb4(9,13) = dimtb4(9,13) + 1
        dimtb4(9,14) = dimtb4(9,14) + s_all
        IF EMPTY(er)
         dimtb4(9,15) = dimtb4(9,15) + 1
         dimtb4(9,16) = dimtb4(9,16) + s_all
        ENDIF
      ENDCASE 
     ENDIF 

  ENDSCAN 

 oExcel.Sheets(2).Select
 WITH oExcel
  .Cells(2,1).Value   = m.qname
  .Cells(5,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
  .Cells(11,3).Value = TRANSFORM(dimtb4(2,5)+dimtb4(3,5), '99999')
  .Cells(11,4).Value = TRANSFORM(dimtb4(2,9)+dimtb4(3,9), '99999')
  .Cells(11,5).Value = TRANSFORM(dimtb4(2,13)+dimtb4(3,13), '99999')
  .Cells(11,6).Value = TRANSFORM(dimtb4(2,1)+dimtb4(3,1), '99999')
  .Cells(11,7).Value  = TRANSFORM(dimtb4(2,6)+dimtb4(3,6), '9999999.99')
  .Cells(11,8).Value  = TRANSFORM(dimtb4(2,10)+dimtb4(3,10), '9999999.99')
  .Cells(11,9).Value  = TRANSFORM(dimtb4(2,14)+dimtb4(3,14), '9999999.99')
  .Cells(11,10).Value = TRANSFORM(dimtb4(2,2)+dimtb4(3,2), '9999999.99')
  .Cells(11,11).Value = TRANSFORM(dimtb4(2,7)+dimtb4(3,7), '99999')
  .Cells(11,12).Value = TRANSFORM(dimtb4(2,11)+dimtb4(3,11), '99999')
  .Cells(11,13).Value = TRANSFORM(dimtb4(2,15)+dimtb4(3,15), '99999')
  .Cells(11,14).Value = TRANSFORM(dimtb4(2,3)+dimtb4(3,3), '99999')
  .Cells(11,15).Value = TRANSFORM(dimtb4(2,8)+dimtb4(3,8), '9999999.99')
  .Cells(11,16).Value = TRANSFORM(dimtb4(2,12)+dimtb4(3,12), '9999999.99')
  .Cells(11,17).Value = TRANSFORM(dimtb4(2,16)+dimtb4(3,16), '9999999.99')
  .Cells(11,18).Value = TRANSFORM(dimtb4(2,4)+dimtb4(3,4), '9999999.99')


  .Cells(12,3).Value  = TRANSFORM(dimtb4(2,5), '99999')
  .Cells(12,4).Value  = TRANSFORM(dimtb4(2,9), '99999')
  .Cells(12,5).Value  = TRANSFORM(dimtb4(2,13), '99999')
  .Cells(12,6).Value  = TRANSFORM(dimtb4(2,1), '99999')
  .Cells(12,7).Value  = TRANSFORM(dimtb4(2,6), '9999999.99')
  .Cells(12,8).Value  = TRANSFORM(dimtb4(2,10), '9999999.99')
  .Cells(12,9).Value  = TRANSFORM(dimtb4(2,14), '9999999.99')
  .Cells(12,10).Value = TRANSFORM(dimtb4(2,2), '9999999.99')
  .Cells(12,11).Value = TRANSFORM(dimtb4(2,7), '99999')
  .Cells(12,12).Value = TRANSFORM(dimtb4(2,11), '99999')
  .Cells(12,13).Value = TRANSFORM(dimtb4(2,15), '99999')
  .Cells(12,14).Value = TRANSFORM(dimtb4(2,3), '99999')
  .Cells(12,15).Value = TRANSFORM(dimtb4(2,8), '9999999.99')
  .Cells(12,16).Value = TRANSFORM(dimtb4(2,12), '9999999.99')
  .Cells(12,17).Value = TRANSFORM(dimtb4(2,16), '9999999.99')
  .Cells(12,18).Value = TRANSFORM(dimtb4(2,4), '9999999.99')


  .Cells(13,3).Value  = TRANSFORM(dimtb4(3,5), '99999')
  .Cells(13,4).Value  = TRANSFORM(dimtb4(3,9), '99999')
  .Cells(13,5).Value  = TRANSFORM(dimtb4(3,13), '99999')
  .Cells(13,6).Value = TRANSFORM(dimtb4(3,1), '99999')
  .Cells(13,7).Value  = TRANSFORM(dimtb4(3,6), '9999999.99')
  .Cells(13,8).Value  = TRANSFORM(dimtb4(3,10), '9999999.99')
  .Cells(13,9).Value  = TRANSFORM(dimtb4(3,14), '9999999.99')
  .Cells(13,10).Value = TRANSFORM(dimtb4(3,2), '9999999.99')
  .Cells(13,11).Value = TRANSFORM(dimtb4(3,7), '99999')
  .Cells(13,12).Value = TRANSFORM(dimtb4(3,11), '99999')
  .Cells(13,13).Value = TRANSFORM(dimtb4(3,15), '99999')
  .Cells(13,14).Value = TRANSFORM(dimtb4(3,3), '99999')
  .Cells(13,15).Value = TRANSFORM(dimtb4(3,8), '9999999.99')
  .Cells(13,16).Value = TRANSFORM(dimtb4(3,12), '9999999.99')
  .Cells(13,17).Value = TRANSFORM(dimtb4(3,16), '9999999.99')
  .Cells(13,18).Value = TRANSFORM(dimtb4(3,4), '9999999.99')


  .Cells(14,3).Value  = TRANSFORM(dimtb4(5,5)+dimtb4(6,5), '99999')
  .Cells(14,4).Value  = TRANSFORM(dimtb4(5,9)+dimtb4(6,9), '99999')
  .Cells(14,5).Value  = TRANSFORM(dimtb4(5,13)+dimtb4(6,13), '99999')
  .Cells(14,6).Value = TRANSFORM(dimtb4(5,1)+dimtb4(6,1), '99999')
  .Cells(14,7).Value  = TRANSFORM(dimtb4(5,6)+dimtb4(6,6), '9999999.99')
  .Cells(14,8).Value  = TRANSFORM(dimtb4(5,10)+dimtb4(6,10), '9999999.99')
  .Cells(14,9).Value  = TRANSFORM(dimtb4(5,14)+dimtb4(6,14), '9999999.99')
  .Cells(14,10).Value = TRANSFORM(dimtb4(5,2)+dimtb4(6,2), '9999999.99')
  .Cells(14,11).Value = TRANSFORM(dimtb4(5,7)+dimtb4(6,7), '99999')
  .Cells(14,12).Value = TRANSFORM(dimtb4(5,11)+dimtb4(6,11), '99999')
  .Cells(14,13).Value = TRANSFORM(dimtb4(5,15)+dimtb4(6,15), '99999')
  .Cells(14,14).Value = TRANSFORM(dimtb4(5,3)+dimtb4(6,3), '99999')
  .Cells(14,15).Value = TRANSFORM(dimtb4(5,8)+dimtb4(6,8), '9999999.99')
  .Cells(14,16).Value = TRANSFORM(dimtb4(5,12)+dimtb4(6,12), '9999999.99')
  .Cells(14,17).Value = TRANSFORM(dimtb4(5,16)+dimtb4(6,16), '9999999.99')
  .Cells(14,18).Value = TRANSFORM(dimtb4(5,4)+dimtb4(6,4), '9999999.99')

  .Cells(15,3).Value  = TRANSFORM(dimtb4(5,5), '99999')
  .Cells(15,4).Value  = TRANSFORM(dimtb4(5,9), '99999')
  .Cells(15,5).Value  = TRANSFORM(dimtb4(5,13), '99999')
  .Cells(15,6).Value  = TRANSFORM(dimtb4(5,1), '99999')
  .Cells(15,7).Value  = TRANSFORM(dimtb4(5,6), '9999999.99')
  .Cells(15,8).Value  = TRANSFORM(dimtb4(5,10), '9999999.99')
  .Cells(15,9).Value  = TRANSFORM(dimtb4(5,14), '9999999.99')
  .Cells(15,10).Value = TRANSFORM(dimtb4(5,2), '9999999.99')
  .Cells(15,11).Value = TRANSFORM(dimtb4(5,7), '99999')
  .Cells(15,12).Value = TRANSFORM(dimtb4(5,11), '99999')
  .Cells(15,13).Value = TRANSFORM(dimtb4(5,15), '99999')
  .Cells(15,14).Value = TRANSFORM(dimtb4(5,3), '99999')
  .Cells(15,15).Value = TRANSFORM(dimtb4(5,8), '9999999.99')
  .Cells(15,16).Value = TRANSFORM(dimtb4(5,12), '9999999.99')
  .Cells(15,17).Value = TRANSFORM(dimtb4(5,16), '9999999.99')
  .Cells(15,18).Value = TRANSFORM(dimtb4(5,4), '9999999.99')

  .Cells(16,3).Value  = TRANSFORM(dimtb4(6,5), '99999')
  .Cells(16,4).Value  = TRANSFORM(dimtb4(6,9), '99999')
  .Cells(16,5).Value  = TRANSFORM(dimtb4(6,13), '99999')
  .Cells(16,6).Value = TRANSFORM(dimtb4(6,1), '99999')
  .Cells(16,7).Value  = TRANSFORM(dimtb4(6,6), '9999999.99')
  .Cells(16,8).Value  = TRANSFORM(dimtb4(6,10), '9999999.99')
  .Cells(16,9).Value  = TRANSFORM(dimtb4(6,14), '9999999.99')
  .Cells(16,10).Value = TRANSFORM(dimtb4(6,2), '9999999.99')
  .Cells(16,11).Value = TRANSFORM(dimtb4(6,7), '99999')
  .Cells(16,12).Value = TRANSFORM(dimtb4(6,11), '99999')
  .Cells(16,13).Value = TRANSFORM(dimtb4(6,15), '99999')
  .Cells(16,14).Value = TRANSFORM(dimtb4(6,3), '99999')
  .Cells(16,15).Value = TRANSFORM(dimtb4(6,8), '9999999.99')
  .Cells(16,16).Value = TRANSFORM(dimtb4(6,12), '9999999.99')
  .Cells(16,17).Value = TRANSFORM(dimtb4(6,16), '9999999.99')
  .Cells(16,18).Value = TRANSFORM(dimtb4(6,4), '9999999.99')

  .Cells(17,3).Value  = TRANSFORM(dimtb4(8,5)+dimtb4(9,5), '99999')
  .Cells(17,4).Value  = TRANSFORM(dimtb4(8,9)+dimtb4(9,9), '99999')
  .Cells(17,5).Value  = TRANSFORM(dimtb4(8,13)+dimtb4(9,13), '99999')
  .Cells(17,6).Value = TRANSFORM(dimtb4(8,1)+dimtb4(9,1), '99999')
  .Cells(17,7).Value  = TRANSFORM(dimtb4(8,6)+dimtb4(9,6), '9999999.99')
  .Cells(17,8).Value  = TRANSFORM(dimtb4(8,10)+dimtb4(9,10), '9999999.99')
  .Cells(17,9).Value  = TRANSFORM(dimtb4(8,14)+dimtb4(9,14), '9999999.99')
  .Cells(17,10).Value = TRANSFORM(dimtb4(8,2)+dimtb4(9,2), '9999999.99')
  .Cells(17,11).Value = TRANSFORM(dimtb4(8,7)+dimtb4(9,7), '99999')
  .Cells(17,12).Value = TRANSFORM(dimtb4(8,11)+dimtb4(9,11), '99999')
  .Cells(17,13).Value = TRANSFORM(dimtb4(8,15)+dimtb4(9,15), '99999')
  .Cells(17,14).Value = TRANSFORM(dimtb4(8,3)+dimtb4(9,3), '99999')
  .Cells(17,15).Value = TRANSFORM(dimtb4(8,8)+dimtb4(9,8), '9999999.99')
  .Cells(17,16).Value = TRANSFORM(dimtb4(8,12)+dimtb4(9,12), '9999999.99')
  .Cells(17,17).Value = TRANSFORM(dimtb4(8,16)+dimtb4(9,16), '9999999.99')
  .Cells(17,18).Value = TRANSFORM(dimtb4(8,4)+dimtb4(9,4), '9999999.99')

  .Cells(18,3).Value  = TRANSFORM(dimtb4(8,5), '99999')
  .Cells(18,4).Value  = TRANSFORM(dimtb4(8,9), '99999')
  .Cells(18,5).Value  = TRANSFORM(dimtb4(8,13), '99999')
  .Cells(18,6).Value = TRANSFORM(dimtb4(8,1), '99999')
  .Cells(18,7).Value  = TRANSFORM(dimtb4(8,6), '9999999.99')
  .Cells(18,8).Value  = TRANSFORM(dimtb4(8,10), '9999999.99')
  .Cells(18,9).Value  = TRANSFORM(dimtb4(8,14), '9999999.99')
  .Cells(18,10).Value = TRANSFORM(dimtb4(8,2), '9999999.99')
  .Cells(18,11).Value = TRANSFORM(dimtb4(8,7), '99999')
  .Cells(18,12).Value = TRANSFORM(dimtb4(8,11), '99999')
  .Cells(18,13).Value = TRANSFORM(dimtb4(8,15), '99999')
  .Cells(18,14).Value = TRANSFORM(dimtb4(8,3), '99999')
  .Cells(18,15).Value = TRANSFORM(dimtb4(8,8), '9999999.99')
  .Cells(18,16).Value = TRANSFORM(dimtb4(8,12), '9999999.99')
  .Cells(18,17).Value = TRANSFORM(dimtb4(8,16), '9999999.99')
  .Cells(18,18).Value = TRANSFORM(dimtb4(8,4), '9999999.99')

  .Cells(19,3).Value  = TRANSFORM(dimtb4(9,5), '99999')
  .Cells(19,4).Value  = TRANSFORM(dimtb4(9,9), '99999')
  .Cells(19,5).Value  = TRANSFORM(dimtb4(9,13), '99999')
  .Cells(19,6).Value = TRANSFORM(dimtb4(9,1), '99999')
  .Cells(19,7).Value  = TRANSFORM(dimtb4(9,6), '9999999.99')
  .Cells(19,8).Value  = TRANSFORM(dimtb4(9,10), '9999999.99')
  .Cells(19,9).Value  = TRANSFORM(dimtb4(9,14), '9999999.99')
  .Cells(19,10).Value = TRANSFORM(dimtb4(9,2), '9999999.99')
  .Cells(19,11).Value = TRANSFORM(dimtb4(9,7), '99999')
  .Cells(19,12).Value = TRANSFORM(dimtb4(9,11), '99999')
  .Cells(19,13).Value = TRANSFORM(dimtb4(9,15), '99999')
  .Cells(19,14).Value = TRANSFORM(dimtb4(9,3), '99999')
  .Cells(19,15).Value = TRANSFORM(dimtb4(9,8), '9999999.99')
  .Cells(19,16).Value = TRANSFORM(dimtb4(9,12), '9999999.99')
  .Cells(19,17).Value = TRANSFORM(dimtb4(9,16), '9999999.99')
  .Cells(19,18).Value = TRANSFORM(dimtb4(9,4), '9999999.99')

 ENDWITH 
 
 WAIT CLEAR 
RETURN 

FUNCTION MakePage2
 WAIT "Формирование листа 2" WINDOW NOWAIT 
 DIMENSION dimtb4(9,20)
 dimtb4 = 0
 
 SELECT dsp
 
 SCAN 

  m.cod = cod 
  m.rslt = rslt
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=2
   LOOP 
  ENDIF 
*  IF !INLIST(m.cod,15001,1017,1018,1027)
*   LOOP 
*  ENDIF 
  IF !INLIST(m.rslt,343,344,345)
*   LOOP 
  ENDIF 

  m.vozr = ROUND((m.tdat1 - dr)/365.25,2)

  IF m.vozr>=18 AND w=1
   dimtb4(2,1) = dimtb4(2,1) + 1
   dimtb4(2,2) = dimtb4(2,2) + s_all
   IF EMPTY(er)
    dimtb4(2,3) = dimtb4(2,3) + 1
    dimtb4(2,4) = dimtb4(2,4) + s_all
   ENDIF
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(2,5) = dimtb4(2,5) + 1
     dimtb4(2,6) = dimtb4(2,6) + s_all
     IF EMPTY(er)
      dimtb4(2,7) = dimtb4(2,7) + 1
      dimtb4(2,8) = dimtb4(2,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(2,9) = dimtb4(2,9) + 1
     dimtb4(2,10) = dimtb4(2,10) + s_all
     IF EMPTY(er)
      dimtb4(2,11) = dimtb4(2,11) + 1
      dimtb4(2,12) = dimtb4(2,12) + s_all
     ENDIF
    CASE m.rslt = 345
     dimtb4(2,13) = dimtb4(2,13) + 1
     dimtb4(2,14) = dimtb4(2,14) + s_all
     IF EMPTY(er)
      dimtb4(2,15) = dimtb4(2,15) + 1
      dimtb4(2,16) = dimtb4(2,16) + s_all
     ENDIF
   ENDCASE 
  ENDIF 

  IF m.vozr>=18 and w=2
   dimtb4(3,1) = dimtb4(3,1) + 1
   dimtb4(3,2) = dimtb4(3,2) + s_all
   IF EMPTY(er)
    dimtb4(3,3) = dimtb4(3,3) + 1
    dimtb4(3,4) = dimtb4(3,4) + s_all
   ENDIF 
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(3,5) = dimtb4(3,5) + 1
     dimtb4(3,6) = dimtb4(3,6) + s_all
     IF EMPTY(er)
      dimtb4(3,7) = dimtb4(3,7) + 1
      dimtb4(3,8) = dimtb4(3,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(3,9) = dimtb4(3,9) + 1
     dimtb4(3,10) = dimtb4(3,10) + s_all
     IF EMPTY(er)
      dimtb4(3,11) = dimtb4(3,11) + 1
      dimtb4(3,12) = dimtb4(3,12) + s_all
     ENDIF
    CASE m.rslt = 345
    dimtb4(3,13) = dimtb4(3,13) + 1
    dimtb4(3,14) = dimtb4(3,14) + s_all
    IF EMPTY(er)
     dimtb4(3,15) = dimtb4(3,15) + 1
     dimtb4(3,16) = dimtb4(3,16) + s_all
    ENDIF
  ENDCASE 
 ENDIF 

 IF BETWEEN(m.vozr,18,60) AND w=1 
  dimtb4(5,1) = dimtb4(5,1) + 1
  dimtb4(5,2) = dimtb4(5,2) + s_all
  IF EMPTY(er)
   dimtb4(5,3) = dimtb4(5,3) + 1
   dimtb4(5,4) = dimtb4(5,4) + s_all
  ENDIF 
  DO CASE 
   CASE m.rslt = 343 
    dimtb4(5,5) = dimtb4(5,5) + 1
    dimtb4(5,6) = dimtb4(5,6) + s_all
    IF EMPTY(er)
     dimtb4(5,7) = dimtb4(5,7) + 1
     dimtb4(5,8) = dimtb4(5,8) + s_all
    ENDIF
   CASE m.rslt = 344
    dimtb4(5,9) = dimtb4(5,9) + 1
    dimtb4(5,10) = dimtb4(5,10) + s_all
    IF EMPTY(er)
     dimtb4(5,11) = dimtb4(5,11) + 1
     dimtb4(5,12) = dimtb4(5,12) + s_all
    ENDIF
   CASE m.rslt = 345
    dimtb4(5,13) = dimtb4(5,13) + 1
    dimtb4(5,14) = dimtb4(5,14) + s_all
    IF EMPTY(er)
     dimtb4(5,15) = dimtb4(5,15) + 1
     dimtb4(5,16) = dimtb4(5,16) + s_all
    ENDIF
  ENDCASE 
 ENDIF 

 IF BETWEEN(m.vozr,18,55) AND w=2
  dimtb4(6,1) = dimtb4(6,1) + 1
  dimtb4(6,2) = dimtb4(6,2) + s_all
  IF EMPTY(er)
   dimtb4(6,3) = dimtb4(6,3) + 1
   dimtb4(6,4) = dimtb4(6,4) + s_all
  ENDIF 
  DO CASE 
   CASE m.rslt = 343 
    dimtb4(6,5) = dimtb4(6,5) + 1
    dimtb4(6,6) = dimtb4(6,6) + s_all
    IF EMPTY(er)
     dimtb4(6,7) = dimtb4(6,7) + 1
     dimtb4(6,8) = dimtb4(6,8) + s_all
    ENDIF
   CASE m.rslt = 344
    dimtb4(6,9) = dimtb4(6,9) + 1
    dimtb4(6,10) = dimtb4(6,10) + s_all
    IF EMPTY(er)
     dimtb4(6,11) = dimtb4(6,11) + 1
     dimtb4(6,12) = dimtb4(6,12) + s_all
    ENDIF
   CASE m.rslt = 345
    dimtb4(6,13) = dimtb4(6,13) + 1
    dimtb4(6,14) = dimtb4(6,14) + s_all
    IF EMPTY(er)
     dimtb4(6,15) = dimtb4(6,15) + 1
     dimtb4(6,16) = dimtb4(6,16) + s_all
    ENDIF
   ENDCASE 
  ENDIF 

     IF m.vozr>60 AND w=1
      dimtb4(8,1) = dimtb4(8,1) + 1
      dimtb4(8,2) = dimtb4(8,2) + s_all
      IF EMPTY(er)
       dimtb4(8,3) = dimtb4(8,3) + 1
       dimtb4(8,4) = dimtb4(8,4) + s_all
      ENDIF 
      DO CASE 
       CASE m.rslt = 343 
        dimtb4(8,5) = dimtb4(8,5) + 1
        dimtb4(8,6) = dimtb4(8,6) + s_all
        IF EMPTY(er)
         dimtb4(8,7) = dimtb4(8,7) + 1
         dimtb4(8,8) = dimtb4(8,8) + s_all
        ENDIF
       CASE m.rslt = 344
        dimtb4(8,9) = dimtb4(8,9) + 1
        dimtb4(8,10) = dimtb4(8,10) + s_all
        IF EMPTY(er)
         dimtb4(8,11) = dimtb4(8,11) + 1
         dimtb4(8,12) = dimtb4(8,12) + s_all
        ENDIF
       CASE m.rslt = 345
        dimtb4(8,13) = dimtb4(8,13) + 1
        dimtb4(8,14) = dimtb4(8,14) + s_all
        IF EMPTY(er)
         dimtb4(8,15) = dimtb4(8,15) + 1
         dimtb4(8,16) = dimtb4(8,16) + s_all
        ENDIF
      ENDCASE 
     ENDIF 

     IF m.vozr>55 and w=2
      dimtb4(9,1) = dimtb4(9,1) + 1
      dimtb4(9,2) = dimtb4(9,2) + s_all
      IF EMPTY(er)
       dimtb4(9,3) = dimtb4(9,3) + 1
       dimtb4(9,4) = dimtb4(9,4) + s_all
      ENDIF 
      DO CASE 
       CASE m.rslt = 343 
        dimtb4(9,5) = dimtb4(9,5) + 1
        dimtb4(9,6) = dimtb4(9,6) + s_all
        IF EMPTY(er)
         dimtb4(9,7) = dimtb4(9,7) + 1
         dimtb4(9,8) = dimtb4(9,8) + s_all
        ENDIF
       CASE m.rslt = 344
        dimtb4(9,9) = dimtb4(9,9) + 1
        dimtb4(9,10) = dimtb4(9,10) + s_all
        IF EMPTY(er)
         dimtb4(9,11) = dimtb4(9,11) + 1
         dimtb4(9,12) = dimtb4(9,12) + s_all
        ENDIF
       CASE m.rslt = 345
        dimtb4(9,13) = dimtb4(9,13) + 1
        dimtb4(9,14) = dimtb4(9,14) + s_all
        IF EMPTY(er)
         dimtb4(9,15) = dimtb4(9,15) + 1
         dimtb4(9,16) = dimtb4(9,16) + s_all
        ENDIF
      ENDCASE 
     ENDIF 

  ENDSCAN 

 oExcel.Sheets(2).Select
 WITH oExcel
  .Cells(2,1).Value   = m.qname
  .Cells(5,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'

  .Cells(10,7).Value  = TRANSFORM(dimtb4(2,1)+dimtb4(3,1), '99999') && !
*  .Cells(10,3).Value = TRANSFORM(ROUND((dimtb4(2,2)+dimtb4(3,2))/1000,2), '9999999.99') && !
  .Cells(10,3).Value = TRANSFORM(dimtb4(2,2)+dimtb4(3,2), '9999999.99') && !
*  .Cells(10,11).Value = TRANSFORM(ROUND((dimtb4(2,4)+dimtb4(3,4))/1000,2), '9999999.99') && !
  .Cells(10,11).Value = TRANSFORM(dimtb4(2,4)+dimtb4(3,4), '9999999.99') && !

  .Cells(11,7).Value  = TRANSFORM(dimtb4(2,1), '99999')
*  .Cells(11,3).Value = TRANSFORM(ROUND(dimtb4(2,2)/1000,2), '9999999.99')
  .Cells(11,3).Value = TRANSFORM(dimtb4(2,2), '9999999.99')
*  .Cells(11,11).Value = TRANSFORM(ROUND(dimtb4(2,4)/1000,2), '9999999.99')
  .Cells(11,11).Value = TRANSFORM(dimtb4(2,4), '9999999.99')

  .Cells(12,7).Value = TRANSFORM(dimtb4(3,1), '99999')
*  .Cells(12,3).Value = TRANSFORM(ROUND(dimtb4(3,2)/1000,2), '9999999.99')
  .Cells(12,3).Value = TRANSFORM(dimtb4(3,2), '9999999.99')
*  .Cells(12,11).Value = TRANSFORM(ROUND(dimtb4(3,4)/1000,2), '9999999.99')
  .Cells(12,11).Value = TRANSFORM(dimtb4(3,4), '9999999.99')

 ENDWITH 
 
 WAIT CLEAR 
RETURN 

FUNCTION MakePage3
 WAIT "Формирование листа 3" WINDOW NOWAIT 

 CREATE CURSOR curdsp (sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curlpu1 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 CREATE CURSOR curlpu2 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 

 CREATE CURSOR curlpu3 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 

 SELECT dsp
 
 DIMENSION dimdsp(9,17)
 dimdsp = 0

 SCAN 
  m.cod    = cod
  m.w      = w
  m.ages   = ages
  m.sn_pol = sn_pol
  m.mcod   = mcod

  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=1
   LOOP 
  ENDIF 
  IF m.ages<18
   LOOP 
  ENDIF 
  
  IF !SEEK(m.mcod, 'curlpu1')
   INSERT INTO curlpu1 (mcod) VALUES (m.mcod)
  ENDIF 

  =incdimdsp(1)
  IF m.w=1 && вторая строка, мужчины
   =incdimdsp(2)
  ENDIF 
  IF m.w=2 && вторая строка, женщины
   =incdimdsp(3)
  ENDIF 
  IF (m.w=1 AND BETWEEN(m.ages, 18, 60)) OR (m.w=2 AND BETWEEN(m.ages, 18, 55))
   =incdimdsp(4)
  ENDIF 
  IF m.w=1 AND BETWEEN(m.ages, 18, 60) && пятая строка, мужчины
   =incdimdsp(5)
   IF !SEEK(m.mcod, 'curlpu2')
    INSERT INTO curlpu2 (mcod) VALUES (m.mcod)
   ENDIF 
  ENDIF 
  IF m.w=2 AND BETWEEN(m.ages, 18, 55) && шестая строка, женщины
   =incdimdsp(6)
   IF !SEEK(m.mcod, 'curlpu2')
    INSERT INTO curlpu2 (mcod) VALUES (m.mcod)
   ENDIF 
  ENDIF 
  IF (m.w=1 AND m.ages>60) OR (m.w=2 AND m.ages>55)
   =incdimdsp(7)
  ENDIF 
  IF m.w=1 and m.ages>60 && восьмая строка, мужчины
   IF !SEEK(m.mcod, 'curlpu3')
    INSERT INTO curlpu3 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(8)
  ENDIF 
  IF m.w=2 AND m.ages>55 && девятая строка, женщины
   IF !SEEK(m.mcod, 'curlpu3')
    INSERT INTO curlpu3 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(9)
  ENDIF 

  IF !SEEK(m.sn_pol, 'curdsp')
  	INSERT INTO curdsp (sn_pol) VALUES (m.sn_pol)
  ENDIF 

 ENDSCAN 

 oExcel.Sheets(3).Select
 oExcel.Cells(6,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
 oExcel.Cells(2,1).Value   = m.qname

 oExcel.Cells(14,3).Value  = TRANSFORM(RECCOUNT('curlpu1'), '9999')
 oExcel.Cells(17,3).Value  = TRANSFORM(RECCOUNT('curlpu2'), '9999')
 oExcel.Cells(20,3).Value  = TRANSFORM(RECCOUNT('curlpu3'), '9999')

 FOR m.nstr=1 TO 9
  WITH oExcel
   .Cells(13+m.nstr,5).Value  = TRANSFORM(dimdsp(m.nstr,2)+dimdsp(m.nstr,5), '999999999.99')
   .Cells(13+m.nstr,6).Value  = TRANSFORM(dimdsp(m.nstr,1), '99999')
   .Cells(13+m.nstr,7).Value  = TRANSFORM(dimdsp(m.nstr,2), '999999999.99')
   .Cells(13+m.nstr,8).Value  = TRANSFORM(dimdsp(m.nstr,3), '99999')
   .Cells(13+m.nstr,9).Value  = TRANSFORM(dimdsp(m.nstr,4), '99999')
   .Cells(13+m.nstr,10).Value = TRANSFORM(dimdsp(m.nstr,5), '999999999.99')
   .Cells(13+m.nstr,11).Value = TRANSFORM(dimdsp(m.nstr,7)+dimdsp(m.nstr,11), '999999999.99')
*   .Cells(13+m.nstr,12).Value = TRANSFORM(dimdsp(m.nstr,6), '99999')
   .Cells(13+m.nstr,12).Value = TRANSFORM(dimdsp(m.nstr,6)-dimdsp(m.nstr,9), '99999')
   .Cells(13+m.nstr,13).Value = TRANSFORM(dimdsp(m.nstr,7), '999999999.99')
   .Cells(13+m.nstr,14).Value = TRANSFORM(dimdsp(m.nstr,9), '99999')
   .Cells(13+m.nstr,15).Value = TRANSFORM(dimdsp(m.nstr,10), '99999')
   .Cells(13+m.nstr,16).Value = TRANSFORM(dimdsp(m.nstr,11), '999999999.99')
   .Cells(13+m.nstr,17).Value = TRANSFORM(dimdsp(m.nstr,12), '99999')
*   .Cells(13+m.nstr,18).Value = TRANSFORM(dimdsp(m.nstr,6)-dimdsp(m.nstr,14)-dimdsp(m.nstr,15)-;
     dimdsp(m.nstr,16)-dimdsp(m.nstr,17)-dimdsp(m.nstr,12), '99999')
   .Cells(13+m.nstr,18).Value = TRANSFORM(dimdsp(m.nstr,6)-dimdsp(m.nstr,9)-dimdsp(m.nstr,14)-dimdsp(m.nstr,15)-;
     dimdsp(m.nstr,16)-dimdsp(m.nstr,17), '99999')
   .Cells(13+m.nstr,19).Value = TRANSFORM(dimdsp(m.nstr,14), '99999')
   .Cells(13+m.nstr,20).Value = TRANSFORM(dimdsp(m.nstr,15), '99999')
   .Cells(13+m.nstr,21).Value = TRANSFORM(dimdsp(m.nstr,16), '99999')
   .Cells(13+m.nstr,22).Value = TRANSFORM(dimdsp(m.nstr,17), '99999')
  ENDWITH 
 ENDFOR 
 
 USE IN curdsp
 USE IN curlpu1
 USE IN curlpu2
 USE IN curlpu3

RETURN 

WAIT CLEAR 
RETURN 

FUNCTION MakePage4
 WAIT "Формирование листа 4" WINDOW NOWAIT 

 CREATE CURSOR curlpu1 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 CREATE CURSOR curpols (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 DIMENSION dimtb11(2,15)
 dimtb11 = 0
 
 SELECT dsp
 SCAN 
  m.cod    = cod
  m.mcod   = mcod
  m.sn_pol = sn_pol
  
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=3
   LOOP 
  ENDIF 
  m.rslt=rslt
  m.c_i = c_i
*  IF !INLIST(m.rslt,320,321,322,323,324,325)
  IF !INLIST(m.rslt,321,322,323,324,325)
   LOOP 
  ENDIF 
*  IF m.rslt=320 AND LEFT(m.c_i,2)!='ДС'
*   LOOP 
*  ENDIF 

  IF !SEEK(m.mcod, 'curlpu1')
   INSERT INTO curlpu1 (mcod) VALUES (m.mcod)
  ENDIF 
*  IF !INLIST(cod,101929,101930,101931,101932)
*   LOOP 
*  ENDIF 

  IF EMPTY(k_u2)
   dimtb11(1,7) = dimtb11(1,7) + 1
   dimtb11(1,8) = dimtb11(1,8) + s_all
   IF EMPTY(er)
    dimtb11(2,7) = dimtb11(2,7) + 1
    dimtb11(2,8) = dimtb11(2,8) + s_all
   ENDIF 
  ELSE 
   dimtb11(1,9) = dimtb11(1,9) + 1
   dimtb11(1,10) = dimtb11(1,10) + k_u2
   dimtb11(1,11) = dimtb11(1,11) + s_all2
*   dimtb11(1,11) = dimtb11(1,11) + s_all + s_all2
   IF !SEEK(m.sn_pol, 'curpols')
    INSERT INTO curpols (sn_pol) VALUES (m.sn_pol)
   ENDIF 
   IF EMPTY(er)
    dimtb11(2,9) = dimtb11(2,9) + 1
    dimtb11(2,10) = dimtb11(2,10) + k_u2
    dimtb11(2,11) = dimtb11(2,11) + s_all2
*    dimtb11(2,11) = dimtb11(2,11) + s_all + s_all2
   ENDIF 
  ENDIF 
    
  dimtb11(2,12) = dimtb11(2,12) + IIF(k_u2ok>0,1,0)
  dimtb11(2,13) = dimtb11(2,13) + IIF(k_u2ok>0 and EMPTY(er),1,0) + k_u2ok
  dimtb11(2,14) = dimtb11(2,14) + IIF(k_u2ok>0 and EMPTY(er), s_all, 0) + s_all2ok
  
  DO CASE 
   CASE rslt=321
    dimtb11(1,2) = dimtb11(1,2) + 1
    IF EMPTY(er)
     dimtb11(2,2) = dimtb11(2,2) + 1
    ENDIF 
   CASE rslt=322
    dimtb11(1,3) = dimtb11(1,3) + 1
    IF EMPTY(er)
     dimtb11(2,3) = dimtb11(2,3) + 1
    ENDIF 
   CASE rslt=323
    dimtb11(1,4) = dimtb11(1,4) + 1
    IF EMPTY(er)
     dimtb11(2,4) = dimtb11(2,4) + 1
    ENDIF 
   CASE rslt=324
    dimtb11(1,5) = dimtb11(1,5) + 1
    IF EMPTY(er)
     dimtb11(2,5) = dimtb11(2,5) + 1
    ENDIF 
   CASE rslt=325
    dimtb11(1,6) = dimtb11(1,6) + 1
    IF EMPTY(er)
     dimtb11(2,6) = dimtb11(2,6) + 1
    ENDIF 
  ENDCASE 
 ENDSCAN 
 
* dimtb11(2,15) = RECCOUNT('curpols')

 oExcel.Sheets(4).Select
 WITH oExcel
  .Cells(3,1).Value   = m.qname
 
  .Cells(8,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'

  .Cells(15,01).Value = RECCOUNT('curlpu1')

  .Cells(15,02).Value = dimtb11(1,8) + dimtb11(1,11)
  .Cells(15,03).Value = dimtb11(1,7)
  .Cells(15,04).Value = dimtb11(1,8)
  .Cells(15,05).Value = dimtb11(1,9)
  .Cells(15,06).Value = dimtb11(1,11)

  .Cells(15,07).Value = dimtb11(2,8)

  .Cells(15,08).Value = dimtb11(2,7)
*  .Cells(15,08).Value = dimtb11(2,2)+dimtb11(2,3)+dimtb11(2,4)+dimtb11(2,5)+dimtb11(2,6)
  .Cells(15,09).Value = dimtb11(2,8)

*  .Cells(15,10).Value = dimtb11(2,15)
  .Cells(15,10).Value = dimtb11(2,12)
*  .Cells(15,11).Value = dimtb11(2,12)
  .Cells(15,11).Value = dimtb11(2,13)
  .Cells(15,12).Value = dimtb11(2,14)

  .Cells(15,13).Value = dimtb11(2,2)
  .Cells(15,14).Value = dimtb11(2,3)
  .Cells(15,15).Value = dimtb11(2,4)
  .Cells(15,16).Value = dimtb11(2,5)
  .Cells(15,17).Value = dimtb11(2,6)
 ENDWITH 

 USE IN curlpu1
 USE IN curpols

 WAIT CLEAR 
RETURN 

FUNCTION MakePage5
 WAIT "Формирование листа 5" WINDOW NOWAIT 

 CREATE CURSOR curlpu1 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 
 CREATE CURSOR curpols (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 DIMENSION dimtb11(2,15)
 dimtb11 = 0
 
 SELECT dsp
 SCAN 
  m.cod    = cod
  m.mcod   = mcod
  m.sn_pol = sn_pol
  
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=3
   LOOP 
  ENDIF 
  m.rslt=rslt
  m.c_i = c_i
  IF !INLIST(m.rslt,347,348,349,350,351)
   LOOP 
  ENDIF 
  IF m.rslt=320 AND LEFT(m.c_i,2)!='ДУ'
   LOOP 
  ENDIF 

  IF !SEEK(m.mcod, 'curlpu1')
   INSERT INTO curlpu1 (mcod) VALUES (m.mcod)
  ENDIF 
*  IF !INLIST(cod,101929,101930,101931,101932)
*   LOOP 
*  ENDIF 

  IF EMPTY(k_u2)
   dimtb11(1,7) = dimtb11(1,7) + 1
   dimtb11(1,8) = dimtb11(1,8) + s_all
   IF EMPTY(er)
    dimtb11(2,7) = dimtb11(2,7) + 1
    dimtb11(2,8) = dimtb11(2,8) + s_all
   ENDIF 
  ELSE 
   dimtb11(1,9) = dimtb11(1,9) + 1
   dimtb11(1,10) = dimtb11(1,10) + k_u2
   dimtb11(1,11) = dimtb11(1,11) + s_all2
*   dimtb11(1,11) = dimtb11(1,11) + s_all + s_all2
   IF !SEEK(m.sn_pol, 'curpols')
    INSERT INTO curpols (sn_pol) VALUES (m.sn_pol)
   ENDIF 
   IF EMPTY(er)
    dimtb11(2,9) = dimtb11(2,9) + 1
    dimtb11(2,10) = dimtb11(2,10) + k_u2
    dimtb11(2,11) = dimtb11(2,11) + s_all2
*    dimtb11(2,11) = dimtb11(2,11) + s_all + s_all2
   ENDIF 
  ENDIF 
  
    
  dimtb11(2,12) = dimtb11(2,12) + IIF(k_u2ok>0,1,0)
  dimtb11(2,13) = dimtb11(2,13) + IIF(k_u2ok>0 and EMPTY(er),1,0) + k_u2ok
  dimtb11(2,14) = dimtb11(2,14) + IIF(k_u2ok>0 and EMPTY(er), s_all, 0) + s_all2ok
  
  DO CASE 
   CASE rslt=347
    dimtb11(1,2) = dimtb11(1,2) + 1
    IF EMPTY(er)
     dimtb11(2,2) = dimtb11(2,2) + 1
    ENDIF 
   CASE rslt=348
    dimtb11(1,3) = dimtb11(1,3) + 1
    IF EMPTY(er)
     dimtb11(2,3) = dimtb11(2,3) + 1
    ENDIF 
   CASE rslt=349
    dimtb11(1,4) = dimtb11(1,4) + 1
    IF EMPTY(er)
     dimtb11(2,4) = dimtb11(2,4) + 1
    ENDIF 
   CASE rslt=350
    dimtb11(1,5) = dimtb11(1,5) + 1
    IF EMPTY(er)
     dimtb11(2,5) = dimtb11(2,5) + 1
    ENDIF 
   CASE rslt=351
    dimtb11(1,6) = dimtb11(1,6) + 1
    IF EMPTY(er)
     dimtb11(2,6) = dimtb11(2,6) + 1
    ENDIF 
  ENDCASE 
 ENDSCAN 
 USE IN dsp
 
 dimtb11(2,15) = RECCOUNT('curpols')

 oExcel.Sheets(5).Select
 WITH oExcel
  .Cells(2,1).Value   = m.qname
 
  .Cells(8,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'

  .Cells(15,01).Value = RECCOUNT('curlpu1')

  .Cells(15,02).Value = dimtb11(1,8) + dimtb11(1,11)
  .Cells(15,03).Value = dimtb11(1,7)
  .Cells(15,04).Value = dimtb11(1,8)
  .Cells(15,05).Value = dimtb11(1,9)
  .Cells(15,06).Value = dimtb11(1,11)

  .Cells(15,07).Value = dimtb11(2,8)

  .Cells(15,08).Value = dimtb11(2,7)
  .Cells(15,09).Value = dimtb11(2,8)

  .Cells(15,10).Value = dimtb11(2,15)
  .Cells(15,11).Value = dimtb11(2,12)
  .Cells(15,12).Value = dimtb11(2,14)

  .Cells(15,13).Value = dimtb11(2,2)
  .Cells(15,14).Value = dimtb11(2,3)
  .Cells(15,15).Value = dimtb11(2,4)
  .Cells(15,16).Value = dimtb11(2,5)
  .Cells(15,17).Value = dimtb11(2,6)
 ENDWITH 

 USE IN curlpu1
 USE IN curpols

 WAIT CLEAR 
RETURN 

FUNCTION MakePage6
 WAIT "Формирование листа 6" WINDOW NOWAIT 
 DIMENSION dimpg6(3,2)
 dimpg6=0
* FOR nmon=1 TO tmonth
*  m.lcperiod = STR(tyear,4)+PADL(nmon,2,'0')
 FOR nperiod=18 TO 0 STEP -1 
  m.lcperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT aisoms
  SCAN 
   m.mcod = mcod
   IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT merror
   SCAN 
    m.cod     = cod
    m.err_mee = LEFT(err_mee,2)
    m.koeff   = koeff
    m.s_all   = s_all
    m.e_cod   = e_cod
    m.e_tip   = e_tip
    m.e_ku    = e_ku
    m.IsVed   = .f.
    m.et      = et
    m.dbeg = CTOD('01.01.'+STR(YEAR(tdat1),4))
    m.dend = tdat2
    m.cdat = CTOD('01.'+SUBSTR(e_period,5,2)+SUBSTR(e_period,1,4))
    IF m.err_mee='W0'
     LOOP 
    ENDIF 
    IF !INLIST(m.et,'2','3','7')
     LOOP 
    ENDIF 
    IF !SEEK(m.cod, 'dspcodes')
     LOOP 
    ENDIF 
    IF !INLIST(dspcodes.tip,4,5,6)
     LOOP 
    ENDIF 
    IF !BETWEEN(m.cdat,m.dbeg,m.dend)
     LOOP 
    ENDIF 
    IF m.koeff<=0
     m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
    ELSE 
    m.e_sall = ROUND(m.s_all * m.koeff,2)
    ENDIF 
    m.badsum  = IIF(m.koeff<=0, m.s_all-m.e_sall, m.e_sall)
    DO CASE 
     CASE dspcodes.tip = 4
      dimpg6(1,1) = dimpg6(1,1) + 1
      dimpg6(1,2) = dimpg6(1,2) + m.badsum
     CASE dspcodes.tip = 5
      dimpg6(2,1) = dimpg6(2,1) + 1
      dimpg6(2,2) = dimpg6(2,2) + m.badsum
     CASE dspcodes.tip = 6
      dimpg6(3,1) = dimpg6(3,1) + 1
      dimpg6(3,2) = dimpg6(3,2) + m.badsum
     OTHERWISE 
    ENDCASE 
   ENDSCAN 
   USE IN merror 
   SELECT aisoms

  ENDSCAN 
  USE IN aisoms

 oExcel.Sheets(6).Select
 WITH oExcel
  .Cells(3,1).Value   = m.qname
  .Cells(5,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
  .Cells(08,2).Value = dimpg6(1,1)
  .Cells(09,2).Value = dimpg6(2,1)
  .Cells(10,2).Value = dimpg6(3,1)

  .Cells(08,3).Value = dimpg6(1,2)
  .Cells(09,3).Value = dimpg6(2,2)
  .Cells(10,3).Value = dimpg6(3,2)
 ENDWITH 

 ENDFOR 

 WAIT CLEAR 
RETURN 

FUNCTION MakePage7
 WAIT "Формирование листа 7" WINDOW NOWAIT 
 DIMENSION dimpg6(1,2)
 dimpg6=0
* FOR nmon=1 TO tmonth
*  m.lcperiod = STR(tyear,4)+PADL(nmon,2,'0')
 FOR nperiod=18 TO 0 STEP -1 
  m.lcperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT aisoms
  SCAN 
   m.mcod = mcod
   IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT merror
   SCAN 
    m.cod     = cod
    m.err_mee = LEFT(err_mee,2)
    m.koeff   = koeff
    m.s_all   = s_all
    m.e_cod   = e_cod
    m.e_tip   = e_tip
    m.e_ku    = e_ku
    m.IsVed   = .f.
    m.et      = et
    m.dbeg = CTOD('01.01.'+STR(YEAR(tdat1),4))
    m.dend = tdat2
    m.cdat = CTOD('01.'+SUBSTR(e_period,5,2)+SUBSTR(e_period,1,4))
    IF m.err_mee='W0'
     LOOP 
    ENDIF 
    IF !INLIST(m.et,'2','3','7')
     LOOP 
    ENDIF 
    IF !SEEK(m.cod, 'dspcodes')
     LOOP 
    ENDIF 
    IF !INLIST(dspcodes.tip,2)
     LOOP 
    ENDIF 
    IF !BETWEEN(m.cdat,m.dbeg,m.dend)
     LOOP 
    ENDIF 
    IF m.koeff<=0
     m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
    ELSE 
    m.e_sall = ROUND(m.s_all * m.koeff,2)
    ENDIF 
    m.badsum  = IIF(m.koeff<=0, m.s_all-m.e_sall, m.e_sall)

    dimpg6(1,1) = dimpg6(1,1) + 1
    dimpg6(1,2) = dimpg6(1,2) + m.badsum

   ENDSCAN 
   USE IN merror 
   SELECT aisoms

  ENDSCAN 
  USE IN aisoms

 oExcel.Sheets(7).Select
 WITH oExcel
  .Cells(5,1).Value   = m.qname
  .Cells(7,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
  .Cells(10,1).Value = dimpg6(1,1)
  .Cells(10,2).Value = dimpg6(1,2)
 ENDWITH 

 ENDFOR 

 WAIT CLEAR 
RETURN 

FUNCTION MakePage8
 WAIT "Формирование листа 8" WINDOW NOWAIT 
 DIMENSION dimpg6(1,2)
 dimpg6=0
* FOR nmon=1 TO tmonth
*  m.lcperiod = STR(tyear,4)+PADL(nmon,2,'0')
 FOR nperiod=18 TO 0 STEP -1 
  m.lcperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT aisoms
  SCAN 
   m.mcod = mcod
   IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT merror
   SCAN 
    m.cod     = cod
    m.err_mee = LEFT(err_mee,2)
    m.koeff   = koeff
    m.s_all   = s_all
    m.e_cod   = e_cod
    m.e_tip   = e_tip
    m.e_ku    = e_ku
    m.IsVed   = .f.
    m.et      = et
    m.dbeg = CTOD('01.01.'+STR(YEAR(tdat1),4))
    m.dend = tdat2
    m.cdat = CTOD('01.'+SUBSTR(e_period,5,2)+SUBSTR(e_period,1,4))
    IF m.err_mee='W0'
     LOOP 
    ENDIF 
    IF !INLIST(m.et,'2','3','7')
     LOOP 
    ENDIF 
    IF !SEEK(m.cod, 'dspcodes')
     LOOP 
    ENDIF 
    IF !INLIST(dspcodes.tip,1)
     LOOP 
    ENDIF 
    IF !BETWEEN(m.cdat,m.dbeg,m.dend)
     LOOP 
    ENDIF 
    IF m.koeff<=0
     m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
    ELSE 
    m.e_sall = ROUND(m.s_all * m.koeff,2)
    ENDIF 
    m.badsum  = IIF(m.koeff<=0, m.s_all-m.e_sall, m.e_sall)

    dimpg6(1,1) = dimpg6(1,1) + 1
    dimpg6(1,2) = dimpg6(1,2) + m.badsum

   ENDSCAN 
   USE IN merror 
   SELECT aisoms

  ENDSCAN 
  USE IN aisoms

 oExcel.Sheets(8).Select
 WITH oExcel
  .Cells(5,1).Value   = m.qname
  .Cells(7,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
  .Cells(10,1).Value = dimpg6(1,1)
  .Cells(10,2).Value = dimpg6(1,2)
 ENDWITH 

 ENDFOR 

 WAIT CLEAR 
RETURN 

FUNCTION MakePage9
 WAIT "Формирование листа 9" WINDOW NOWAIT 
 DIMENSION dimpg6(1,2)
 dimpg6=0
* FOR nmon=1 TO tmonth
*  m.lcperiod = STR(tyear,4)+PADL(nmon,2,'0')
 FOR nperiod=18 TO 0 STEP -1 
  m.lcperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT aisoms
  SCAN 
   m.mcod = mcod
   IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    IF USED('merror')
     USE IN merror
    ENDIF 
    IF USED('talon')
     USE IN talon
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT merror
   SET RELATION TO recid INTO talon 
   SCAN 
    m.cod     = cod
    m.err_mee = LEFT(err_mee,2)
    m.koeff   = koeff
    m.s_all   = s_all
    m.e_cod   = e_cod
    m.e_tip   = e_tip
    m.e_ku    = e_ku
    m.IsVed   = .f.
    m.et      = et
    m.dbeg = CTOD('01.01.'+STR(YEAR(tdat1),4))
    m.dend = tdat2
    m.cdat = CTOD('01.'+SUBSTR(e_period,5,2)+SUBSTR(e_period,1,4))
    IF m.err_mee='W0'
     LOOP 
    ENDIF 
    IF !INLIST(m.et,'2','3','7')
     LOOP 
    ENDIF 
    IF !SEEK(m.cod, 'dspcodes')
     LOOP 
    ENDIF 
    IF !INLIST(dspcodes.tip,3)
     LOOP 
    ENDIF 
    IF !BETWEEN(m.cdat,m.dbeg,m.dend)
     LOOP 
    ENDIF 
    m.rslt=talon.rslt
    m.c_i = talon.c_i
    IF !INLIST(m.rslt,320,321,322,323,324,325)
     LOOP 
    ENDIF 
    IF m.rslt=320 AND LEFT(m.c_i,2)!='ДС'
     LOOP 
    ENDIF 
    IF m.koeff<=0
     m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
    ELSE 
    m.e_sall = ROUND(m.s_all * m.koeff,2)
    ENDIF 
    m.badsum  = IIF(m.koeff<=0, m.s_all-m.e_sall, m.e_sall)

    dimpg6(1,1) = dimpg6(1,1) + 1
    dimpg6(1,2) = dimpg6(1,2) + m.badsum

   ENDSCAN 
   SET RELATION OFF INTO talon 
   USE IN merror 
   USE IN talon 
   SELECT aisoms

  ENDSCAN 
  USE IN aisoms

 oExcel.Sheets(9).Select
 WITH oExcel
  .Cells(5,1).Value   = m.qname
  .Cells(7,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
  .Cells(10,1).Value = dimpg6(1,1)
  .Cells(10,2).Value = dimpg6(1,2)
 ENDWITH 

 ENDFOR 

 WAIT CLEAR 
RETURN 

FUNCTION MakePage10
 WAIT "Формирование листа 10" WINDOW NOWAIT 
 DIMENSION dimpg6(1,2)
 dimpg6=0
* FOR nmon=1 TO tmonth
*  m.lcperiod = STR(tyear,4)+PADL(nmon,2,'0')
 FOR nperiod=18 TO 0 STEP -1 
  m.lcperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  
  SELECT aisoms
  SCAN 
   m.mcod = mcod
   IF !fso.FolderExists(pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF RECCOUNT('merror')<=0
    IF USED('merror')
     USE IN merror
    ENDIF 
    IF USED('talon')
     USE IN talon
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT merror
   SET RELATION TO recid INTO talon 
   SCAN 
    m.cod     = cod
    m.err_mee = LEFT(err_mee,2)
    m.koeff   = koeff
    m.s_all   = s_all
    m.e_cod   = e_cod
    m.e_tip   = e_tip
    m.e_ku    = e_ku
    m.IsVed   = .f.
    m.et      = et
    m.dbeg = CTOD('01.01.'+STR(YEAR(tdat1),4))
    m.dend = tdat2
    m.cdat = CTOD('01.'+SUBSTR(e_period,5,2)+SUBSTR(e_period,1,4))
    IF m.err_mee='W0'
     LOOP 
    ENDIF 
    IF !INLIST(m.et,'2','3','7')
     LOOP 
    ENDIF 
    IF !SEEK(m.cod, 'dspcodes')
     LOOP 
    ENDIF 
    IF !INLIST(dspcodes.tip,3)
     LOOP 
    ENDIF 
    IF !BETWEEN(m.cdat,m.dbeg,m.dend)
     LOOP 
    ENDIF 
    m.rslt=talon.rslt
    m.c_i = talon.c_i
    IF !INLIST(m.rslt,347,348,349,350,351)
     LOOP 
    ENDIF 
    IF m.rslt=320 AND LEFT(m.c_i,2)!='ДУ'
     LOOP 
    ENDIF 
    IF m.koeff<=0
     m.e_sall = fsumm(m.e_cod, m.e_tip, m.e_ku, m.IsVed)
    ELSE 
    m.e_sall = ROUND(m.s_all * m.koeff,2)
    ENDIF 
    m.badsum  = IIF(m.koeff<=0, m.s_all-m.e_sall, m.e_sall)

    dimpg6(1,1) = dimpg6(1,1) + 1
    dimpg6(1,2) = dimpg6(1,2) + m.badsum

   ENDSCAN 
   SET RELATION OFF INTO talon 
   USE IN merror 
   USE IN talon 
   SELECT aisoms

  ENDSCAN 
  USE IN aisoms

 oExcel.Sheets(10).Select
 WITH oExcel
  .Cells(5,1).Value   = m.qname
  .Cells(7,1).Value   = 'за '+NameOfMonth(tMonth)+' '+STR(tyear,4)+;
   ' года (нарастающим итогом - начиная со счетов за январь '+STR(tyear,4)+' года)'
  .Cells(10,1).Value = dimpg6(1,1)
  .Cells(10,2).Value = dimpg6(1,2)
 ENDWITH 

 ENDFOR 

 WAIT CLEAR 
RETURN 

FUNCTION IsWDR(w, pol, age, vozr1, vozr2)
 PRIVATE w, pol, age, dr1, dr2
 IF m.w!=m.pol
  RETURN .F.
 ENDIF 
 IF !BETWEEN(m.age, m.vozr1, m.vozr2)
  RETURN .F.
 ENDIF 
RETURN .T.

FUNCTION incdimdsp(nstr)
 PRIVATE nstr
  dimdsp(m.nstr,1) = dimdsp(m.nstr,1) + 1
  dimdsp(m.nstr,2) = dimdsp(m.nstr,2) + s_all

  IF !EMPTY(k_u2)
   dimdsp(m.nstr,3) = dimdsp(m.nstr,3) + 1
   dimdsp(m.nstr,4) = dimdsp(m.nstr,4) + k_u2
   dimdsp(m.nstr,5) = dimdsp(m.nstr,5) + s_all2
  ENDIF 

  IF EMPTY(er)
   dimdsp(m.nstr,6) = dimdsp(m.nstr,6) + 1
   dimdsp(m.nstr,7) = dimdsp(m.nstr,7) + s_all
   IF !EMPTY(k_u2)
    dimdsp(m.nstr,9) = dimdsp(m.nstr,9) + 1
    dimdsp(m.nstr,10) = dimdsp(m.nstr,10) + k_u2ok
    dimdsp(m.nstr,11) = dimdsp(m.nstr,11) + s_all2ok
   ENDIF 

   DO CASE 
    CASE INLIST(rslt,352,353,357,358) && INLIST(rslt,316,352,353,354)
     dimdsp(m.nstr,12) = dimdsp(m.nstr,12) + 1
    CASE INLIST(rslt,317,352)
     dimdsp(m.nstr,13) = dimdsp(m.nstr,13) + 1
    CASE INLIST(rslt,318,353)
     dimdsp(m.nstr,14) = dimdsp(m.nstr,14) + 1
    CASE INLIST(rslt,319,354)
     dimdsp(m.nstr,15) = dimdsp(m.nstr,15) + 1
    CASE INLIST(rslt,355,357)
     dimdsp(m.nstr,16) = dimdsp(m.nstr,16) + 1
    CASE INLIST(rslt,356,358)
     dimdsp(m.nstr,17) = dimdsp(m.nstr,17) + 1
   ENDCASE 
  ENDIF 

RETURN 
