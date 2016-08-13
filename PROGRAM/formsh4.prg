FUNCTION FormSh4(para1)
 m.exptip = para1 && 1 - Ï˝˝, 2 - ˝ÍÏÔ
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ œŒ œ–»ÀŒ∆≈Õ»ﬁ 8?'+CHR(13)+CHR(10),4+32,;
  IIF(m.exptip=2,'‘Œ–Ã¿ ÿ-4 (› Ãœ)','‘Œ–Ã¿ ÿ-4 (Ã››)'))=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee+'\'+m.gcperiod)
  fso.CreateFolder(pmee+'\'+m.gcperiod)
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\priln8.xlt')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À PRILN8.XLT!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 m.docname = 'PrilN8'
 m.dotname = ptempl+'\'+m.docname+'.xlt'
 
 DIMENSION dimerr(77,3)
 dimerr=0
 
 PUBLIC oExcel AS Excel.Application
 WAIT "«‡ÔÛÒÍ MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 WAIT "–¿—◊≈“..." WINDOW NOWAIT 
 SELECT aisoms
 SCAN 
  m.mcod     = mcod
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
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
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
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
   m.err_mee = ALLTRIM(err_mee)
   IF EMPTY(m.err_mee) OR m.err_mee='W0'
    LOOP 
   ENDIF 
   m.et = et
   IF (m.exptip=1 AND !INLIST(m.et,'2','3')) OR (m.exptip=2 AND !INLIST(m.et,'4','5','6'))
    LOOP 
   ENDIF 

   m.osn230 = PADR(ALLTRIM(osn230),5)
   m.s1     = s_1
   
   DO CASE 
    CASE m.osn230 = '1.1.1'
     dimerr(1,1) = dimerr(1,1) + 1
     dimerr(1,2) = dimerr(1,2) + m.s1
    CASE m.osn230 = '1.1.2'
     dimerr(2,1) = dimerr(2,1) + 1
     dimerr(2,2) = dimerr(2,2) + m.s1
    CASE m.osn230 = '1.1.3'
     dimerr(3,1) = dimerr(3,1) + 1
     dimerr(3,2) = dimerr(3,2) + m.s1
    CASE m.osn230 = '1.2.1'
     dimerr(4,1) = dimerr(4,1) + 1
     dimerr(4,2) = dimerr(4,2) + m.s1
    CASE m.osn230 = '1.2.2'
     dimerr(5,1) = dimerr(5,1) + 1
     dimerr(5,2) = dimerr(5,2) + m.s1
    CASE m.osn230 = '1.3.1'
     dimerr(6,1) = dimerr(6,1) + 1
     dimerr(6,2) = dimerr(6,2) + m.s1
    CASE m.osn230 = '1.3.2'
     dimerr(7,1) = dimerr(7,1) + 1
     dimerr(7,2) = dimerr(7,2) + m.s1
    CASE m.osn230 = '1.4. '
     dimerr(8,1) = dimerr(8,1) + 1
     dimerr(8,2) = dimerr(8,2) + m.s1
    CASE m.osn230 = '1.5. '
     dimerr(9,1) = dimerr(9,1) + 1
     dimerr(9,2) = dimerr(9,2) + m.s1
    CASE m.osn230 = '2.1. '
     dimerr(10,1) = dimerr(10,1) + 1
     dimerr(10,2) = dimerr(10,2) + m.s1
    CASE m.osn230 = '2.2.1'
     dimerr(11,1) = dimerr(11,1) + 1
     dimerr(11,2) = dimerr(11,2) + m.s1
    CASE m.osn230 = '2.2.2'
     dimerr(12,1) = dimerr(12,1) + 1
     dimerr(12,2) = dimerr(12,2) + m.s1
    CASE m.osn230 = '2.2.3'
     dimerr(13,1) = dimerr(13,1) + 1
     dimerr(13,2) = dimerr(13,2) + m.s1
    CASE m.osn230 = '2.2.4'
     dimerr(14,1) = dimerr(14,1) + 1
     dimerr(14,2) = dimerr(14,2) + m.s1
    CASE m.osn230 = '2.2.5'
     dimerr(15,1) = dimerr(15,1) + 1
     dimerr(15,2) = dimerr(15,2) + m.s1
    CASE m.osn230 = '2.2.6'
     dimerr(16,1) = dimerr(16,1) + 1
     dimerr(16,2) = dimerr(16,2) + m.s1
    CASE m.osn230 = '2.3. '
     dimerr(17,1) = dimerr(17,1) + 1
     dimerr(17,2) = dimerr(17,2) + m.s1
    CASE m.osn230 = '2.4.1'
     dimerr(18,1) = dimerr(18,1) + 1
     dimerr(18,2) = dimerr(18,2) + m.s1
    CASE m.osn230 = '2.4.2'
     dimerr(19,1) = dimerr(19,1) + 1
     dimerr(19,2) = dimerr(19,2) + m.s1
    CASE m.osn230 = '2.4.3'
     dimerr(20,1) = dimerr(20,1) + 1
     dimerr(20,2) = dimerr(20,2) + m.s1

    CASE m.osn230 = '2.4.4'
     dimerr(21,1) = dimerr(21,1) + 1
     dimerr(21,2) = dimerr(21,2) + m.s1
    CASE m.osn230 = '2.4.5'
     dimerr(22,1) = dimerr(22,1) + 1
     dimerr(22,2) = dimerr(22,2) + m.s1
    CASE m.osn230 = '2.4.6'
     dimerr(23,1) = dimerr(23,1) + 1
     dimerr(23,2) = dimerr(23,2) + m.s1
    CASE m.osn230 = '3.1. '
     dimerr(24,1) = dimerr(24,1) + 1
     dimerr(24,2) = dimerr(24,2) + m.s1
    CASE m.osn230 = '3.2.1'
     dimerr(25,1) = dimerr(25,1) + 1
     dimerr(25,2) = dimerr(25,2) + m.s1
    CASE m.osn230 = '3.2.2'
     dimerr(26,1) = dimerr(26,1) + 1
     dimerr(26,2) = dimerr(26,2) + m.s1
    CASE m.osn230 = '3.2.3'
     dimerr(27,1) = dimerr(27,1) + 1
     dimerr(27,2) = dimerr(27,2) + m.s1
    CASE m.osn230 = '3.2.4'
     dimerr(28,1) = dimerr(28,1) + 1
     dimerr(28,2) = dimerr(28,2) + m.s1
    CASE m.osn230 = '3.2.5'
     dimerr(29,1) = dimerr(29,1) + 1
     dimerr(29,2) = dimerr(29,2) + m.s1
    CASE m.osn230 = '3.3.1'
     dimerr(30,1) = dimerr(30,1) + 1
     dimerr(30,2) = dimerr(30,2) + m.s1

    CASE m.osn230 = '3.3.2'
     dimerr(31,1) = dimerr(31,1) + 1
     dimerr(31,2) = dimerr(31,2) + m.s1
    CASE m.osn230 = '3.4. '
     dimerr(32,1) = dimerr(32,1) + 1
     dimerr(32,2) = dimerr(32,2) + m.s1
    CASE m.osn230 = '3.5. '
     dimerr(33,1) = dimerr(33,1) + 1
     dimerr(33,2) = dimerr(33,2) + m.s1
    CASE m.osn230 = '3.6. '
     dimerr(34,1) = dimerr(34,1) + 1
     dimerr(34,2) = dimerr(34,2) + m.s1
    CASE m.osn230 = '3.7. '
     dimerr(35,1) = dimerr(35,1) + 1
     dimerr(35,2) = dimerr(35,2) + m.s1
    CASE m.osn230 = '3.8. '
     dimerr(36,1) = dimerr(36,1) + 1
     dimerr(36,2) = dimerr(36,2) + m.s1
    CASE m.osn230 = '3.9. '
     dimerr(37,1) = dimerr(37,1) + 1
     dimerr(37,2) = dimerr(37,2) + m.s1
    CASE m.osn230 = '3.10.'
     dimerr(38,1) = dimerr(38,1) + 1
     dimerr(38,2) = dimerr(38,2) + m.s1
    CASE m.osn230 = '3.11.'
     dimerr(39,1) = dimerr(39,1) + 1
     dimerr(39,2) = dimerr(39,2) + m.s1
    CASE m.osn230 = '3.12.'
     dimerr(40,1) = dimerr(40,1) + 1
     dimerr(40,2) = dimerr(40,2) + m.s1
    CASE m.osn230 = '3.13.'
     dimerr(41,1) = dimerr(41,1) + 1
     dimerr(41,2) = dimerr(41,2) + m.s1
    CASE m.osn230 = '3.14.'
     dimerr(42,1) = dimerr(42,1) + 1
     dimerr(42,2) = dimerr(42,2) + m.s1

    CASE m.osn230 = '4.1. '
     dimerr(43,1) = dimerr(43,1) + 1
     dimerr(43,2) = dimerr(43,2) + m.s1
    CASE m.osn230 = '4.2. '
     dimerr(44,1) = dimerr(44,1) + 1
     dimerr(44,2) = dimerr(44,2) + m.s1
    CASE m.osn230 = '4.3. '
     dimerr(45,1) = dimerr(45,1) + 1
     dimerr(45,2) = dimerr(45,2) + m.s1
    CASE m.osn230 = '4.4. '
     dimerr(46,1) = dimerr(46,1) + 1
     dimerr(46,2) = dimerr(46,2) + m.s1
    CASE m.osn230 = '4.5. '
     dimerr(47,1) = dimerr(47,1) + 1
     dimerr(47,2) = dimerr(47,2) + m.s1
    CASE m.osn230 = '4.6.1'
     dimerr(48,1) = dimerr(48,1) + 1
     dimerr(48,2) = dimerr(48,2) + m.s1
    CASE m.osn230 = '4.6.2'
     dimerr(49,1) = dimerr(49,1) + 1
     dimerr(49,2) = dimerr(49,2) + m.s1
    CASE m.osn230 = '5.1.1'
     dimerr(50,1) = dimerr(50,1) + 1
     dimerr(50,2) = dimerr(50,2) + m.s1
    CASE m.osn230 = '5.1.2'
     dimerr(51,1) = dimerr(51,1) + 1
     dimerr(51,2) = dimerr(51,2) + m.s1
    CASE m.osn230 = '5.1.3'
     dimerr(52,1) = dimerr(52,1) + 1
     dimerr(52,2) = dimerr(52,2) + m.s1
    CASE m.osn230 = '5.1.4'
     dimerr(53,1) = dimerr(53,1) + 1
     dimerr(53,2) = dimerr(53,2) + m.s1
    CASE m.osn230 = '5.1.5'
     dimerr(54,1) = dimerr(54,1) + 1
     dimerr(54,2) = dimerr(54,2) + m.s1
    CASE m.osn230 = '5.1.6'
     dimerr(55,1) = dimerr(55,1) + 1
     dimerr(55,2) = dimerr(55,2) + m.s1
    CASE m.osn230 = '5.2.1'
     dimerr(56,1) = dimerr(56,1) + 1
     dimerr(56,2) = dimerr(56,2) + m.s1
    CASE m.osn230 = '5.2.2'
     dimerr(57,1) = dimerr(57,1) + 1
     dimerr(57,2) = dimerr(57,2) + m.s1
    CASE m.osn230 = '5.2.3'
     dimerr(58,1) = dimerr(58,1) + 1
     dimerr(58,2) = dimerr(58,2) + m.s1
    CASE m.osn230 = '5.2.4'
     dimerr(59,1) = dimerr(59,1) + 1
     dimerr(59,2) = dimerr(59,2) + m.s1
    CASE m.osn230 = '5.2.5'
     dimerr(60,1) = dimerr(60,1) + 1
     dimerr(60,2) = dimerr(60,2) + m.s1
    CASE m.osn230 = '5.3.1'
     dimerr(61,1) = dimerr(61,1) + 1
     dimerr(61,2) = dimerr(61,2) + m.s1
    CASE m.osn230 = '5.3.2'
     dimerr(62,1) = dimerr(62,1) + 1
     dimerr(62,2) = dimerr(62,2) + m.s1
    CASE m.osn230 = '5.3.3'
     dimerr(63,1) = dimerr(63,1) + 1
     dimerr(63,2) = dimerr(63,2) + m.s1
    CASE m.osn230 = '5.4.1'
     dimerr(64,1) = dimerr(64,1) + 1
     dimerr(64,2) = dimerr(64,2) + m.s1
    CASE m.osn230 = '5.4.2'
     dimerr(65,1) = dimerr(65,1) + 1
     dimerr(65,2) = dimerr(65,2) + m.s1
    CASE m.osn230 = '5.5.1'
     dimerr(66,1) = dimerr(66,1) + 1
     dimerr(66,2) = dimerr(66,2) + m.s1
    CASE m.osn230 = '5.5.2'
     dimerr(67,1) = dimerr(67,1) + 1
     dimerr(67,2) = dimerr(67,2) + m.s1
    CASE m.osn230 = '5.5.3'
     dimerr(68,1) = dimerr(68,1) + 1
     dimerr(68,2) = dimerr(68,2) + m.s1
    CASE m.osn230 = '5.6. '
     dimerr(69,1) = dimerr(69,1) + 1
     dimerr(69,2) = dimerr(69,2) + m.s1
    CASE m.osn230 = '5.7.1'
     dimerr(70,1) = dimerr(70,1) + 1
     dimerr(70,2) = dimerr(70,2) + m.s1
    CASE m.osn230 = '5.7.2'
     dimerr(71,1) = dimerr(71,1) + 1
     dimerr(71,2) = dimerr(71,2) + m.s1
    CASE m.osn230 = '5.7.3'
     dimerr(72,1) = dimerr(72,1) + 1
     dimerr(72,2) = dimerr(72,2) + m.s1
    CASE m.osn230 = '5.7.4'
     dimerr(73,1) = dimerr(73,1) + 1
     dimerr(73,2) = dimerr(73,2) + m.s1
    CASE m.osn230 = '5.7.5'
     dimerr(74,1) = dimerr(74,1) + 1
     dimerr(74,2) = dimerr(74,2) + m.s1
    CASE m.osn230 = '5.7.6'
     dimerr(75,1) = dimerr(75,1) + 1
     dimerr(75,2) = dimerr(75,2) + m.s1
    OTHERWISE 
    dimerr(76,1) = dimerr(76,1) - 1
    dimerr(76,2) = dimerr(76,2) - m.s1

   ENDCASE 
   dimerr(77,1) = dimerr(77,1) + 1
   dimerr(77,2) = dimerr(77,2) + m.s1

  ENDSCAN 
  SET RELATION OFF INTO talon 
  USE IN talon 
  USE IN merror 

  SELECT aisoms 
  
 ENDSCAN
 USE IN aisoms
 WAIT CLEAR 

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Œ“◊≈“¿..." WINDOW NOWAIT 
 m.nOpBooks = oExcel.Workbooks.Count 
 IF m.nOpBooks>0
  FOR m.nBook=1 TO m.nOpBooks
   m.cBookName = LOWER(ALLTRIM(oExcel.Workbooks.Item(m.nBook).Name))
   IF INLIST(m.cBookName,LOWER(m.DocName),'sh4_')
    oExcel.Workbooks.Item(m.nBook).Close 
   ENDIF 
  NEXT 
 ENDIF 
 oBook       = oExcel.WorkBooks.Add(m.dotname)
 oSheet      = oexcel.ActiveSheet
 IF m.exptip=1
  oSheet.name = 'Ã›› Á‡ '+LOWER(nameofmonth(tmonth))+' '+STR(tyear,4)+' „.'
 ELSE 
  oSheet.name = '› Ãœ Á‡ '+LOWER(nameofmonth(tmonth))+' '+STR(tyear,4)+' „.'
 ENDIF 

 WITH oExcel
  .Cells(3,3) = dimerr(1,1)
  .Cells(3,4) = dimerr(1,2)
  .Cells(3,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(1,1)/dimerr(77,1))*100,4), 0)
  .Cells(4,3) = dimerr(2,1)
  .Cells(4,4) = dimerr(2,2)
  .Cells(4,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(2,1)/dimerr(77,1))*100,4), 0)
  .Cells(5,3) = dimerr(3,1)
  .Cells(5,4) = dimerr(3,2)
  .Cells(5,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(3,1)/dimerr(77,1))*100,4), 0)
  .Cells(6,3) = dimerr(4,1)
  .Cells(6,4) = dimerr(4,2)
  .Cells(6,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(4,1)/dimerr(77,1))*100,4), 0)
  .Cells(7,3) = dimerr(5,1)
  .Cells(7,4) = dimerr(5,2)
  .Cells(7,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(5,1)/dimerr(77,1))*100,4), 0)
  .Cells(8,3) = dimerr(6,1)
  .Cells(8,4) = dimerr(6,2)
  .Cells(8,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(6,1)/dimerr(77,1))*100,4), 0)
  .Cells(9,3) = dimerr(7,1)
  .Cells(9,4) = dimerr(7,2)
  .Cells(9,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(7,1)/dimerr(77,1))*100,4), 0)
  .Cells(10,3) = dimerr(8,1)
  .Cells(10,4) = dimerr(8,2)
  .Cells(10,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(8,1)/dimerr(77,1))*100,4), 0)

  .Cells(11,3) = dimerr(9,1)
  .Cells(11,4) = dimerr(9,2)
  .Cells(11,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(9,1)/dimerr(77,1))*100,4), 0)
  .Cells(12,3) = dimerr(10,1)
  .Cells(12,4) = dimerr(10,2)
  .Cells(12,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(10,1)/dimerr(77,1))*100,4), 0)
  .Cells(13,3) = dimerr(11,1)
  .Cells(13,4) = dimerr(11,2)
  .Cells(13,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(11,1)/dimerr(77,1))*100,4), 0)
  .Cells(14,3) = dimerr(12,1)
  .Cells(14,4) = dimerr(12,2)
  .Cells(14,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(12,1)/dimerr(77,1))*100,4), 0)
  .Cells(15,3) = dimerr(13,1)
  .Cells(15,4) = dimerr(13,2)
  .Cells(15,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(13,1)/dimerr(77,1))*100,4), 0)
  .Cells(16,3) = dimerr(14,1)
  .Cells(16,4) = dimerr(14,2)
  .Cells(16,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(14,1)/dimerr(77,1))*100,4), 0)
  .Cells(17,3) = dimerr(15,1)
  .Cells(17,4) = dimerr(15,2)
  .Cells(17,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(15,1)/dimerr(77,1))*100,4), 0)
  .Cells(18,3) = dimerr(16,1)
  .Cells(18,4) = dimerr(16,2)
  .Cells(18,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(16,1)/dimerr(77,1))*100,4), 0)
  .Cells(19,3) = dimerr(17,1)
  .Cells(19,4) = dimerr(17,2)
  .Cells(19,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(17,1)/dimerr(77,1))*100,4), 0)
  .Cells(20,3) = dimerr(18,1)
  .Cells(20,4) = dimerr(18,2)
  .Cells(20,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(18,1)/dimerr(77,1))*100,4), 0)
  .Cells(21,3) = dimerr(19,1)
  .Cells(21,4) = dimerr(19,2)
  .Cells(21,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(19,1)/dimerr(77,1))*100,4), 0)
  .Cells(22,3) = dimerr(20,1)
  .Cells(22,4) = dimerr(20,2)
  .Cells(22,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(20,1)/dimerr(77,1))*100,4), 0)
  .Cells(23,3) = dimerr(21,1)
  .Cells(23,4) = dimerr(21,2)
  .Cells(23,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(21,1)/dimerr(77,1))*100,4), 0)
  .Cells(24,3) = dimerr(22,1)
  .Cells(24,4) = dimerr(22,2)
  .Cells(24,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(22,1)/dimerr(77,1))*100,4), 0)
  .Cells(25,3) = dimerr(23,1)
  .Cells(25,4) = dimerr(23,2)
  .Cells(25,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(23,1)/dimerr(77,1))*100,4), 0)

  .Cells(26,3) = dimerr(24,1)
  .Cells(26,4) = dimerr(24,2)
  .Cells(26,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(24,1)/dimerr(77,1))*100,4), 0)
  .Cells(27,3) = dimerr(25,1)
  .Cells(27,4) = dimerr(25,2)
  .Cells(27,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(25,1)/dimerr(77,1))*100,4), 0)
  .Cells(28,3) = dimerr(26,1)
  .Cells(28,4) = dimerr(26,2)
  .Cells(28,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(26,1)/dimerr(77,1))*100,4), 0)
  .Cells(29,3) = dimerr(27,1)
  .Cells(29,4) = dimerr(27,2)
  .Cells(29,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(27,1)/dimerr(77,1))*100,4), 0)
  .Cells(30,3) = dimerr(28,1)
  .Cells(30,4) = dimerr(28,2)
  .Cells(30,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(28,1)/dimerr(77,1))*100,4), 0)
  .Cells(31,3) = dimerr(29,1)
  .Cells(31,4) = dimerr(29,2)
  .Cells(31,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(29,1)/dimerr(77,1))*100,4), 0)
  .Cells(32,3) = dimerr(30,1)
  .Cells(32,4) = dimerr(30,2)
  .Cells(32,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(30,1)/dimerr(77,1))*100,4), 0)
  .Cells(33,3) = dimerr(31,1)
  .Cells(33,4) = dimerr(31,2)
  .Cells(33,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(31,1)/dimerr(77,1))*100,4), 0)
  .Cells(34,3) = dimerr(32,1)
  .Cells(34,4) = dimerr(32,2)
  .Cells(34,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(32,1)/dimerr(77,1))*100,4), 0)
  .Cells(35,3) = dimerr(33,1)
  .Cells(35,4) = dimerr(33,2)
  .Cells(35,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(33,1)/dimerr(77,1))*100,4), 0)
  .Cells(36,3) = dimerr(34,1)
  .Cells(36,4) = dimerr(34,2)
  .Cells(36,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(34,1)/dimerr(77,1))*100,4), 0)
  .Cells(37,3) = dimerr(35,1)
  .Cells(37,4) = dimerr(35,2)
  .Cells(37,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(35,1)/dimerr(77,1))*100,4), 0)
  .Cells(38,3) = dimerr(36,1)
  .Cells(38,4) = dimerr(36,2)
  .Cells(38,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(36,1)/dimerr(77,1))*100,4), 0)
  .Cells(39,3) = dimerr(37,1)
  .Cells(39,4) = dimerr(37,2)
  .Cells(39,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(37,1)/dimerr(77,1))*100,4), 0)

  .Cells(40,3) = dimerr(38,1)
  .Cells(40,4) = dimerr(38,2)
  .Cells(40,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(38,1)/dimerr(77,1))*100,4), 0)
  .Cells(41,3) = dimerr(39,1)
  .Cells(41,4) = dimerr(39,2)
  .Cells(41,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(39,1)/dimerr(77,1))*100,4), 0)
  .Cells(42,3) = dimerr(40,1)
  .Cells(42,4) = dimerr(40,2)
  .Cells(42,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(40,1)/dimerr(77,1))*100,4), 0)
  .Cells(43,3) = dimerr(41,1)
  .Cells(43,4) = dimerr(41,2)
  .Cells(43,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(41,1)/dimerr(77,1))*100,4), 0)
  .Cells(44,3) = dimerr(42,1)
  .Cells(44,4) = dimerr(42,2)
  .Cells(44,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(42,1)/dimerr(77,1))*100,4), 0)
  .Cells(45,3) = dimerr(43,1)
  .Cells(45,4) = dimerr(43,2)
  .Cells(45,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(43,1)/dimerr(77,1))*100,4), 0)
  .Cells(46,3) = dimerr(44,1)
  .Cells(46,4) = dimerr(44,2)
  .Cells(46,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(44,1)/dimerr(77,1))*100,4), 0)
  .Cells(47,3) = dimerr(45,1)
  .Cells(47,4) = dimerr(45,2)
  .Cells(47,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(45,1)/dimerr(77,1))*100,4), 0)

  .Cells(48,3) = dimerr(46,1)
  .Cells(48,4) = dimerr(46,2)
  .Cells(48,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(46,1)/dimerr(77,1))*100,4), 0)
  .Cells(49,3) = dimerr(47,1)
  .Cells(49,4) = dimerr(47,2)
  .Cells(49,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(47,1)/dimerr(77,1))*100,4), 0)
  .Cells(50,3) = dimerr(48,1)
  .Cells(50,4) = dimerr(48,2)
  .Cells(50,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(48,1)/dimerr(77,1))*100,4), 0)
  .Cells(51,3) = dimerr(49,1)
  .Cells(51,4) = dimerr(49,2)
  .Cells(51,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(49,1)/dimerr(77,1))*100,4), 0)
  .Cells(52,3) = dimerr(50,1)
  .Cells(52,4) = dimerr(50,2)
  .Cells(52,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(50,1)/dimerr(77,1))*100,4), 0)
  .Cells(53,3) = dimerr(51,1)
  .Cells(53,4) = dimerr(51,2)
  .Cells(53,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(51,1)/dimerr(77,1))*100,4), 0)

  .Cells(54,3) = dimerr(52,1)
  .Cells(54,4) = dimerr(52,2)
  .Cells(54,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(52,1)/dimerr(77,1))*100,4), 0)
  .Cells(55,3) = dimerr(53,1)
  .Cells(55,4) = dimerr(53,2)
  .Cells(55,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(53,1)/dimerr(77,1))*100,4), 0)
  .Cells(56,3) = dimerr(54,1)
  .Cells(56,4) = dimerr(54,2)
  .Cells(56,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(54,1)/dimerr(77,1))*100,4), 0)
  .Cells(57,3) = dimerr(55,1)
  .Cells(57,4) = dimerr(55,2)
  .Cells(57,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(55,1)/dimerr(77,1))*100,4), 0)
  .Cells(58,3) = dimerr(56,1)
  .Cells(58,4) = dimerr(56,2)
  .Cells(58,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(56,1)/dimerr(77,1))*100,4), 0)

  .Cells(59,3) = dimerr(57,1)
  .Cells(59,4) = dimerr(57,2)
  .Cells(59,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(57,1)/dimerr(77,1))*100,4), 0)
  .Cells(60,3) = dimerr(58,1)
  .Cells(60,4) = dimerr(58,2)
  .Cells(60,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(58,1)/dimerr(77,1))*100,4), 0)
  .Cells(61,3) = dimerr(59,1)
  .Cells(61,4) = dimerr(59,2)
  .Cells(61,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(59,1)/dimerr(77,1))*100,4), 0)

  .Cells(62,3) = dimerr(60,1)
  .Cells(62,4) = dimerr(60,2)
  .Cells(62,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(60,1)/dimerr(77,1))*100,4), 0)
  .Cells(63,3) = dimerr(61,1)
  .Cells(63,4) = dimerr(61,2)
  .Cells(63,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(61,1)/dimerr(77,1))*100,4), 0)

  .Cells(64,3) = dimerr(62,1)
  .Cells(64,4) = dimerr(62,2)
  .Cells(64,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(62,1)/dimerr(77,1))*100,4), 0)
  .Cells(65,3) = dimerr(63,1)
  .Cells(65,4) = dimerr(63,2)
  .Cells(65,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(63,1)/dimerr(77,1))*100,4), 0)
  .Cells(66,3) = dimerr(64,1)
  .Cells(66,4) = dimerr(64,2)
  .Cells(66,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(64,1)/dimerr(77,1))*100,4), 0)
  .Cells(67,3) = dimerr(65,1)
  .Cells(67,4) = dimerr(65,2)
  .Cells(67,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(65,1)/dimerr(77,1))*100,4), 0)
  .Cells(68,3) = dimerr(66,1)
  .Cells(68,4) = dimerr(66,2)
  .Cells(68,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(66,1)/dimerr(77,1))*100,4), 0)
  .Cells(69,3) = dimerr(67,1)
  .Cells(69,4) = dimerr(67,2)
  .Cells(69,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(67,1)/dimerr(77,1))*100,4), 0)
  .Cells(70,3) = dimerr(68,1)
  .Cells(70,4) = dimerr(68,2)
  .Cells(70,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(68,1)/dimerr(77,1))*100,4), 0)
  .Cells(71,3) = dimerr(69,1)
  .Cells(71,4) = dimerr(69,2)
  .Cells(71,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(69,1)/dimerr(77,1))*100,4), 0)
  .Cells(72,3) = dimerr(70,1)
  .Cells(72,4) = dimerr(70,2)
  .Cells(72,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(70,1)/dimerr(77,1))*100,4), 0)
  .Cells(73,3) = dimerr(71,1)
  .Cells(73,4) = dimerr(71,2)
  .Cells(73,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(71,1)/dimerr(77,1))*100,4), 0)
  .Cells(74,3) = dimerr(72,1)
  .Cells(74,4) = dimerr(72,2)
  .Cells(74,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(72,1)/dimerr(77,1))*100,4), 0)
  .Cells(75,3) = dimerr(73,1)
  .Cells(75,4) = dimerr(73,2)
  .Cells(75,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(73,1)/dimerr(77,1))*100,4), 0)
  .Cells(76,3) = dimerr(74,1)
  .Cells(76,4) = dimerr(74,2)
  .Cells(76,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(74,1)/dimerr(77,1))*100,4), 0)
  .Cells(77,3) = dimerr(75,1)
  .Cells(77,4) = dimerr(75,2)
  .Cells(77,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(75,1)/dimerr(77,1))*100,4), 0)

  .Cells(78,3) = dimerr(77,1)
  .Cells(78,4) = dimerr(77,2)
  .Cells(78,5) = IIF(dimerr(77,1)>0, ROUND((dimerr(77,1)/dimerr(77,1))*100,4), 0)
 ENDWITH 
 WAIT CLEAR 

 m.DocName = 'sh4_'+STR(m.exptip,1)
 IF fso.FileExists(pmee+'\'+m.gcperiod+'\'+m.DocName+'.xls')
  TRY 
   fso.DeleteFile(pmee+'\'+m.gcperiod+'\'+m.DocName+'.xls')
   oBook.SaveAs(pmee+'\'+m.gcperiod+'\'+m.DocName,18)
  CATCH  
   MESSAGEBOX('‘¿…À '+m.DocName+'.XLS Œ “–€“!',0+64,'')
  ENDTRY 
 ELSE 
  oBook.SaveAs(pmee+'\'+m.gcperiod+'\'+m.DocName,18)
 ENDIF 
 oExcel.Visible = .t.

RETURN 