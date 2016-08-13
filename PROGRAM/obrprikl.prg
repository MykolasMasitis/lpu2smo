PROCEDURE ObrPrikl
 IF MESSAGEBOX(CHR(13)+CHR(10)+'—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ œŒ Œ¡–¿Ÿ¿≈ÃŒ—“»?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\obrprikl.xlt')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ Œ“◊≈“¿!'+CHR(13)+CHR(10),0+16,'obrprikl.xlt')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+16,gcperiod)
  RETURN 
 ENDIF 

 IF !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À PEOPLE.DBF ¬ œ≈–»Œƒ≈!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À TALON.DBF ¬ œ≈–»Œƒ≈!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\people', 'people', 'shar', 'sn_pol')>0
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\talon', 'talon', 'shar', 'sn_pol')>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  RETURN 
 ENDIF 
 
 m.lcperiod = STR(IIF(tmonth=1, tyear-1, tyear),4) + PADL(IIF(tmonth=1, 12, tmonth-1),2,'0')
 m.lIsOk = .T.
 IF m.lIsOk = .T.
  IF EMPTY(pattppl)
   m.lIsOk = .F.
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF !fso.FolderExists(pattppl)
   m.lIsOk = .F.
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF !fso.FileExists(pattppl+'\attppl.cfg')
   m.lIsOk = .F.
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF OpenFile(pattppl+'\attppl.cfg', 'attppl', 'shar')>0
   IF USED('attppl')
    USE IN attppl
   ENDIF 
   m.lIsOk = .F.
  ELSE 
   m.attpplbase = ALLTRIM(attppl.pbase)
   USE IN attppl
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF !fso.FolderExists(m.attpplbase)
   m.lIsOk = .F.
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF !fso.FolderExists(m.attpplbase+'\'+m.lcperiod)
   m.lIsOk = .F.
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF !fso.FileExists(m.attpplbase+'\'+m.lcperiod+'\aisoms.dbf')
   m.lIsOk = .F.
  ENDIF 
 ENDIF 
 IF m.lIsOk = .T.
  IF OpenFile(m.attpplbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar')>0
   m.lIsOk = .F.
  ENDIF 
 ENDIF 

 IF !m.lIsOk
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒŒ—“”œ   ATTPPL!'+CHR(13)+CHR(10),0+16,'')
*  IF USED('aisoms')
*   USE IN aisoms
*  ENDIF 
*  IF USED('people')
*   USE IN people
*  ENDIF 
*  IF USED('talon')
*   USE IN talon
*  ENDIF 
*  RETURN 
 ENDIF 
 
 m.totprikl = 0
 IF m.lIsOk
 SELECT aisoms
 SCAN 
  m.totprikl = m.totprikl + (ch_mgf + ad_mgf)
 ENDSCAN 
 ENDIF 

 m.IsVisible = .t. 
 m.IsQuit    = .f.
 
 dotname = ptempl+'\obrprikl.xlt'
 docname = pout+'\op'+m.gcperiod

 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 
 WITH oExcel
  .ReferenceStyle= -4150  && xlR1C1
  .SheetsInNewWorkbook = 1
 ENDWITH 

 oDoc = oExcel.WorkBooks.Add(dotname)

 SELECT people
 COUNT FOR !EMPTY(prmcod) TO m.isobr
 
 CREATE CURSOR curppl (sn_pol c(17), mcod c(7), prmcod c(7))
 INDEX ON mcod + sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 
 SELECT talon 
 SET ORDER TO 
 SET RELATION TO sn_pol INTO people
 SCAN 
  IF (tdat1-people.dr)/365.25<14
   LOOP 
  ENDIF 

  m.mcod = mcod
  m.sn_pol = sn_pol
  m.prmcod = people.prmcod 
  m.var = m.mcod + m.sn_pol
  IF !SEEK(m.var, 'curppl')
   INSERT INTO curppl (sn_pol, mcod, prmcod) VALUES (m.sn_pol, m.mcod, m.prmcod)
  ENDIF 
  
 ENDSCAN 
 SET RELATION OFF INTO people
 
 m.inlpu = 0
 m.outdslpu  = 0
 m.outvedlpu = 0
 SELECT curppl
 SCAN 
  IF EMPTY(prmcod)
   LOOP 
  ENDIF 

  IF mcod=prmcod 
   m.inlpu = m.inlpu + 1
  ELSE 

   IF LEFT(mcod,1)=='0'
     m.outdslpu = m.outdslpu + 1
   ELSE 
    m.outvedlpu = m.outvedlpu + 1
   ENDIF 

  ENDIF 
   
 ENDSCAN 

 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('talon')
  USE IN talon
 ENDIF 

 WITH oExcel
  .Cells(2,1).Value = 'Œ·‡˘‡ÂÏÓÒÚ¸ ÔËÍÂÔÎÂÌÌÓ„Ó Ì‡ÒÂÎÂÌËˇ ‚ ÃŒ „. ÃÓÒÍ‚˚ Á‡ '+;
   NameOfMonth(tmonth) + ' '+STR(tyear,4)+' „Ó‰‡'
  .Cells(11,1).Value = ALLTRIM(qname)
*  .Cells(11,2).Value = m.totprikl
*  .Cells(11,3).Value = m.isobr
*  .Cells(11,4).Value = m.inlpu
*  .Cells(11,5).Value = m.isobr - m.inlpu
*  .Cells(11,6).Value = m.outdslpu
*  .Cells(11,7).Value = m.isobr - m.inlpu - m.outdslpu

  .Cells(11,2).Value = TRANSFORM(ROUND(m.totprikl/1000,3),'999.999')
  .Cells(11,3).Value = TRANSFORM(ROUND(m.isobr/1000,3),'999.999')+' ('+TRANSFORM(ROUND((m.isobr/m.totprikl)*100,2),'99.99')+')'
  .Cells(11,4).Value = TRANSFORM(ROUND(m.inlpu/1000,3),'999.999')+' ('+TRANSFORM(ROUND((m.inlpu/m.isobr)*100,2),'99.99')+')'
  .Cells(11,5).Value = TRANSFORM(ROUND((m.isobr - m.inlpu)/1000,3),'999.999')+' ('+TRANSFORM(ROUND(((m.isobr - m.inlpu)/m.isobr)*100,2),'99.99')+')'
  .Cells(11,6).Value = TRANSFORM(ROUND(m.outdslpu/1000,3),'999.999')+' ('+TRANSFORM(ROUND((m.outdslpu/(m.isobr - m.inlpu))*100,2),'99.99')+')'
  .Cells(11,7).Value = TRANSFORM(ROUND((m.isobr - m.inlpu - m.outdslpu)/1000,3),'999.999')+' ('+TRANSFORM(ROUND(((m.isobr - m.inlpu - m.outdslpu)/(m.isobr - m.inlpu))*100,2),'99.99')+')'
 ENDWITH 

 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls')
 ENDIF 
 oDoc.SaveAs(DocName,18)
 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
  
 
RETURN 