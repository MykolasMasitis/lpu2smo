PROCEDURE VolControls
 m.IsVisible = .T.
 m.IsQuit    = .F.

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÔÎÐÌÓ ÊÎÍÒÐÎËß ÎÁÚÅÌÎÂ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(ptempl+'\volcontrols.xlt')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË ØÀÁËÎÍÀ VOLCONTROLS.XLT'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pcommon+'\lpudogs.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÑÏÐÀÂÎ×ÍÈÊ ÄÎÃÎÂÎÐÎÂ LPUDOGS.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'mcod')>0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT lpudogs
 IF EMPTY(FIELD('kv01')) OR EMPTY(FIELD('kv02')) OR EMPTY(FIELD('kv03')) OR EMPTY(FIELD('kv04'))
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'Â ÑÏÐÀÂÎ×ÍÈÊÅ LPUDOGS.DBF ÎÒÑÓÒÑÒÂÓÞÒ ÏÎËß KV01,KV02,KV03,KV04!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 DotName = pTempl + "\VolControls.xlt"
 DocName = pOut + "\VC" + UPPER(m.qcod)
 
 IF fso.FileExists(DocName+'.xls')
  fso.DeleteFile(DocName+'.xls')
 ENDIF 

 PUBLIC oExcel AS Excel.Application

 WAIT "Çàïóñê MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 WAIT "ÔÎÐÌÈÐÎÂÀÍÈÅ ÎÒ×ÅÒÀ..." WINDOW NOWAIT 
 oDoc = oExcel.WorkBooks.Add(dotname)
 
 FOR nmonth=1 TO 12
* FOR nmonth=1 TO 3
  tdir = STR(tyear,4)+PADL(nmonth,2,'0')
  IF !fso.FolderExists(pbase+'\'+tdir)
   LOOP 
  ENDIF
  IF !fso.FileExists(pbase+'\'+tdir+'\aisoms.dbf')
   LOOP 
  ENDIF  
  IF OpenFile(pbase+'\'+tdir+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+tdir+'\pr4', 'pr4', 'shar', 'lpuid')>0
   IF USED('pr4')
    USE IN pr4
   ENDIF 
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 

  WAIT tdir+'...' WINDOW NOWAIT 

  SELECT lpudogs
  m.nstr = 10
  m.npp = 1
  SCAN 
   m.mcod    = mcod 
   m.lpu_id  = lpu_id
   m.lpuname = IIF(SEEK(m.lpu_id, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
   m.kv01    = kv01
   m.kv02    = kv02
   m.kv03    = kv03
   m.kv04    = kv04
   m.plan    = m.kv01 + m.kv02 + m.kv03 + m.kv04

   IF SEEK(m.lpu_id, 'pilot')
    m.IsPilot = .T.
   ELSE 
    m.IsPilot = .F.
   ENDIF 

   IF !SEEK(m.lpu_id, 'aisoms')
    m.koplate = 0
   ELSE 
    IF m.IsPilot = .F.
     m.vvir    = aisoms.s_pred - aisoms.sum_flk - (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b
     m.koplate = IIF(m.vvir>0, m.vvir, aisoms.s_avans)
    ELSE 
     IF SEEK(m.lpu_id, 'pr4')
      m.koplate = pr4.finval - pr4.s_others + pr4.s_guests + (pr4.s_kompl + pr4.s_dst) + ;
       pr4.s_npilot + pr4.s_empty - (aisoms.e_mee+aisoms.e_ekmp)
     ELSE 
      m.koplate = 0
     ENDIF 
    ENDIF 
   ENDIF 

   m.valcell = oExcel.Cells(m.nstr,2).Value
   IF ISNULL(m.valcell)
    oExcel.Cells(m.nstr,1).Value = m.npp
    oExcel.Cells(m.nstr,2).Value = m.mcod
    oExcel.Cells(m.nstr,3).Value = m.lpuname
    oExcel.Cells(m.nstr,4).Value = m.plan
     
    oExcel.Cells(m.nstr,7).Value  = m.kv01
    oExcel.Cells(m.nstr,12).Value = m.kv02
    oExcel.Cells(m.nstr,17).Value = m.kv03
    oExcel.Cells(m.nstr,22).Value = m.kv04
   ENDIF 
    
   DO CASE 
    CASE nmonth = 1
     oExcel.Cells(m.nstr,8).Value   = m.koplate
     oExcel.Cells(m.nstr,11).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,11).Value), 0, oExcel.Cells(m.nstr,11).Value) + m.koplate
    CASE nmonth = 2
     oExcel.Cells(m.nstr,9).Value   = m.koplate
     oExcel.Cells(m.nstr,11).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,11).Value), 0, oExcel.Cells(m.nstr,11).Value) + m.koplate
    CASE nmonth = 3
     oExcel.Cells(m.nstr,10).Value  = m.koplate
     oExcel.Cells(m.nstr,11).Value   = m.kv01 - (IIF(ISNULL(oExcel.Cells(m.nstr,11).Value), 0, oExcel.Cells(m.nstr,11).Value) + m.koplate)
    CASE nmonth = 4
     oExcel.Cells(m.nstr,13).Value  = m.koplate
     oExcel.Cells(m.nstr,16).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,16).Value), 0, oExcel.Cells(m.nstr,16).Value) + m.koplate
    CASE nmonth = 5
     oExcel.Cells(m.nstr,14).Value  = m.koplate
     oExcel.Cells(m.nstr,16).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,16).Value), 0, oExcel.Cells(m.nstr,16).Value) + m.koplate
    CASE nmonth = 6
     oExcel.Cells(m.nstr,15).Value  = m.koplate
     oExcel.Cells(m.nstr,16).Value   = m.kv02 - (IIF(ISNULL(oExcel.Cells(m.nstr,16).Value), 0, oExcel.Cells(m.nstr,16).Value) + m.koplate)
    CASE nmonth = 7
     oExcel.Cells(m.nstr,18).Value  = m.koplate
     oExcel.Cells(m.nstr,21).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,21).Value), 0, oExcel.Cells(m.nstr,21).Value) + m.koplate
    CASE nmonth = 8
     oExcel.Cells(m.nstr,19).Value  = m.koplate
     oExcel.Cells(m.nstr,21).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,21).Value), 0, oExcel.Cells(m.nstr,21).Value) + m.koplate
    CASE nmonth = 9
     oExcel.Cells(m.nstr,20).Value  = m.koplate
     oExcel.Cells(m.nstr,21).Value   = m.kv03 - (IIF(ISNULL(oExcel.Cells(m.nstr,21).Value), 0, oExcel.Cells(m.nstr,21).Value) + m.koplate)
    CASE nmonth = 10
     oExcel.Cells(m.nstr,23).Value  = m.koplate
     oExcel.Cells(m.nstr,26).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,26).Value), 0, oExcel.Cells(m.nstr,26).Value) + m.koplate
    CASE nmonth = 11
     oExcel.Cells(m.nstr,24).Value  = m.koplate
     oExcel.Cells(m.nstr,26).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,26).Value), 0, oExcel.Cells(m.nstr,26).Value) + m.koplate
    CASE nmonth = 12
     oExcel.Cells(m.nstr,25).Value  = m.koplate
     oExcel.Cells(m.nstr,26).Value   = m.kv04 - (IIF(ISNULL(oExcel.Cells(m.nstr,26).Value), 0, oExcel.Cells(m.nstr,26).Value) + m.koplate)
   ENDCASE 

   oExcel.Cells(m.nstr,5).Value   = IIF(ISNULL(oExcel.Cells(m.nstr,5).Value), 0, oExcel.Cells(m.nstr,5).Value) + m.koplate

   m.akoplate = IIF(!ISNULL(oExcel.Cells(m.nstr,5).Value), oExcel.Cells(m.nstr,5).Value, 0)
    
   IF m.akoplate>0 AND m.plan>0
    oExcel.Cells(m.nstr,6).Value = m.plan - m.akoplate
   ENDIF 
    
   m.nstr = m.nstr + 1
   m.npp = m.npp + 1
  ENDSCAN 

  USE IN aisoms
  USE IN pr4
  
  WAIT CLEAR 
 ENDFOR 
 
 
 
 WAIT CLEAR 

 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('pilot')
  USE IN pilot
 ENDIF 
 IF USED('pr4')
  USE IN pr4
 ENDIF 

 oDoc.SaveAs(DocName)
* oDoc.SaveAs(DocName, 0)
 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 

RETURN 