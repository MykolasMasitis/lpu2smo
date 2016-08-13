PROCEDURE FormDDDS

 IF MESSAGEBOX(CHR(13)+CHR(10)+'—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“€ œŒ ƒ»—œ¿Õ—≈–»«¿÷»»?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 dspfile = pbase+'\'+gcperiod+'\'+'dsp'
 IF !fso.FileExists(dspfile+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'DSP-‘¿…À Õ≈ —‘Œ–Ã»–Œ¬¿Õ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod") > 0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN
 ENDIF 
 IF OpenFile(dspfile, 'dsp', 'shar', 'uniqq')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 
 DDDotName = pTempl + "\DDxxxxqqmmy.dot"
 DSDotName = pTempl + "\DSxxxxqqmmy.dot"
 IF !fso.FileExists(DDDotName)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ DDxxxxqqmmy.dot!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(DSDotName)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ DSxxxxqqmmy.dot!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 m.IsVisible = .F.
 m.IsQuit    = .F.
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)

 WAIT "«¿œ”—  WORD..." WINDOW NOWAIT 
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 
 
 SELECT mcod DISTINCT FROM dsp WHERE period=m.gcperiod INTO CURSOR curlpu
 
 SELECT curlpu 

 SCAN 
  m.mcod = mcod 
  WAIT m.mcod WINDOW NOWAIT 
  m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')

  lcpath = pbase+'\'+gcperiod+'\'+m.mcod
  DDDocName = lcpath + "\DD" + STR(m.lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  DSDocName = lcpath + "\DS" + STR(m.lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  
  IF fso.FileExists(DDDocName+'.doc')
   fso.DeleteFile(DDDocName+'.doc')
  ENDIF 
  IF fso.FileExists(DDDocName+'.pdf')
   fso.DeleteFile(DDDocName+'.pdf')
  ENDIF 
  IF fso.FileExists(DSDocName+'.doc')
   fso.DeleteFile(DSDocName+'.doc')
  ENDIF 
  IF fso.FileExists(DSDocName+'.pdf')
   fso.DeleteFile(DSDocName+'.pdf')
  ENDIF 

  CREATE CURSOR ddcurs (cod n(6), k_u n(3), k_uok n(3), s_all n(11,2), s_allok n(11,2))
  INDEX on cod TAG cod 
  SET ORDER TO cod
  
  CREATE CURSOR dscurs (cod n(6), k_u n(3), k_uok n(3), s_all n(11,2), s_allok n(11,2))
  INDEX ON cod TAG cod 
  SET ORDER TO cod

  SELECT dsp
  SCAN 
   SCATTER FIELDS EXCEPT recid,mcod,period MEMVAR 
   m.mmcod = mcod
   
   IF period!=m.gcperiod
    LOOP 
   ENDIF 

   IF m.mcod!=m.mmcod
    LOOP 
   ENDIF 
   
   IF m.cod<99999
    IF m.cod = 15001
     LOOP 
    ENDIF 

    IF SEEK(m.cod, 'ddcurs')
     m.ok_u     = ddcurs.k_u
     m.ok_uok   = ddcurs.k_uok
     m.os_all   = ddcurs.s_all
     m.os_allok = ddcurs.s_allok
     
     UPDATE ddcurs SET k_u = m.ok_u + 1, ;
      k_uok = m.ok_uok + IIF(EMPTY(m.er), 1, 0), ;
      s_all = m.os_all + m.s_all, ;
      s_allok = m.os_allok + IIF(EMPTY(m.er), m.s_all, 0) WHERE cod = m.cod 
      
    ELSE 
     INSERT INTO ddcurs (cod, k_u, k_uok, s_all, s_allok) VALUES ;
      (m.cod, 1, IIF(EMPTY(m.er), 1, 0), ;
       m.s_all, IIF(EMPTY(m.er),m.s_all,0))
    ENDIF 

   ELSE 

    IF SEEK(m.cod, 'dscurs')
     m.ok_u     = dscurs.k_u
     m.ok_uok   = dscurs.k_uok
     m.os_all   = dscurs.s_all
     m.os_allok = dscurs.s_allok
     
     UPDATE dscurs SET k_u = m.ok_u + 1, ;
      k_uok = m.ok_uok + IIF(EMPTY(m.er), 1, 0), ;
      s_all = m.os_all + m.s_all, ;
      s_allok = m.os_allok + IIF(EMPTY(m.er), m.s_all, 0) WHERE cod = m.cod 
      
    ELSE 
     INSERT INTO dscurs (cod, k_u, k_uok, s_all, s_allok) VALUES ;
      (m.cod, 1, IIF(EMPTY(m.er), 1, 0), ;
       m.s_all, IIF(EMPTY(m.er),m.s_all,0))
    ENDIF 

   ENDIF 
   
  ENDSCAN 
  
  IF RECCOUNT('ddcurs')>0
   =MakeDDList()
  ENDIF 
  USE IN ddcurs

  IF RECCOUNT('dscurs')>0
   =MakeDSList()
  ENDIF 
  USE IN dscurs
  
  SELECT curlpu
  
  WAIT CLEAR 

 ENDSCAN 

 USE IN sprlpu
 USE IN dsp
 USE IN curlpu

 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')

RETURN 

FUNCTION MakeDDList

 SELECT ddcurs
 m.totpaz   = 0
 m.totsum   = 0
 m.totpazok = 0
 m.totsumok = 0

 GO TOP 
  
 oDoc = oWord.Documents.Add(dddotname)
 oDoc.Bookmarks('qname').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('lpuname').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('period').Select  
 oWord.Selection.TypeText(m.period+' „Ó‰')
  
 oDoc.Tables(1).Cell(4,1).Select
 oWord.Selection.TypeText(PADL(cod,6))
 oDoc.Tables(1).Cell(4,3).Select
 oWord.Selection.TypeText(PADL(k_u,3))
 oDoc.Tables(1).Cell(4,4).Select
 oWord.Selection.TypeText(TRANSFORM(s_all,'9999999.99'))
 oDoc.Tables(1).Cell(4,5).Select
 oWord.Selection.TypeText(PADL(k_uok,3))
 oDoc.Tables(1).Cell(4,6).Select
 oWord.Selection.TypeText(TRANSFORM(s_allok,'9999999.99'))

 m.totpaz   = m.totpaz   + k_u
 m.totsum   = m.totsum   + s_all
 m.totpazok = m.totpazok + k_uok
 m.totsumok = m.totsumok + s_allok

 n=4
 SKIP 
 DO WHILE !EOF()
  oDoc.Tables(1).Cell(n,1).Select
  oWord.Selection.InsertRowsBelow
  n=n+1
  oDoc.Tables(1).Cell(n,1).Select
  oWord.Selection.TypeText(PADL(cod,6))
  oDoc.Tables(1).Cell(n,3).Select
  oWord.Selection.TypeText(PADL(k_u,3))
  oDoc.Tables(1).Cell(n,4).Select
  oWord.Selection.TypeText(TRANSFORM(s_all,'9999999.99'))
  oDoc.Tables(1).Cell(n,5).Select
  oWord.Selection.TypeText(PADL(k_uok,3))
  oDoc.Tables(1).Cell(n,6).Select
  oWord.Selection.TypeText(TRANSFORM(s_allok,'9999999.99'))

  m.totpaz   = m.totpaz   + k_u
  m.totsum   = m.totsum   + s_all
  m.totpazok = m.totpazok + k_uok
  m.totsumok = m.totsumok + s_allok

  SKIP 

 ENDDO 
  
 oDoc.Bookmarks('totpaz').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totpaz, '9999'))
 oDoc.Bookmarks('totsum').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totsum, '9999999.99'))
 oDoc.Bookmarks('totpazok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totpazok, '9999'))
 oDoc.Bookmarks('totsumok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totsumok, '9999999.99'))

 TRY 
  oDoc.SaveAs(DDDocName, 17)
  oDoc.Close(0)
 CATCH 
  oDoc.SaveAs(DDDocName, 0)
  oDoc.Close(0)
 ENDTRY 

RETURN 

FUNCTION MakeDSList

 SELECT dscurs

 m.totpaz   = 0
 m.totsum   = 0
 m.totpazok = 0
 m.totsumok = 0

 GO TOP 
  
 oDoc = oWord.Documents.Add(dsdotname)
 oDoc.Bookmarks('qname').Select  
 oWord.Selection.TypeText(m.qname)
 oDoc.Bookmarks('lpuname').Select  
 oWord.Selection.TypeText(m.lpuname)
 oDoc.Bookmarks('period').Select  
 oWord.Selection.TypeText(m.period+' „Ó‰')

 oDoc.Tables(1).Cell(4,1).Select
 oWord.Selection.TypeText(PADL(cod,6))
 oDoc.Tables(1).Cell(4,3).Select
 oWord.Selection.TypeText(PADL(k_u,3))
 oDoc.Tables(1).Cell(4,4).Select
 oWord.Selection.TypeText(TRANSFORM(s_all,'9999999.99'))
 oDoc.Tables(1).Cell(4,5).Select
 oWord.Selection.TypeText(PADL(k_uok,3))
 oDoc.Tables(1).Cell(4,6).Select
 oWord.Selection.TypeText(TRANSFORM(s_allok,'9999999.99'))
  
 m.totpaz   = m.totpaz   + k_u
 m.totsum   = m.totsum   + s_all
 m.totpazok = m.totpazok + k_uok
 m.totsumok = m.totsumok + s_allok

 n=4
 SKIP 
 DO WHILE !EOF()
  oDoc.Tables(1).Cell(n,1).Select
  oWord.Selection.InsertRowsBelow
  n=n+1
  oDoc.Tables(1).Cell(n,1).Select
  oWord.Selection.TypeText(PADL(cod,6))
  oDoc.Tables(1).Cell(n,3).Select
  oWord.Selection.TypeText(PADL(k_u,3))
  oDoc.Tables(1).Cell(n,4).Select
  oWord.Selection.TypeText(TRANSFORM(s_all,'9999999.99'))
  oDoc.Tables(1).Cell(n,5).Select
  oWord.Selection.TypeText(PADL(k_uok,3))
  oDoc.Tables(1).Cell(n,6).Select
  oWord.Selection.TypeText(TRANSFORM(s_allok,'9999999.99'))

  m.totpaz   = m.totpaz   + k_u
  m.totsum   = m.totsum   + s_all
  m.totpazok = m.totpazok + k_uok
  m.totsumok = m.totsumok + s_allok
   
  SKIP 

 ENDDO 

 oDoc.Bookmarks('totpaz').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totpaz, '9999'))
 oDoc.Bookmarks('totsum').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totsum, '9999999.99'))
 oDoc.Bookmarks('totpazok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totpazok, '9999'))
 oDoc.Bookmarks('totsumok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totsumok, '9999999.99'))

 TRY 
  oDoc.SaveAs(DSDocName, 17)
  oDoc.Close(0)
 CATCH 
  oDoc.SaveAs(DSDocName, 0)
  oDoc.Close(0)
 ENDTRY 

RETURN 