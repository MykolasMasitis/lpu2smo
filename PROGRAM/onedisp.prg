PROCEDURE OneDisp 

m.IsVisible = .T.
m.IsQuit    = .F.
m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)

DDDotName = pTempl + "\DDxxxxqqmmy.dot"
DSDotName = pTempl + "\DSxxxxqqmmy.dot"

TRY 
 oWord=GETOBJECT(,"Word.Application")
CATCH 
 oWord=CREATEOBJECT("Word.Application")
ENDTRY 

 m.mcod  = aisoms.mcod
 m.lpuid = aisoms.lpuid
 
 IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+m.mcod+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\talon.dbf')
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À TALON.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN  
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'cod')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\e'+m.mcod, 'serror', 'share', 'rid')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('serror')
   USE IN serror
  ENDIF 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\people', 'people', 'share', 'sn_pol')>0
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('serror')
   USE IN serror
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
 ENDIF 
 
 m.lIsDSDisp = .f.
 m.lIsDDDisp = .f.

 IF SEEK(1900, 'talon')
  m.lIsDDDisp = .t.
 ENDIF 
 IF m.lIsDDDisp = .f.
  IF SEEK(1901, 'talon')
   m.lIsDDDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDDDisp = .f.
  IF SEEK(1902, 'talon')
   m.lIsDDDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDDDisp = .f.
  IF SEEK(1903, 'talon')
   m.lIsDDDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDDDisp = .f.
  IF SEEK(1904, 'talon')
   m.lIsDDDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDDDisp = .f.
  IF SEEK(1905, 'talon')
   m.lIsDDDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDSDisp = .f.
  IF SEEK(101929, 'talon')
   m.lIsDSDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDSDisp = .f.
  IF SEEK(101930, 'talon')
   m.lIsDSDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDSDisp = .f.
  IF SEEK(101931, 'talon')
   m.lIsDSDisp = .t.
  ENDIF 
 ENDIF 
 IF m.lIsDSDisp = .f.
  IF SEEK(101932, 'talon')
   m.lIsDSDisp = .t.
  ENDIF 
 ENDIF 
  
 IF m.lIsDDDisp = .f. AND m.lIsDSDisp = .f.
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('serror')
   USE IN serror
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'¬€¡–¿ÕÕ€Ã Àœ” ƒ»—œ¿Õ—≈–»«¿÷»ﬂ Õ≈ œ–Œ¬Œƒ»À¿—‹!'+CHR(13)+CHR(10),0+64,'')
 RETURN  
  
 ENDIF 

 m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')

 lcpath = pbase+'\'+gcperiod+'\'+m.mcod
 DDDocName = lcpath + "\DD" + STR(m.lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 DSDocName = lcpath + "\DS" + STR(m.lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 SELECT talon 
 SET RELATION TO recid INTO serror
 SET RELATION TO sn_pol INTO people ADDITIVE 

 CREATE CURSOR ddcurs (cod n(6), k_u n(3), k_uok n(3), s_all n(11,2), s_allok n(11,2))
 INDEX on cod TAG cod 
 SET ORDER TO cod
  
 CREATE CURSOR dscurs (cod n(6), k_u n(3), k_uok n(3), s_all n(11,2), s_allok n(11,2))
 INDEX ON cod TAG cod 
 SET ORDER TO cod

 SELECT talon 
 
 SCAN 
  m.cod   = cod
  IF !INLIST(m.cod,1900,1901,1902,1903,1904,1905,101929,101930,101931,101932)
   LOOP 
  ENDIF 
  
  m.d_u   = d_u
  m.k_u   = k_u
  m.s_all = s_all
  m.key   = m.mcod + PADL(recid,6,'0')
  m.er    = ''

   m.id_smo = recid
   m.sn_pol = sn_pol
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.w      = people.w
   m.dr     = people.dr
   m.er     = serror.c_err
   
   m.ages   = YEAR(tdat1) - YEAR(m.dr)

   IF INLIST(m.cod, 1900,1901,1902,1903,1904,1905)
    IF SEEK(m.cod, 'ddcurs')
     m.ok_u     = ddcurs.k_u
     m.ok_uok   = ddcurs.k_uok
     m.os_all   = ddcurs.s_all
     m.os_allok = ddcurs.s_allok
     
     UPDATE ddcurs SET k_u = m.ok_u + m.k_u, ;
      k_uok = m.ok_uok + IIF(EMPTY(m.er), m.k_u, 0), ;
      s_all = m.os_all + m.s_all, ;
      s_allok = m.os_allok + IIF(EMPTY(m.er), m.s_all, 0) WHERE cod = m.cod 
      
    ELSE 
     INSERT INTO ddcurs (cod, k_u, k_uok, s_all, s_allok) VALUES ;
      (m.cod, m.k_u, IIF(EMPTY(m.er), m.k_u, 0), ;
       m.s_all, IIF(EMPTY(m.er),m.s_all,0))
    ENDIF 

   ELSE 

    IF SEEK(m.cod, 'dscurs')
     m.ok_u     = dscurs.k_u
     m.ok_uok   = dscurs.k_uok
     m.os_all   = dscurs.s_all
     m.os_allok = dscurs.s_allok
     
     UPDATE dscurs SET k_u = m.ok_u + m.k_u, ;
      k_uok = m.ok_uok + IIF(EMPTY(m.er), m.k_u, 0), ;
      s_all = m.os_all + m.s_all, ;
      s_allok = m.os_allok + IIF(EMPTY(m.er), m.s_all, 0) WHERE cod = m.cod 
      
    ELSE 
     INSERT INTO dscurs (cod, k_u, k_uok, s_all, s_allok) VALUES ;
      (m.cod, m.k_u, IIF(EMPTY(m.er), m.k_u, 0), ;
       m.s_all, IIF(EMPTY(m.er),m.s_all,0))
    ENDIF 

   ENDIF 
   
 ENDSCAN 
 SET RELATION OFF INTO serror
 SET RELATION OFF INTO people
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 
 SELECT ddcurs
 m.totpaz   = 0
 m.totsum   = 0
 m.totpazok = 0
 m.totsumok = 0
 IF RECCOUNT('ddcurs')=0
  USE 
 ELSE 
  IF fso.FileExists(DDDocName+'.doc')
   fso.DeleteFile(DDDocName+'.doc')
  ENDIF 

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
  CATCH 
   oDoc.SaveAs(DDDocName, 0)
  ENDTRY 

  USE IN ddcurs

 ENDIF 
 
 m.totpaz   = 0
 m.totsum   = 0
 m.totpazok = 0
 m.totsumok = 0
 SELECT dscurs
 IF RECCOUNT('dscurs')=0
  USE 
 ELSE 
  IF fso.FileExists(DSDocName+'.doc')
   fso.DeleteFile(DSDocName+'.doc')
  ENDIF 

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
  CATCH 
   oDoc.SaveAs(DSDocName, 0)
  ENDTRY 

  USE IN dscurs

 ENDIF 

 SELECT aisoms
 
IF USED('dsp')
 USE IN dsp
ENDIF 

IF IsVisible == .t. 
 oWord.Visible = .t.
ELSE 
 oDoc.Close(0)
 IF IsQuit
  oWord.Quit
 ENDIF 
ENDIF 
