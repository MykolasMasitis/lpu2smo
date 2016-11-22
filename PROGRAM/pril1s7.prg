PROCEDURE Pril1S7
 IF MESSAGEBOX('¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ œ–»ÀŒ∆≈Õ»≈ π1'+CHR(13)+CHR(10)+'(œ–» ¿« π725 ÓÚ 17.11.2014 „.)',4+32,'')=7
  RETURN 
 ENDIF 
 dotname = 'pril1s7'
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 IF !fso.FileExists(ptempl+'\'+dotname+'.dot')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ƒŒ ”Ã≈Õ“¿'+CHR(13)+CHR(10)+dotname+'.dot',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À SPRLPUXX.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF
 
 CREATE CURSOR curtbl (mcod c(7), lpuid n(4), lpuname c(120), kolst n(6), koldst n(6), kolamb n(6), s_all n(11,2))
 INDEX on mcod TAG mcod 
 SET ORDER TO mcod 
 
 SELECT aisoms
 SET RELATION TO mcod INTO sprlpu
 SCAN 
  m.mcod    = mcod
  m.lpuid   = lpuid
  m.lpuname = sprlpu.fullname
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   IF USED('err')
    USE IN err 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 

  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  CREATE CURSOR curamb (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR curdst (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  CREATE CURSOR curst (c_i c(30))
  INDEX on c_i TAG c_i
  SET ORDER TO c_i
  
  m.s_all = 0
  SELECT talon 
  SET RELATION TO recid INTO err
  SCAN 
   IF !EMPTY(err.rid)
    LOOP 
   ENDIF 
   m.cod    = cod
   m.sn_pol = sn_pol 
   m.c_i    = c_i 
   m.s_all  = m.s_all + s_all
   DO CASE 
    CASE IsUsl(m.cod)
     IF !SEEK(m.sn_pol, 'curamb')
      INSERT INTO curamb FROM MEMVAR 
     ENDIF 
    CASE IsKd(m.cod)
     IF !SEEK(m.sn_pol, 'curdst')
      INSERT INTO curdst FROM MEMVAR 
     ENDIF 
    CASE IsMes(m.cod) OR IsVMP(m.cod)
     IF !SEEK(m.c_i, 'curst')
      INSERT INTO curst FROM MEMVAR 
     ENDIF 
    OTHERWISE 
   ENDCASE 
  ENDSCAN 
  SET RELATION OFF INTO err 
  USE IN talon 
  USE IN err 
  
  m.kolst  = RECCOUNT('curst')
  m.koldst = RECCOUNT('curdst')
  m.kolamb = RECCOUNT('curamb')
  
  USE IN curst
  USE IN curdst
  USE IN curamb
  
  INSERT INTO curtbl FROM MEMVAR 

  SELECT aisoms 

 ENDSCAN 
 SET RELATION OFF INTO sprlpu
 USE IN aisoms
 USE IN sprlpu
 WAIT CLEAR 
 
 SELECT curtbl
 
 COPY TO &pout\&gcperiod\pril1s7
 GO TOP 

 m.IsVisible = .T.
 m.IsQuit    = .F.
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)

 WAIT "«¿œ”—  WORD..." WINDOW NOWAIT 
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 
 
 lcpath = pout+'\'+gcperiod
 IF !fso.FolderExists(lcpath)
  fso.CreateFolder(lcpath)
 ENDIF 
 docname = lcpath + "\pril1s7"
  
 IF fso.FileExists(docname+'.doc')
  fso.DeleteFile(docname+'.doc')
 ENDIF 

 oDoc = oWord.Documents.Add(ptempl+'\'+dotname)
 oDoc.Bookmarks('period').Select  
 oWord.Selection.TypeText(m.period+' „Ó‰')

 oDoc.Tables(1).Cell(4,2).Select
 oWord.Selection.TypeText(ALLTRIM(lpuname))
 oDoc.Tables(1).Cell(4,3).Select
 oWord.Selection.TypeText(TRANSFORM(kolst, '99999'))
 oDoc.Tables(1).Cell(4,4).Select
 oWord.Selection.TypeText(TRANSFORM(koldst, '99999'))
 oDoc.Tables(1).Cell(4,5).Select
 oWord.Selection.TypeText(TRANSFORM(kolamb, '99999'))
 oDoc.Tables(1).Cell(4,7).Select
 oWord.Selection.TypeText(TRANSFORM(s_all, '99999999.99'))
  
 m.totst  = 0
 m.totdst = 0
 m.totamb = 0
 m.totsum = 0

 m.totst  = m.totst  + kolst
 m.totdst = m.totdst + koldst
 m.totamb = m.totamb + kolamb
 m.totsum = m.totsum + s_all

 n=4
 SKIP 
 DO WHILE !EOF()
  oDoc.Tables(1).Cell(n,1).Select
  oWord.Selection.InsertRowsBelow
  n=n+1
  oDoc.Tables(1).Cell(n,2).Select
  oWord.Selection.TypeText(ALLTRIM(lpuname))
  oDoc.Tables(1).Cell(n,3).Select
  oWord.Selection.TypeText(TRANSFORM(kolst, '99999'))
  oDoc.Tables(1).Cell(n,4).Select
  oWord.Selection.TypeText(TRANSFORM(koldst, '99999'))
  oDoc.Tables(1).Cell(n,5).Select
  oWord.Selection.TypeText(TRANSFORM(kolamb, '99999'))
  oDoc.Tables(1).Cell(n,7).Select
  oWord.Selection.TypeText(TRANSFORM(s_all, '99999999.99'))

  m.totst  = m.totst  + kolst
  m.totdst = m.totdst + koldst
  m.totamb = m.totamb + kolamb
  m.totsum = m.totsum + s_all
   
  SKIP 

 ENDDO 

 oDoc.Bookmarks('totst').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totst, '999999'))
 oDoc.Bookmarks('totdst').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totdst, '999999'))
 oDoc.Bookmarks('totamb').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totamb, '999999'))
 oDoc.Bookmarks('totsum').Select  
 oWord.Selection.TypeText(TRANSFORM(m.totsum, '999999999.99'))

 oDoc.SaveAs(DocName+'.doc', 0)

 USE IN curtbl
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 

RETURN