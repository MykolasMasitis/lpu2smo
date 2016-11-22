PROCEDURE LoadResults
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÇÀÃÐÓÇÈÒÜ ÐÅÇÓËÜÒÀÒÛ ÝÊÑÏÅÐÒÈÇ?',4+32,'')=7
  RETURN 
 ENDIF 

 LoadDir = fso.GetParentFolderName(pbin)+'\UPLOAD'
 IF !fso.FolderExists(LoadDir)
  fso.CreateFolder(LoadDir)
 ENDIF 

 oal = SYS(5)+SYS(2003)
 SET DEFAULT TO (LoadDir)
 XLSXFile = GETFILE('xls','','',0,'Óêàæèòå íà ôàéë!')
 SET DEFAULT TO (oal)
 m.dbfrep = STRTRAN(LOWER(XLSXFile),'.xls','')
 
 IF EMPTY(XLSXFile)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÍÈ×ÅÃÎ ÍÅ ÂÛÁÐÀËÈ!'+CHR(13)+CHR(10),0+48,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\nsi\sookodxx', 'sookod', 'shar', 'osn230')>0
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 

 WAIT "ÇÀÏÓÑÊ EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 
 
 TRY 
  WAIT "ÇÀÃÐÓÇÊÀ ÔÀÉËÀ..." WINDOW NOWAIT 
  oDoc = oExcel.WorkBooks.Open(XLSXFile,.T.)
  WAIT CLEAR 
 CATCH 
  oExcel.quit
  MESSAGEBOX('ÍÅ ÓÄÀËÎÑÜ ÎÒÊÐÛÒÜ ÂÛÁÐÀÍÍÛÉ ÔÀÉË!',0+64,'')
  RETURN 
 ENDTRY 

 nSheets = oExcel.Sheets.Count
 
 FOR nSheet=1 TO nSheets
  oExcel.Sheets(nSheet).Select
  m.lstname = LOWER(ALLTRIM(oExcel.Sheets(nSheet).Name))
  IF m.lstname='ñ÷åò'
   MESSAGEBOX(oExcel.Sheets(nSheet).Name,0+64,'')
 
   CREATE CURSOR onesheet (recid n(6), mcod char(7), period char(6), ;
   	sn_pol char(25), c_i char(25), cod n(6), ds char(6), k_u n(3), tip c(1), s_all n(11,2), ;
   	et c(1), d_exp d, err_mee c(3), osn230 c(5), e_period c(6), koeff n(4,2), straf n(4,2), ;
   	docexp c(7), s_1 n(11,2), s_2 n(11,2))
 
   nRow = 4
   DO WHILE !ISNULL(oexcel.Cells(nRow,1).Value)
 
    m.mcod   = IIF(!ISNULL(oexcel.Cells(nRow,2).Value), ALLTRIM(oexcel.Cells(nRow,2).Value), '')
    m.sn_pol = IIF(!ISNULL(oexcel.Cells(nRow,8).Value), ALLTRIM(oexcel.Cells(nRow,8).Value), '')
    m.c_i    = IIF(!ISNULL(oexcel.Cells(nRow,10).Value), ALLTRIM(oexcel.Cells(nRow,10).Value), '')
    m.cod    = IIF(!ISNULL(oexcel.Cells(nRow,12).Value), INT(VAL(ALLTRIM(oexcel.Cells(nRow,12).Value))), 0)
    m.ds     = IIF(!ISNULL(oexcel.Cells(nRow,14).Value), ALLTRIM(oexcel.Cells(nRow,14).Value), '')
    m.k_u    = IIF(!ISNULL(oexcel.Cells(nRow,25).Value), IIF(VARTYPE(oexcel.Cells(nRow,25).Value)!='N', INT(VAL(oexcel.Cells(nRow,25).Value)),oexcel.Cells(nRow,25).Value ), 0)
    m.tip    = IIF(!ISNULL(oexcel.Cells(nRow,27).Value), ALLTRIM(oexcel.Cells(nRow,27).Value), '')
    m.s_all  = IIF(!ISNULL(oexcel.Cells(nRow,28).Value), IIF(VARTYPE(oexcel.Cells(nRow,28).Value)='N', oexcel.Cells(nRow,28).Value, VAL(oexcel.Cells(nRow,28).Value)), 0)
    m.period = IIF(!ISNULL(oexcel.Cells(nRow,52).Value), ALLTRIM(oexcel.Cells(nRow,52).Value), '')
    m.et     = IIF(!ISNULL(oexcel.Cells(nRow,67).Value), ALLTRIM(oexcel.Cells(nRow,67).Value), '')
    m.d_exp  = IIF(!ISNULL(oexcel.Cells(nRow,42).Value), ;
    	IIF(INLIST(VARTYPE(oexcel.Cells(nRow,42).Value),'T','D'), oexcel.Cells(nRow,42).Value, {}),{})
    m.e_period = LEFT(DTOS(m.d_exp),6)

    m.osn230 = IIF(ALLTRIM(oexcel.Cells(nRow,68).Value)='1', ALLTRIM(oexcel.Cells(nRow,29).Value), '0.0.0')
    m.err_mee = IIF(SEEK(m.osn230, 'sookod'), sookod.er_c, '')

    m.koeff  = IIF(ALLTRIM(oexcel.Cells(nRow,68).Value)='1', IIF(VARTYPE(oexcel.Cells(nRow,30).Value)='N',;
    	oexcel.Cells(nRow,30).Value, 0), 0)
*    m.s_1    = IIF(ALLTRIM(oexcel.Cells(nRow,68).Value)='1', oexcel.Cells(nRow,31).Value, 0)
    m.s_1    = IIF(ALLTRIM(oexcel.Cells(nRow,68).Value)='1', ;
    	IIF(!ISNULL(oexcel.Cells(nRow,31).Value), IIF(VARTYPE(oexcel.Cells(nRow,31).Value)='N', oexcel.Cells(nRow,31).Value, VAL(oexcel.Cells(nRow,31).Value)), 0),;
    	0)
    m.s_2    = 0

    m.docexp = IIF(!ISNULL(oexcel.Cells(nRow,46).Value) and VARTYPE(oexcel.Cells(nRow,46).Value)='N', ;
    	ALLTRIM(STR(oexcel.Cells(nRow,46).Value)), '')

    INSERT INTO onesheet FROM MEMVAR 
 
    nRow = nRow + 1
   ENDDO 
  ENDIF 
 ENDFOR 

 oExcel.Quit
 
 USE IN sookod

 SELECT onesheet
* MESSAGEBOX(m.dbfrep,0+64,'')
* COPY TO &dbfrep
 
 SELECT * FROM onesheet WHERE !EMPTY(period) AND !EMPTY(mcod) ORDER BY period,mcod INTO CURSOR crs READWRITE 
 
 SELECT crs
 m.period = 'qw'
 m.mcod = 'we'

 SCAN 
  IF m.mcod!=mcod AND m.period!=period
   m.period = ALLTRIM(period)
   m.mcod   = mcod
   m.pfile  = m.pbase+'\'+m.period+'\'+m.mcod+'\talon'
   IF USED('talon')
    SELECT talon 
    SET ORDER TO 
    DELETE TAG qwert
    USE IN talon 
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 

   IF !fso.FileExists(m.pbase+'\'+m.period+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.pbase+'\'+m.period+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 

   IF OpenFile(m.pbase+'\'+m.period+'\'+m.mcod+'\talon', 'talon', 'excl')>0
    IF USED('talon')
     USE IN talon 
    ENDIF 
    SELECT crs
    LOOP 
   ELSE 
    SELECT talon 
*    INDEX on c_i+PADL(cod,6,'0')+ds+PADL(k_u,3,'0')+tip TAG qwert 
    INDEX on sn_pol+PADL(cod,6,'0')+ds+tip TAG qwert 
    SET ORDER TO qwert
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.period+'\'+m.mcod+'\m'+m.mcod, 'merror', 'excl')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    USE IN talon 
    SELECT crs
    LOOP 
   ENDIF 
  ELSE 
   IF !USED('talon')
    LOOP 
   ENDIF 
   IF !USED('merror')
    LOOP 
   ENDIF 
  ENDIF && IF m.mcod!=mcod AND m.period!=period
  
  SELECT crs
*  m.virr  = c_i+PADL(cod,6,'0')+ds+PADL(k_u,3,'0')+tip
  m.virr  = sn_pol+PADL(cod,6,'0')+ds+tip
  m.recid = IIF(SEEK(m.virr, 'talon', 'qwert'), talon.recid, 0)
  REPLACE recid WITH m.recid 
  
 ENDSCAN 

 IF USED('talon')
  USE IN talon 
 ENDIF 
 IF USED('merror')
  USE IN merror
 ENDIF 

 SELECT crs 
 COPY TO &dbfrep
 MESSAGEBOX(m.dbfrep,0+64,'')

RETURN 
