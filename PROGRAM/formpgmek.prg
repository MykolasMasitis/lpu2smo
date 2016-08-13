PROCEDURE FormPGMEK

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
  'ФОРМУ ПГ ПО МЭК?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 m.pgdat1 = m.tdat1
 m.pgdat2 = m.tdat2
 m.ischecked = .f.
 DO FORM SelPeriod
 IF m.ischecked = .f.
  RETURN 
 ENDIF 

 PUBLIC oExcel AS Excel.Application
 WAIT "Запуск MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 oExcel.Cells.Font.Name='Calibri'
 nSheet = 1

 BookName = pOut+'\PgMEK'
 IF fso.FileExists(BookName+'.xls')
  fso.DeleteFile(BookName+'.xls')
 ENDIF 

 =MakeHeadOfPage()

 m.svamb       = 0
 m.sumsvamb    = 0
 m.svamb_ok    = 0
 m.sumsvamb_ok = 0
 m.svamb_bad   = 0
 m.sumsvamb_bad = 0
 m.svamb_erz   = 0
 m.svamb_uma   = 0 && Услуга в во время госпитализации
 m.svamb_others  = 0 
 m.sumsvamb_erz   = 0
 m.sumsvamb_uma   = 0 && Услуга в во время госпитализации
 m.sumsvamb_others  = 0 

 m.svstac      = 0
 m.sumsvstac   = 0
 m.svstac_ok   = 0
 m.sumsvstac_ok = 0
 m.svstac_bad  = 0
 m.sumsvstac_bad  = 0
 m.svstac_erz  = 0
 m.svstac_others  = 0 
 m.sumsvstac_erz  = 0
 m.sumsvstac_others  = 0 

 m.svdstac     = 0
 m.sumsvdstac  = 0
 m.svdstac_ok  = 0
 m.sumsvdstac_ok = 0
 m.svdstac_bad = 0
 m.sumsvdstac_bad = 0
 m.svdstac_erz = 0
 m.svdstac_others  = 0 
 m.sumsvdstac_erz = 0
 m.sumsvdstac_others  = 0 

 m.curdat = m.pgdat1-1
 m.curmonth = 0
 DO WHILE  m.curdat<m.pgdat2
  m.curdat = m.curdat + 1
  IF MONTH(m.curdat)!=m.curmonth
   m.curmonth = MONTH(m.curdat)
  ELSE 
   LOOP 
  ENDIF 
  
  lcperiod =  STR(YEAR(m.curdat),4)+PADL(m.curmonth,2,'0')
  IF !fso.FolderExists(pbase+'\'+lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'Tarif', 'SHARED', 'cod') > 0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   LOOP 
  ENDIF 
  
  WAIT lcperiod WINDOW NOWAIT 

  SELECT aisoms
  SCAN 
   m.mcod = mcod
   m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  
   IF !fso.FolderExists(pBase+'\'+gcPeriod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\people.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
  
   tnresult = 0
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people', 'people', 'shar')
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'serrors', 'shar', 'rid')
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'rerrors', 'shar', 'rrid', 'again')
  
   IF tnresult>0
    IF USED('rerrors')
     USE IN rerrors
    ENDIF 
    IF USED('serrors')
     USE IN serrors
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   
   SELECT people 
   SET ORDER TO sn_pol
   SET RELATION TO recid INTO rerrors
  
   SELECT talon 
   SET RELATION TO sn_pol INTO people 
   SET RELATION TO recid INTO serrors ADDITIVE 
  
   m.amb       = 0
   m.sumamb    = 0
   m.amb_ok    = 0
   m.sumamb_ok = 0
   m.amb_bad   = 0
   m.sumamb_bad   = 0
   m.amb_erz   = 0
   m.amb_uma   = 0 && Услуга в во время госпитализации
   m.amb_others  = 0 
   m.sumamb_erz   = 0
   m.sumamb_uma   = 0 && Услуга в во время госпитализации
   m.sumamb_others  = 0 

   m.stac       = 0
   m.sumstac    = 0
   m.stac_ok    = 0
   m.sumstac_ok = 0
   m.stac_bad  = 0
   m.sumstac_bad  = 0
   m.stac_erz  = 0
   m.stac_others  = 0 
   m.sumstac_erz  = 0
   m.sumstac_others  = 0 

   m.dstac       = 0
   m.sumdstac    = 0
   m.dstac_ok    = 0
   m.sumdstac_ok = 0
   m.dstac_bad = 0
   m.sumdstac_bad = 0
   m.dstac_erz = 0
   m.dstac_others  = 0 
   m.sumdstac_erz = 0
   m.sumdstac_others  = 0 

   CREATE CURSOR difcards (tip n(1), sn_pol c(25))
   INDEX FOR tip=1 ON sn_pol TAG amb
   INDEX FOR tip=2 ON sn_pol TAG st
   INDEX FOR tip=3 ON sn_pol TAG dg
  
   CREATE CURSOR difcardsbad (tip n(1), sn_pol c(25))
   INDEX FOR tip=1 ON sn_pol TAG amb
   INDEX FOR tip=2 ON sn_pol TAG st
   INDEX FOR tip=3 ON sn_pol TAG dg

   SELECT talon 

   SCAN
    m.cod    = cod
    m.sn_pol = sn_pol
    m.c_i    = c_i
    DO CASE 
     CASE IsPlk(m.cod)
      m.sumamb = m.sumamb + s_all
      IF !SEEK(m.sn_pol, 'difcards', 'amb')
       m.amb = m.amb + 1
       INSERT INTO difcards (tip, sn_pol) VALUES (1, m.sn_pol)
      ENDIF 
      IF !EMPTY(serrors.c_err) && Если МЭК
       m.sumamb_bad = m.sumamb_bad + s_all
       DO CASE 
        CASE INLIST(rerrors.c_err,'ERA','ECA')
         m.sumamb_erz = m.sumamb_erz + s_all
        CASE INLIST(serrors.c_err, 'UMA')
         m.sumamb_uma = m.sumamb_uma + s_all
        OTHERWISE 
         m.sumamb_others = m.sumamb_others + s_all
       ENDCASE 
       IF !SEEK(m.sn_pol, 'difcardsbad', 'amb')
        m.amb_bad = m.amb_bad + 1
        INSERT INTO difcardsbad (tip, sn_pol) VALUES (1, m.sn_pol)
        DO CASE 
         CASE INLIST(rerrors.c_err,'ERA','ECA')
          m.amb_erz = m.amb_erz + 1
         CASE INLIST(serrors.c_err, 'UMA')
          m.amb_uma = m.amb_uma + 1
         OTHERWISE 
          m.amb_others = m.amb_others + 1
        ENDCASE 
       ENDIF 
      ENDIF 

     CASE IsGsp(m.cod) OR IsVMP(m.cod)
      m.sumstac = m.sumstac + s_all
      IF !SEEK(m.c_i, 'difcards', 'st')
       m.stac = m.stac + 1
       INSERT INTO difcards (tip, sn_pol) VALUES (2, m.c_i)
      ENDIF 
      IF !EMPTY(serrors.c_err)
       m.sumstac_bad = m.sumstac_bad + s_all
       DO CASE 
        CASE INLIST(rerrors.c_err,'ERA','ECA')
         m.sumstac_erz = m.sumstac_erz + s_all
        OTHERWISE 
         m.sumstac_others = m.sumstac_others + s_all
       ENDCASE 
       IF !SEEK(m.c_i, 'difcardsbad', 'st')
        m.stac_bad = m.stac_bad + 1
        INSERT INTO difcardsbad (tip, sn_pol) VALUES (2, m.c_i)
        DO CASE 
         CASE INLIST(rerrors.c_err,'ERA','ECA')
          m.stac_erz = m.stac_erz + 1
         OTHERWISE 
          m.stac_others = m.stac_others + 1
        ENDCASE 
       ENDIF 
      ENDIF 

     CASE IsDst(m.cod)
      m.sumdstac = m.sumdstac + s_all
      IF !SEEK(m.sn_pol, 'difcards', 'dg')
       m.dstac = m.dstac + 1
       INSERT INTO difcards (tip, sn_pol) VALUES (3, m.sn_pol)
      ENDIF 
      IF !EMPTY(serrors.c_err)
       m.sumdstac_bad = m.sumdstac_bad + s_all
       DO CASE 
        CASE INLIST(rerrors.c_err,'ERA','ECA')
         m.sumdstac_erz = m.sumdstac_erz + s_all
        OTHERWISE 
         m.sumdstac_others = m.sumdstac_others + s_all
       ENDCASE 
       IF !SEEK(m.sn_pol, 'difcardsbad', 'dg')
        m.dstac_bad = m.dstac_bad + 1
        INSERT INTO difcardsbad (tip, sn_pol) VALUES (3, m.sn_pol)
        DO CASE 
         CASE INLIST(rerrors.c_err,'ERA','ECA')
          m.dstac_erz = m.dstac_erz + 1
         OTHERWISE 
          m.dstac_others = m.dstac_others + 1
        ENDCASE 
       ENDIF 
      ENDIF 
     OTHERWISE 
    ENDCASE 
   ENDSCAN 

   m.amb_ok   = m.amb   - m.amb_bad
   m.stac_ok  = m.stac  - m.stac_bad
   m.dstac_ok = m.dstac - m.dstac_bad

   m.sumamb_ok   = m.sumamb   - m.sumamb_bad
   m.sumstac_ok  = m.sumstac  - m.sumstac_bad
   m.sumdstac_ok = m.sumdstac - m.sumdstac_bad

   m.svamb        = m.svamb + m.amb
   m.sumsvamb     = m.sumsvamb + m.sumamb
   m.svamb_ok     = m.svamb_ok + m.amb_ok
   m.sumsvamb_ok  = m.sumsvamb_ok + m.sumamb_ok
   m.svamb_bad    = m.svamb_bad + m.amb_bad
   m.sumsvamb_bad = m.sumsvamb_bad + m.sumamb_bad
   m.svamb_erz    = m.svamb_erz + m.amb_erz
   m.svamb_uma    = m.svamb_uma + m.amb_uma
   m.svamb_others = m.svamb_others + m.amb_others
   m.sumsvamb_erz    = m.sumsvamb_erz + m.sumamb_erz
   m.sumsvamb_uma    = m.sumsvamb_uma + m.sumamb_uma
   m.sumsvamb_others = m.sumsvamb_others + m.sumamb_others

   m.svstac         = m.svstac + m.stac
   m.sumsvstac      = m.sumsvstac + m.sumstac
   m.svstac_ok      = m.svstac_ok + m.stac_ok
   m.sumsvstac_ok   = m.sumsvstac_ok + m.sumstac_ok
   m.svstac_bad     = m.svstac_bad + m.stac_bad
   m.sumsvstac_bad  = m.sumsvstac_bad + m.sumstac_bad
   m.svstac_erz     = m.svstac_erz +  m.stac_erz
   m.svstac_others  = m.svstac_others + m.stac_others
   m.sumsvstac_erz     = m.sumsvstac_erz +  m.sumstac_erz
   m.sumsvstac_others  = m.sumsvstac_others + m.sumstac_others

   m.svdstac         = m.svdstac + m.dstac
   m.sumsvdstac      = m.sumsvdstac + m.sumdstac
   m.svdstac_ok      = m.svdstac_ok + m.dstac_ok
   m.sumsvdstac_ok   = m.sumsvdstac_ok + m.sumdstac_ok
   m.svdstac_bad     = m.svdstac_bad + m.dstac_bad
   m.sumsvdstac_bad  = m.sumsvdstac_bad + m.sumdstac_bad
   m.svdstac_erz     = m.svdstac_erz + m.dstac_erz
   m.svdstac_others  = m.svdstac_others + m.dstac_others
   m.sumsvdstac_erz     = m.sumsvdstac_erz + m.sumdstac_erz
   m.sumsvdstac_others  = m.sumsvdstac_others + m.sumdstac_others

   SET RELATION OFF INTO serrors
   SET RELATION OFF INTO people
   USE IN serrors
   USE 
   SELECT people 
   SET RELATION OFF INTO rerrors 
   USE IN rerrors
   USE 
   
   IF USED('rerrors')
    USE IN rerrors
   ENDIF 
   IF USED('serrors')
    USE IN serrors
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
  ENDSCAN 
  USE 
  USE IN tarif
  
  USE IN difcards
  USE IN difcardsbad
  
  WAIT CLEAR 
 
 ENDDO 

 WITH oExcel.ActiveSheet 
  .Cells(03,3)  = m.svamb                && Амбулаторно-поликлиническая помощь
  .Cells(03,4)  = m.svstac               && Стационарная помощь
  .Cells(03,5)  = m.svdstac              && Стационар-замещающая помощь
  .Cells(03,6)  = m.svamb+m.svstac+m.svdstac && Всего
  .Cells(03,7)  = m.sumsvamb                && Амбулаторно-поликлиническая помощь
  .Cells(03,8)  = m.sumsvstac               && Стационарная помощь
  .Cells(03,9)  = m.sumsvdstac              && Стационар-замещающая помощь
  .Cells(03,10) = m.sumsvamb+m.sumsvstac+m.sumsvdstac && Всего

  .Cells(04,3) = m.svamb_bad 
  .Cells(04,4) = m.svstac_bad
  .Cells(04,5) = m.svdstac_bad
  .Cells(04,6) = m.svamb_bad+m.svstac_bad+m.svdstac_bad
  .Cells(04,7) = m.sumsvamb_bad 
  .Cells(04,8) = m.sumsvstac_bad
  .Cells(04,9) = m.sumsvdstac_bad
  .Cells(04,10) = m.sumsvamb_bad+m.sumsvstac_bad+m.sumsvdstac_bad

  .Cells(05,3) = m.svamb_bad 
  .Cells(05,4) = m.svstac_bad
  .Cells(05,5) = m.svdstac_bad
  .Cells(05,6) = m.svamb_bad+m.svstac_bad+m.svdstac_bad
  .Cells(05,7) = m.sumsvamb_bad 
  .Cells(05,8) = m.sumsvstac_bad
  .Cells(05,9) = m.sumsvdstac_bad
  .Cells(05,10) = m.sumsvamb_bad+m.sumsvstac_bad+m.sumsvdstac_bad

  .Cells(06,3) = m.svamb_others
  .Cells(06,4) = m.svstac_others
  .Cells(06,5) = m.svdstac_others
  .Cells(06,6) = m.svamb_others+m.svstac_others+m.svdstac_others
  .Cells(06,7) = m.sumsvamb_others
  .Cells(06,8) = m.sumsvstac_others
  .Cells(06,9) = m.sumsvdstac_others
  .Cells(06,10) = m.sumsvamb_others+m.sumsvstac_others+m.sumsvdstac_others

  .Cells(07,3) = m.svamb_erz 
  .Cells(07,4) = m.svstac_erz
  .Cells(07,5) = m.svdstac_erz
  .Cells(07,6) = m.svamb_erz+m.svstac_erz+m.svdstac_erz
  .Cells(07,7) = m.sumsvamb_erz 
  .Cells(07,8) = m.sumsvstac_erz
  .Cells(07,9) = m.sumsvdstac_erz
  .Cells(07,10) = m.sumsvamb_erz+m.sumsvstac_erz+m.sumsvdstac_erz

  .Cells(08,3) = 0 
  .Cells(08,4) = 0
  .Cells(08,5) = 0
  .Cells(08,6) = 0

  .Cells(09,3) = 0 
  .Cells(09,4) = 0
  .Cells(09,5) = 0
  .Cells(09,6) = 0

  .Cells(10,3) = 0 
  .Cells(10,4) = 0
  .Cells(10,5) = 0
  .Cells(10,6) = 0

  .Cells(12,3) = m.svamb_uma
  .Cells(12,4) = 0
  .Cells(12,5) = 0
  .Cells(12,6) = m.svamb_uma
  .Cells(12,7) = m.sumsvamb_uma
  .Cells(12,8) = 0
  .Cells(12,9) = 0
  .Cells(12,10) = m.sumsvamb_uma

  .Cells(11,3) = m.svamb_uma
  .Cells(11,4) = 0
  .Cells(11,5) = 0
  .Cells(11,6) = m.svamb_uma
  .Cells(11,7) = m.sumsvamb_uma
  .Cells(11,8) = 0
  .Cells(11,9) = 0
  .Cells(11,10) = m.sumsvamb_uma

  .Cells(16,3) = m.svamb_ok
  .Cells(16,4) = m.svstac_ok
  .Cells(16,5) = m.svdstac_ok
  .Cells(16,6) = m.svamb_ok+m.svstac_ok+m.svdstac_ok
  .Cells(16,7) = m.sumsvamb_ok
  .Cells(16,8) = m.sumsvstac_ok
  .Cells(16,9) = m.sumsvdstac_ok
  .Cells(16,10) = m.sumsvamb_ok+m.sumsvstac_ok+m.sumsvdstac_ok

  .Cells(17,3) = m.svamb
  .Cells(17,4) = m.svstac
  .Cells(17,5) = m.svdstac
  .Cells(17,6) = m.svamb+m.svstac+m.svdstac
  .Cells(17,7) = m.sumsvamb
  .Cells(17,8) = m.sumsvstac
  .Cells(17,9) = m.sumsvdstac
  .Cells(17,10) = m.sumsvamb+m.sumsvstac+m.sumsvdstac

  .Name = 'Сводка'
 ENDWITH 

 FOR iii=2 TO 10
  oexcel.Columns(iii).AutoFit
 ENDFOR 

 oBook.SaveAs(BookName)
 oExcel.Visible = .t.
 
RETURN 
 

FUNCTION MakeHeadOfPage
  oExcel.Columns(1).NumberFormat='@'
  oExcel.Columns(2).NumberFormat='@'

  WITH oExcel.ActiveSheet 
   .Cells(01,01) = 'Таблица 3.1'
  
   .Cells(03,01) = 'Количество предъявленных к оплате счетов за '+;
    'оказанную медицинскую помощь по территориальной программе ОМС'
   .Cells(03,2) = '1'

   .Cells(04,01) = 'Всего выявлено счетов, содержащих нарушения'
   .Cells(04,2) = '2'

   .Cells(05,01) = 'Выявлено нарушений в оформлении и предъявлении '+;
    'на оплату счетов и реестров счетов, в т.ч.:'
   .Cells(05,2) = '3'

   .Cells(06,01) = 'нарушения, связанные с оформлением счетов и реестров счетов'
   .Cells(06,2) = '3.1'

   .Cells(07,01) = 'нарушения, связанные с принадлежностью застрахованного лица к СМО'
   .Cells(07,2) = '3.2'

   .Cells(08,01) = 'нарушения, связанные с включением в реестр медицинской помощи, '+;
    'не входящей в территориальную программу ОМС' 
   .Cells(08,2) = '3.3'

   .Cells(09,01) = 'нарушения, связанные с необоснованным применением тарифа '+;
    'на медицинскую помощь'
   .Cells(09,2) = '3.4'

   .Cells(10,01) = 'нарушения, связанные с включением в реестр счетов нелицензированных'+;
    ' видов медицинской деятельности'
   .Cells(10,2) = '3.5'

   .Cells(11,01) = 'нарушения, связанные с повторным или необоснованным включением '+;
    'в реестр счетов медицинской помощи, в т.ч.:'
   .Cells(11,2) = '3.6'

   .Cells(12,01) = 'включение в счет амбулаторных посещений в период пребывания '+;
    'застрахованного лица в круглосуточном стационаре'
   .Cells(12,2) = '3.6.1'

   .Cells(13,01) = 'включение в счет пациенто-дней пребывания застрахованного лица '+;
    'в дневном стационаре в период пребывания пациента в круглосуточном стационаре'
   .Cells(13,2) = '3.6.2'

   .Cells(14,01) = 'повторное выставление счета на оплату случаев оказанной медицинской помощи, '+;
     'которые были оплачены ранее'
   .Cells(14,2) = '3.6.3'

   .Cells(15,01) = 'прочие нарушения в соответствии с Перечнем (МЭЭ)'
   .Cells(15,2) = '3.7'

   .Cells(16,01) = 'Количество принятых к оплате счетов за оказанную медицинскую помощь по '+;
    'территориальной программе ОМС'
   .Cells(16,2) = '4'

   .Cells(17,01) = 'Количество проверенных счетов'
   .Cells(17,2) = '5'

   .Cells(01,03) = 'Амбулаторно-поликлиническая помощь'
   .Cells(01,03).Orientation = 90
   .Cells(01,04) = 'Стационарная помощь'
   .Cells(01,04).Orientation = 90
   .Cells(01,05) = 'Стационар-замещающая помощь'
   .Cells(01,05).Orientation = 90
   .Cells(01,06) = 'Всего'
   .Cells(01,06).Orientation = 90
   .Cells(01,07) = 'Амбулаторно-поликлиническая помощь'
   .Cells(01,07).Orientation = 90
   .Cells(01,08) = 'Стационарная помощь'
   .Cells(01,08).Orientation = 90
   .Cells(01,09) = 'Стационар-замещающая помощь'
   .Cells(01,09).Orientation = 90
   .Cells(01,10) = 'Всего'
   .Cells(01,10).Orientation = 90
   
   .Rows("1:1").WrapText = .t.
   
   .Columns("A:A").ColumnWidth = 45
*   .Columns("A:A").Select
*   oExcel.Selection.WrapText = .t.
*   .Columns("A:A").Select
   .Columns("A:A").WrapText = .t.
  ENDWITH 

RETURN 