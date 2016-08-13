PROCEDURE ModernStac

* IF tmonth<6
*  MESSAGEBOX('ОТЧЕТ ФОРМИРУЕТСЯ НАЧИНАЯ С ИЮНЯ 2012 ГОДА!', 0+16, '')
*  RETURN 
* ENDIF 

 IF !fso.FolderExists(pOut)
  fso.CreateFolder(pOut)
 ENDIF 

 IF !fso.FolderExists(pOut+'\'+gcperiod)
  fso.CreateFolder(pOut+'\'+gcperiod)
 ENDIF 

 IF !fso.FolderExists(pOut+'\'+gcperiod+'\Модернизация стационаров')
  fso.CreateFolder(pOut+'\'+gcperiod+'\Модернизация стационаров')
 ENDIF 

 prioddir = pout+'\'+STR(tYear,4)+PADL(tMonth,2,'0')
 
 IF OpenFile(pcommon+'\stac_mod', "stac_mod", "shar", "mcod")>0
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\tar_s', "tar_s", "shar", "cod")>0
  USE IN stac_mod  
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "lpu_id")>0
  USE IN stac_mod  
  USE IN tar_s
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar', 'lpu_id')>0
  USE IN sprlpu
  USE IN stac_mod  
  USE IN tar_s
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shar', 'cod')>0
  USE IN sprabo
  USE IN sprlpu
  USE IN stac_mod  
  USE IN tar_s
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'lpuid')>0
  USE IN tarif
  USE IN sprabo
  USE IN sprlpu
  USE IN stac_mod  
  USE IN tar_s
  RETURN 
 ENDIF 
 
 STmdr = pcommon+'\stmdr'
 IF !fso.FileExists(STmdr+'.dbf')
  CREATE TABLE &STmdr (lpu_id n(4), mcod c(7), ;
   sum01 n(11,2), sum02 n(11,2), sum03 n(11,2), sum04 n(11,2), sum05 n(11,2), sum06 n(11,2),;
   sum07 n(11,2), sum08 n(11,2), sum09 n(11,2), sum10 n(11,2), sum11 n(11,2), sum12 n(11,2))
  INDEX ON lpu_id TAG lpu_id
  INDEX ON mcod TAG mcod
  USE 
 ENDIF 

 sumname = 'sum'+PADL(tMonth,2,'0')

 =OpenFile(pcommon+'\stmdr', 'stmdr', 'shar', 'lpu_id')

 SELECT stac_mod
 SCAN 
  m.lpu_id = lpu_id
  m.mcod = IIF(SEEK(m.lpu_id, 'aisoms'), aisoms.mcod, '')
  WAIT m.mcod WINDOW NOWAIT 
  m.sumstmod = 0
  IF !EMPTY(m.mcod)
   IF aisoms.s_pred-aisoms.sum_flk > 0 && Если счета принятые счета ненулевые!
    IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
     =OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')
     =OpenFile(pbase+'\'+gcperiod+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')
     SELECT talon
     SET RELATION TO recid INTO error 
     SCAN 
      m.tip = tip 
      IF !EMPTY(m.tip)
       m.cod = cod 
       m.k_u = k_u
       m.n_kd = IIF(SEEK(m.cod, 'tarif'), tarif.n_kd, 0)
       DO CASE 
       CASE !INLIST(FLOOR(m.cod/1000),84,184,83,183)
        IF INLIST(m.tip,'0','Д','П') && Законченный МЭС
         m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), tar_s.del_1, 0)
        ELSE 
         m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), IIF(m.k_u<m.n_kd, ROUND(tar_s.del_2*m.k_u,2), tar_s.del_1), 0)
        ENDIF 
       CASE INLIST(FLOOR(m.cod/1000),84,184)
        m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), tar_s.del_1, 0)
       CASE FLOOR(m.cod/1000) = 083
        DO CASE  
         CASE m.k_u < m.n_kd
          m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), ROUND(m.k_u*tar_s.del_2,2), 0)
         CASE m.k_u = m.n_kd
          m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), tar_s.del_1, 0)
         CASE m.k_u > m.n_kd AND m.k_u <= 30
          m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), ROUND(m.k_u*tar_s.del_2,2), 0)
         CASE m.k_u > m.n_kd AND m.k_u > 30
          m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), ROUND(30*tar_s.del_2,2), 0)
        ENDCASE  
       CASE FLOOR(m.cod/1000) = 183
        m.sumstmod = m.sumstmod + IIF(SEEK(m.cod, 'tar_s'), IIF(m.k_u<=30, ROUND(m.k_u*tar_s.del_2,2), ROUND(30*tar_s.del_2,2)), 0)
       OTHERWISE 
      ENDCASE 
      ENDIF 
     ENDSCAN 
     SET RELATION OFF INTO error
     USE 
     USE IN error
     IF !SEEK(m.lpu_id, 'stmdr')
      INSERT INTO stmdr (lpu_id, mcod, &sumname) VALUES (m.lpu_id, m.mcod, m.sumstmod)
     ELSE 
      UPDATE stmdr SET &sumname=m.sumstmod WHERE lpu_id=m.lpu_id
     ENDIF 
    ELSE 
     IF !SEEK(m.lpu_id, 'stmdr')
      m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
      INSERT INTO stmdr (lpu_id, mcod, &sumname) VALUES (m.lpu_id, m.mcod, 0)
     ELSE 
      UPDATE stmdr SET &sumname=m.sumstmod WHERE lpu_id=m.lpu_id
     ENDIF 
    ENDIF 
   ELSE 
    IF !SEEK(m.lpu_id, 'stmdr')
     m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
     INSERT INTO stmdr (lpu_id, mcod, &sumname) VALUES (m.lpu_id, m.mcod, 0)
    ELSE 
     UPDATE stmdr SET &sumname=m.sumstmod WHERE lpu_id=m.lpu_id
    ENDIF 
   ENDIF 
  ELSE  && Если не найдено в aisoms!
   IF !SEEK(m.lpu_id, 'stmdr')
    m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
    INSERT INTO stmdr (lpu_id, mcod, &sumname) VALUES (m.lpu_id, m.mcod, 0)
   ELSE 
    UPDATE stmdr SET &sumname=m.sumstmod WHERE lpu_id=m.lpu_id
   ENDIF 
  ENDIF 
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 

 USE IN tarif
 USE IN sprabo
* USE IN sprlpu
 USE IN stac_mod  
 USE IN tar_s
 USE IN aisoms
 
 WAIT "ЗАПУСК MS WORD..." WINDOW NOWAIT 
 PUBLIC oWord as Word.Application
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 WAIT CLEAR 

 DotName = pTempl + "\prot_moder.dot"
 DotNameSv = pTempl + "\sv_modern.dot"

 DocNameSV = prioddir+'\Модернизация стационаров\Сводная ведомость'
 oDocSV = oWord.Documents.Add(dotnamesv)
 oTable = oDocSV.Tables(1)

 m.ppriod = 'на 01 ' + ;
  NameOfMonth2(IIF(tMonth<11, tMonth+2, tMonth+2-12))+' '+;
  STR(IIF(tMonth>=11, tYear+1, tYear),4)+ ' года'
 nCell = 0

 SELECT stmdr 
 SET ORDER TO mcod
 m.totst_mon = 0
 m.totst_itog = 0
 SCAN
  WAIT m.mcod WINDOW NOWAIT 
  m.lpu_id = lpu_id 
  m.mcod = mcod
  m.lpuname = IIF(SEEK(m.lpu_id, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')+', '+sprlpu.cokr+', '+sprlpu.mcod
  m.sumst_mon = &sumname
  
  m.sumst_itog = 0
  FOR ikl=1 TO tmonth
   sumtemp = 'sum'+PADL(ikl,2,'0')
   m.sumst_itog = m.sumst_itog + &sumtemp
  NEXT 

  m.totst_mon  = m.totst_mon  + m.sumst_mon
  m.totst_itog = m.totst_itog + m.sumst_itog

  DocName   = prioddir+'\Модернизация стационаров\Pm' + m.mcod
  DocNameSh = 'Pm' + m.mcod
  oDoc = oWord.Documents.Add(dotname)

  oDoc.Bookmarks('lpuname').Select  
  oWord.Selection.TypeText(m.lpuname)
  oDoc.Bookmarks('Period').Select  
  oWord.Selection.TypeText(m.ppriod)

  oDoc.Bookmarks('sumst_mon').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sumst_mon, '99 999 999.99'))
  oDoc.Bookmarks('sumst_itog').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sumst_itog, '99 999 999.99'))

  oDoc.SaveAs(DocName, 0)
  oDoc.Close

  oTable.Cell(3+nCell,1).Select
  oWord.Selection.TypeText(m.lpuname)
  oTable.Cell(3+nCell,2).Select
  oWord.Selection.TypeText(TRANSFORM(m.sumst_mon, '99 999 999.99'))
  oTable.Cell(3+nCell,3).Select
  oWord.Selection.TypeText(TRANSFORM(m.sumst_itog, '99 999 999.99'))

  oTable.Cell(3+nCell,1).Select
  oWord.Selection.InsertRowsBelow
  
  nCell = nCell + 1

  WAIT CLEAR 

 ENDSCAN 
 WAIT CLEAR 
 USE 
 USE IN sprlpu

 oTable.Cell(4+nCell,1).Select
 oWord.Selection.TypeText('Итого')
 oTable.Cell(4+nCell,2).Select
 oWord.Selection.TypeText(TRANSFORM(m.totst_mon, '99 999 999.99'))
 oTable.Cell(4+nCell,3).Select
 oWord.Selection.TypeText(TRANSFORM(m.totst_itog, '99 999 999.99'))

 oDocSV.Bookmarks('Period').Select  
 oWord.Selection.TypeText(m.ppriod)
 oDocSV.SaveAs(DocNameSV, 0)
 oDocSV.Close

 oWord.Quit
 
 MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64, '')

RETURN 