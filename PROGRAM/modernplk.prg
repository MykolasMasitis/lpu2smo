PROCEDURE ModernPlk
 IF !fso.FolderExists(pOut)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pOut, 0+16, '')
  RETURN 
 ENDIF 

 prioddir = pout+'\'+STR(tYear,4)+PADL(tMonth,2,'0')

 IF !fso.FolderExists(prioddir)
  fso.CreateFolder(prioddir)
 ENDIF 

 IF !fso.FolderExists(prioddir+'\Модернизация поликлиник')
  fso.CreateFolder(prioddir+'\Модернизация поликлиник')
 ENDIF 

 IF OpenFile(pcommon+'\lpu_m', "lpu_m", "shar", "mcod")>0
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\usl_m', "usl_m", "shar", "cod")>0
  USE IN lpu_m  
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod")>0
  USE IN lpu_m  
  USE IN usl_m
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\tarimu48', "tarimu", "shar", "cod")>0
  USE IN lpu_m  
  USE IN usl_m
  USE IN tarimu
  RETURN 
 ENDIF 

 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar')
 
 MDRFile = 'mdr' + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 MDRFdir = prioddir + '\' + MDRFile
 
 IF fso.FileExists(MDRFdir+'.dbf')
  fso.DeleteFile(MDRFdir+'.dbf')
 ENDIF 
 
 IF !fso.FileExists(MDRFdir+'.dbf')
  IF tMonth==1
   CREATE TABLE &prioddir\&MDRFile (mcod c(7), cokr c(2), ;
    sumall n(11,2), ddall n(11), sumddall n(11,2), sum01 n(11,2), dd01 n(11), sumdd01 n(11,2)) 
   USE 
   =OpenFile("&prioddir\&MDRFile", "mdrfile", "shar")
   SELECT lpu_m
   SCAN 
    SCATTER MEMVAR 
    INSERT INTO mdrfile FROM MEMVAR 
   ENDSCAN 
  ELSE 
   prioddir_prv = pout+'\'+STR(tYear,4)+PADL(tMonth-1,2,'0')
   MDRFile_prv = 'mdr' + m.qcod + PADL(tMonth-1,2,'0') + RIGHT(STR(tYear,4),2)
   MDRFdir_prv = prioddir_prv + '\' + MDRFile_prv
   IF fso.FileExists(MDRFdir_prv+'.dbf')
    fso.CopyFile(prioddir_prv + '\' + MDRFile_prv+'.dbf', prioddir + '\' + MDRFile+'.dbf')
    =OpenFile("&prioddir\&MDRFile", "mdrfile", "excl")
    SELECT mdrfile
    columnname = 'sum'+PADL(tMonth,2,'0')
    ALTER TABLE mdrfile ADD COLUMN (columnname) n(11,2)
    columnname = 'dd'+PADL(tMonth,2,'0')
    ALTER TABLE mdrfile ADD COLUMN (columnname) n(11)
    columnname = 'sumdd'+PADL(tMonth,2,'0')
    ALTER TABLE mdrfile ADD COLUMN (columnname) n(11,2)
    IF VARTYPE(lpu_id) != 'N'
     ALTER TABLE mdrfile ADD COLUMN lpu_id n(4)
    ENDIF 
   ELSE 
    MESSAGEBOX('ОТСУТСТВУЕТ СВОДНЫЙ ФАЙЛ ПО МОДЕРНИЗАЦИИ'+CHR(13)+CHR(10)+;
     MDRFile_prv + CHR(13)+CHR(10)+;
     'ЗА ПРЕДЫДУЩИЙ ПЕРИОД!', 0+16, '')
     USE IN lpu_m
     USE IN usl_m
     USE IN sprlpu
     USE IN tarimu
     USE IN sprabo
    RETURN 
   ENDIF 
  ENDIF 
 ENDIF 

 PUBLIC oWord as Word.Application

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 DotName = pTempl + "\prot_moder.dot"
 DotNameSv = pTempl + "\sv_modern.dot"

 m.ppriod = 'на 01 ' + ;
  NameOfMonth2(IIF(tMonth<11, tMonth+2, tMonth+2-12))+' '+;
  STR(IIF(tMonth>=11, tYear+1, tYear),4)+ ' года'

* m.ppriod = 'на 01 '+NameOfMonth2(tMonth+2)+' '+STR(tYear,4)+ ' года'

 DocNameSV = prioddir+'\Модернизация поликлиник\Sv_modern'
 oDocSV = oWord.Documents.Add(dotnamesv)
 oTable = oDocSV.Tables(1)

 SELECT mdrfile
 nCell = 0

 SCAN 
  m.mcod = mcod
  WAIT m.mcod WINDOW NOWAIT 
  IF !SEEK(m.mcod, 'sprlpu')
   m.ismcod = IIF(SUBSTR(m.mcod,3,2)='01', SUBSTR(m.mcod,1,2)+'05'+SUBSTR(m.mcod,5,3), m.mcod)
   IF !SEEK(m.ismcod, 'sprlpu') 
    MESSAGEBOX('MCOD '+m.mcod+' не найден в актуальном справочнике ЛПУ!'+CHR(13)+CHR(10)+;
     'Попытка заменить его на '+m.ismcod+' не дала результатов!'+CHR(13)+CHR(10)+;
     'Исправьте mcod в справочнике '+MDRFile_prv+CHR(13)+CHR(10)+;
     'удалите только что созданный справочник '+MDRFile+CHR(13)+CHR(10)+;
     'и запустите расчет модернизации еще раз!'+CHR(13)+CHR(10), 0+48, m.mcod)
    LOOP 
   ELSE
    REPLACE mcod WITH m.ismcod
    m.lpu_id = sprlpu.lpu_id
    IF EMPTY(lpu_id)
     REPLACE lpu_id WITH m.lpu_id
    ENDIF 
   ENDIF
  ELSE  
   m.lpu_id = sprlpu.lpu_id
   IF EMPTY(lpu_id)
    REPLACE lpu_id WITH m.lpu_id
   ENDIF 
  ENDIF 

  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')+', '+sprlpu.cokr+', '+sprlpu.mcod
  m.lpuid = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
  nCell = nCell + 1

  m.sum1_mon   = 0
  m.kol_dd     = 0
  m.sum_dd_mon = 0

  IF fso.FileExists(prioddir+'\l'+m.mcod+'.dbf') && Если есть принятая инфа

   =OpenFile(prioddir+'\l'+m.mcod+'.dbf', "lfile", "shar")

   SELECT lfile
   SCAN 
    m.cod = cod 
    m.k_u = k_u
    m.old_price = IIF(SEEK(m.cod, 'tarimu'), tarimu.tarif, 0)
    m.old_sum = m.old_price*m.k_u
    IF SEEK(m.cod, 'usl_m')
     m.sum1_mon  = m.sum1_mon  + (pr_all - m.old_sum)
    ENDIF 
    IF INLIST(m.cod, 101927, 101928) && Диспансеризация!
     m.kol_dd = m.kol_dd + m.k_u
     m.sum_dd_mon  = m.sum_dd_mon  + pr_all
    ENDIF 
   ENDSCAN 
   USE 
   SELECT mdrfile

  ENDIF 

  columnname1 = 'sum'+PADL(tMonth,2,'0')
*  REPLACE &columnname1 WITH m.sum1_mon && Плановые показатели достигнуты!
  REPLACE &columnname1 WITH 0           && Плановые показатели достигнуты!
  columnname2 = 'dd'+PADL(tMonth,2,'0')
  REPLACE &columnname2 WITH m.kol_dd
  columnname3 = 'sumdd'+PADL(tMonth,2,'0')
  REPLACE &columnname3 WITH m.sum_dd_mon
  
*  REPLACE sumall   WITH sumall+m.sum1_mon && Плановые показатели достигнуты!
  REPLACE ddall    WITH ddall+m.kol_dd
  REPLACE sumddall WITH sumddall+m.sum_dd_mon
  
  DocName   = prioddir+'\Модернизация поликлиник\Pm' + m.mcod
  DocNameSh = 'Pm' + m.mcod
  oDoc = oWord.Documents.Add(dotname)

  oDoc.Bookmarks('lpuname').Select  
  oWord.Selection.TypeText(m.lpuname)
  oDoc.Bookmarks('Period').Select  
  oWord.Selection.TypeText(m.ppriod)

  oDoc.Bookmarks('sum1_mon').Select  
*  oWord.Selection.TypeText(TRANSFORM(m.sum1_mon, '99 999 999.99')) && Плановые показатели достигнуты!
  oWord.Selection.TypeText(TRANSFORM(0, '99 999 999.99'))
  oDoc.Bookmarks('sum1_itog').Select  
  oWord.Selection.TypeText(TRANSFORM(sumall, '99 999 999.99'))
  oDoc.Bookmarks('kol_dd').Select  
  oWord.Selection.TypeText(TRANSFORM(m.kol_dd, '9999999'))
  oDoc.Bookmarks('sum_dd_mon').Select  
  oWord.Selection.TypeText(TRANSFORM(m.sum_dd_mon, '99 999 999.99'))
  oDoc.Bookmarks('sum_dd_itog').Select  
  oWord.Selection.TypeText(TRANSFORM(sumddall, '99 999 999.99'))

  oDoc.SaveAs(DocName, 0)
  oDoc.Close

*  m.cTO  = IIF(SEEK(m.lpuid, 'sprabo', 'lpu_id'), 'usr010@'+ALLTRIM(sprabo.abn_name), '')
  
*  m.un_id    = SYS(3)
*  m.bansfile = 'b_mdr_' + m.mcod
*  m.tansfile = 't_mdr_' + m.mcod
*  m.dfile    = 'd_mdr_' + m.mcod
*  m.mmid     = m.un_id+'.USR010'+'@'+m.qmail
*  m.csubj    = 'Otchet po modernizacii'

*  poi = fso.CreateTextFile(prioddir + '\' + m.tansfile)

*  poi.WriteLine('To: '+m.cTO)
*  poi.WriteLine('Message-Id: ' + m.mmid)
*  poi.WriteLine('Subject: ' + m.csubj)
*  poi.WriteLine('Content-Type: multipart/mixed')
*  poi.WriteLine('Attachment: '+m.dfile+' '+DocNameSh+'.doc')
 
*  poi.Close
 
*  fso.CopyFile(prioddir+'\Модернизация поликлиник\'+DocNameSh+'.doc', pAisOms+'\usr010\output\'+m.dfile)
*  fso.CopyFile(prioddir+'\'+m.tansfile, pAisOms+'\usr010\output\'+m.bansfile)
*  fso.DeleteFile(prioddir+'\'+m.tansfile)
  
  IF sumall>0 OR sumddall>0 
   oTable.Cell(3+nCell,1).Select
   oWord.Selection.TypeText(m.lpuname)
   oTable.Cell(3+nCell,2).Select
   oWord.Selection.TypeText('-')
   oTable.Cell(3+nCell,3).Select
   oWord.Selection.TypeText('-')
   oTable.Cell(3+nCell,4).Select
*   oWord.Selection.TypeText(TRANSFORM(m.sum1_mon, '99 999 999.99')) && Плановые показатели достигнуты!
   oWord.Selection.TypeText(TRANSFORM(0, '99 999 999.99')) && Плановые показатели достигнуты!
   oTable.Cell(3+nCell,5).Select
   oWord.Selection.TypeText(TRANSFORM(sumall, '99 999 999.99'))
   oTable.Cell(3+nCell,6).Select
   oWord.Selection.TypeText(TRANSFORM(m.kol_dd, '9999999'))
   oTable.Cell(3+nCell,7).Select
   oWord.Selection.TypeText(TRANSFORM(m.sum_dd_mon, '99 999 999.99'))
   oTable.Cell(3+nCell,8).Select
   oWord.Selection.TypeText(TRANSFORM(sumddall, '99 999 999.99'))

   oTable.Cell(3+nCell,1).Select
   oWord.Selection.InsertRowsBelow
  ENDIF 


 ENDSCAN 

 clmname1 = 'sum'+PADL(tMonth,2,'0')
 clmname2 = 'dd'+PADL(tMonth,2,'0')
 clmname3 = 'sumdd'+PADL(tMonth,2,'0')
 SUM &clmname1, sumall, &clmname2, &clmname3, sumddall TO ;
  m.sum1_mon_tot, m.sum1_itog_tot, m.kol_dd_tot, m.sum_dd_mon_tot, m.sum_dd_itog_tot
 
 oTable.Cell(4+nCell,1).Select
 oWord.Selection.TypeText('Итого')
 oTable.Cell(4+nCell,2).Select
 oWord.Selection.TypeText('-')
 oTable.Cell(4+nCell,3).Select
 oWord.Selection.TypeText('-')
 oTable.Cell(4+nCell,4).Select
* oWord.Selection.TypeText(TRANSFORM(m.sum1_mon_tot, '99 999 999.99')) && Плановые показатели достигнуты!
 oWord.Selection.TypeText(TRANSFORM(0, '99 999 999.99')) 
 oTable.Cell(4+nCell,5).Select
 oWord.Selection.TypeText(TRANSFORM(m.sum1_itog_tot, '99 999 999.99'))
 oTable.Cell(4+nCell,6).Select
 oWord.Selection.TypeText(TRANSFORM(m.kol_dd_tot, '9999999'))
 oTable.Cell(4+nCell,7).Select
 oWord.Selection.TypeText(TRANSFORM(m.sum_dd_mon_tot, '99 999 999.99'))
 oTable.Cell(4+nCell,8).Select
 oWord.Selection.TypeText(TRANSFORM(m.sum_dd_itog_tot, '99 999 999.99'))

 oDocSV.SaveAs(DocNameSV, 0)
 oDocSV.Close

 WAIT CLEAR 
 
 USE
 USE IN lpu_m
 USE IN usl_m
 USE IN sprlpu
 USE IN tarimu
 USE IN sprabo

 oWord.Quit
 
 MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64, '')

RETURN 