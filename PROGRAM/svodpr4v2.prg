PROCEDURE SvodPr4V2
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ СВОД ПРИЛОЖЕНИЯ 4?'+CHR(13)+CHR(10)+;
 '(ВАРИАНТ 2)'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\pr4.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ PR4.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\pr4', 'pr4', 'shar')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'lpu_id') > 0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
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

 WAIT "ФОРМИРОВАНИЕ ОТЧЕТА..." WINDOW NOWAIT 

 oExcel.UseSystemSeparators = .F.
 oExcel.DecimalSeparator = '.'
 oExcel.ReferenceStyle= -4150  && xlR1C1
 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 oSheet = oBook.WorkSheets(1)
 oSheet.Select
 oSheet.Name = 'Сводная'
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,19))
 oRange.Merge
 oExcel.Cells(1,1).Value='Сводный акт об оплате расчетов по подушевому финансированию'
 oExcel.Cells(1,1).HorizontalAlignment = -4108
 oExcel.Cells(1,1).Font.Size = 12
 oExcel.Cells(1,1).Font.Bold = .F.
 oExcel.Cells(1,1).Font.Italic = .T.
 oExcel.Rows(1).RowHeight = 30
 oExcel.Rows(1).VerticalAlignment = -4108
 
 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 21
 oExcel.Columns(3).ColumnWidth  = 8
 oExcel.Columns(4).ColumnWidth  = 12
 oExcel.Columns(5).ColumnWidth  = 18
 oExcel.Columns(6).ColumnWidth  = 18
 oExcel.Columns(7).ColumnWidth  = 18
 oExcel.Columns(8).ColumnWidth  = 18
 oExcel.Columns(9).ColumnWidth  = 18
 oExcel.Columns(10).ColumnWidth = 18
 oExcel.Columns(11).ColumnWidth = 18
 oExcel.Columns(12).ColumnWidth = 18
 oExcel.Columns(13).ColumnWidth = 18
 oExcel.Columns(14).ColumnWidth = 18
 oExcel.Columns(15).ColumnWidth = 18
 oExcel.Columns(16).ColumnWidth = 18
 oExcel.Columns(17).ColumnWidth = 18
 oExcel.Columns(18).ColumnWidth = 18
 oExcel.Columns(19).ColumnWidth = 18
 oExcel.Columns(20).ColumnWidth = 18
 
 oExcel.Cells(2,1).Value  = '№ п\п'
 oExcel.Cells(2,2).Value  = 'Наименование ЛПУ (юридического лица) в разрезе по районам'
 oExcel.Cells(2,3).Value  = 'Округ'
 oExcel.Cells(2,4).Value  = 'Код МО'
 oExcel.Cells(2,5).Value  = 'Кол-во прикрепленного населения'
 oExcel.Cells(2,6).Value  = 'МЭК'
 oExcel.Cells(2,7).Value  = 'Расчетный объем подушевого финаснирования'
 oExcel.Cells(2,8).Value  = 'АВАНС'
 oExcel.Cells(2,9).Value = 'Фактический объем финансирования'
 oExcel.Cells(2,10).Value = 'Сумма средств, подлежащих исключению из финансирования по ;
  результам взаиморасчета с другими ГП за пацентов, прикрепленных к данному ГП, но;
  получивших амбулаторную медицинскую помощь в других ГП'
 oExcel.Cells(2,11).Value = 'Сумма средств, полученных на пациентов, прикрепленных ;
  к другим ГП, но в отчетном периоде пролеченных в данном ГП'
 oExcel.Cells(2,12).Value  = 'Сумма средств на пациентов, прикрепленных к данному ГП ;
  и получивших амбулаторную медицинскую помощь в данном ГП, сверх полученного авансирования '
 oExcel.Cells(2,13).Value  = 'Сумма дополнительных средств,полученных ГП за комплексные услуги ;
  профилактического направления и средств за оказанную медицинскую помощь в дневных ;
  стационарах'
 oExcel.Cells(2,14).Value  = 'Остаток средств от финансирования по подушевому нормативу'
 oExcel.Cells(2,15).Value  = 'Сумма средсвт за медицинскую помощь,оказанную в данном ГП ;
  гражданам,прикрепленным к ГП,не участвующим в подушевом финансировании'
 oExcel.Cells(2,16).Value  = 'Сумма средств за медицинскую помощь, оказанную в данном ГП ;
  гражданам, не прикрепленным к городским поликлиникам'
 oExcel.Cells(2,17).Value = 'Дефектная величина'
 oExcel.Cells(2,18).Value = 'Сумма средств по результатам проведения МЭЭ и ЭКМП'
 oExcel.Cells(2,19).Value = 'ИТОГО сумма средств к перечислению ;
  (гр.6+гр.7+гр.8+гр.9-гр.10'
 oExcel.Cells(2,20).Value = 'Итого к оплате'

 oExcel.Rows(2).RowHeight = 130
 oExcel.Rows(2).HorizontalAlignment = 1
 oExcel.Rows(2).VerticalAlignment = -4160
 oExcel.Rows(2).WrapText = .t.

 oExcel.Cells(3,1).Value  = '1'
 oExcel.Cells(3,2).Value  = '2'
 oExcel.Cells(3,3).Value  = '3'
 oExcel.Cells(3,4).Value  = '4'
 oExcel.Cells(3,5).Value  = '5'
 oExcel.Cells(3,6).Value  = '6'
 oExcel.Cells(3,7).Value  = '7'
 oExcel.Cells(3,8).Value  = '8'
 oExcel.Cells(3,9).Value  = '9'
 oExcel.Cells(3,10).Value = '10'
 oExcel.Cells(3,11).Value = '11'
 oExcel.Cells(3,12).Value = '12'
 oExcel.Cells(3,13).Value = '13'
 oExcel.Cells(3,14).Value = '14'
 oExcel.Cells(3,15).Value = '15'
 oExcel.Cells(3,16).Value = '16'
 oExcel.Cells(3,17).Value = '17'
 oExcel.Cells(3,18).Value = '18'
 oExcel.Cells(3,19).Value = '19'
 oExcel.Cells(3,20).Value = '20'

 nRow  = 4
 nnRow = 1

 m.col5  = 0
 m.col6  = 0
 m.col7  = 0
 m.col8  = 0
 m.col9  = 0
 m.col10 = 0
 m.col11 = 0
 m.col12 = 0
 m.col13 = 0
 m.col14 = 0
 m.col15 = 0
 m.col16 = 0
 m.col17 = 0
 m.col18 = 0
 m.col19 = 0
 m.col20 = 0

 SELECT pr4
 SET RELATION TO mcod INTO aisoms 
 SCAN 
  SCATTER MEMVAR 
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
  m.cokr = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.cokr, '')
  m.lpudog = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')

*  m.str01 = adnorm*adults + chnorm*childs
  m.str01 = finval
  m.str02 = aisoms.s_avans 
  m.str31 = s_others
  m.str32 = s_guests
  m.str33 = s_own - m.str02
  m.str03 = m.str02 - m.str31 + m.str32 + m.str33
*  m.str04 = s_kompl + s_dst
  m.str04 = aisoms.s_dop
  m.str05 = m.str01 - IIF(m.str03<0, -1*m.str03, m.str03)
  m.str06 = s_npilot
  m.str07 = s_empty
  m.str08  = aisoms.e_mee+aisoms.e_ekmp
  m.str09 = IIF(m.str05>=0, m.str05, 0)+m.str04+m.str06+m.str07
  m.koplate = m.str01 - m.str02 - m.str31 + m.str32 + m.str04 + m.str06 + m.str07 - m.str08

  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.lpuname
  oExcel.Cells(nRow,3).Value  = m.cokr
  oExcel.Cells(nRow,4).Value  = m.mcod
*  oExcel.Cells(nRow,5).Value  = adults + childs
  oExcel.Cells(nRow,5).Value  = pazval
  oExcel.Cells(nRow,6).Value  = aisoms.sum_flk
*  oExcel.Cells(nRow,7).Value  = adnorm*adults + chnorm*childs
  oExcel.Cells(nRow,7).Value  = finval
  oExcel.Cells(nRow,8).Value  = aisoms.s_avans
  oExcel.Cells(nRow,9).Value = IIF(m.koplate<0, -1*m.koplate, 0)
  oExcel.Cells(nRow,10).Value = m.str31
  oExcel.Cells(nRow,11).Value = m.str32
  oExcel.Cells(nRow,12).Value = m.str33
  oExcel.Cells(nRow,13).Value = m.str04
  oExcel.Cells(nRow,14).Value = m.str05
  oExcel.Cells(nRow,15).Value = m.str06
  oExcel.Cells(nRow,16).Value = m.str07 
  oExcel.Cells(nRow,17).Value = s_bad
  oExcel.Cells(nRow,18).Value = m.str08
  oExcel.Cells(nRow,19).Value = IIF(m.str05>=0, m.str05, 0) + m.str04 + m.str06 + m.str07 - m.str08
  oExcel.Cells(nRow,20).Value = IIF(m.koplate>=0, m.koplate, 0)
  
*  m.col5  = m.col5 + (adults + childs)
  m.col5  = m.col5 + pazval
  m.col6  = m.col6 + aisoms.sum_flk
*  m.col7  = m.col7 + (adnorm*adults + chnorm*childs)
  m.col7  = m.col7 + finval
  m.col8  = m.col8 + aisoms.s_avans
  m.col9  = m.col9 + IIF(m.koplate<0, -1*m.koplate, 0)
  m.col10 = m.col10 + m.str31
  m.col11 = m.col11 + m.str32
  m.col12 = m.col12 + m.str33
  m.col13 = m.col13 + m.str04
  m.col14 = m.col14 + m.str05
  m.col15 = m.col15 + m.str06
  m.col16 = m.col16 + m.str07
  m.col17 = m.col17 + s_bad
  m.col18 = m.col18 + m.str08
  m.col19 = m.col19 + (IIF(m.str05>=0, m.str05, 0) + m.str04 + m.str06 + m.str07 - m.str08)
  m.col20 = m.col20 + IIF(m.koplate>=0, m.koplate, 0)
  
  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 SET RELATION OFF INTO aisoms
 
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 
 WAIT CLEAR 
 
 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 oExcel.Cells(nRow,6).Value  = m.col6
 oExcel.Cells(nRow,7).Value  = m.col7
 oExcel.Cells(nRow,8).Value  = m.col8
 oExcel.Cells(nRow,9).Value  = m.col9
 oExcel.Cells(nRow,10).Value  = m.col10
 oExcel.Cells(nRow,11).Value  = m.col11
 oExcel.Cells(nRow,12).Value  = m.col12
 oExcel.Cells(nRow,13).Value  = m.col13
 oExcel.Cells(nRow,14).Value = m.col14
 oExcel.Cells(nRow,15).Value = m.col15
 oExcel.Cells(nRow,16).Value = m.col16
 oExcel.Cells(nRow,17).Value = m.col17
 oExcel.Cells(nRow,18).Value = m.col18
 oExcel.Cells(nRow,19).Value = m.col19
 oExcel.Cells(nRow,20).Value = m.col20

 BookName = 'svpr4'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pbase+'\'+m.gcperiod+'\'+BookName+'.xls',18)
 oExcel.Visible = .T.
 
RETURN 