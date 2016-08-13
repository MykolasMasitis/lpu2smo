FUNCTION OMS6CPDF(lcPath, IsVisible, IsQuit)
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\smo' ALIAS smo IN 0 SHARED ORDER code 
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx' ALIAS sprcokr IN 0 SHARED ORDER cokr
 IF !USED('sprlpu')
  m.WasUsedSprLpu = .f.
  =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "lpu_id")
 ELSE 
  m.WasUsedSprLpu = .t.
 ENDIF 

 SELECT AisOms
 
 m.paz    = paz
 m.s_pred = s_pred 
 m.sumz   = sumz
 m.sumh   = sumh

 m.s_iskl = sum_flk
 
 m.s_prin = m.s_pred - m.sumz - m.sumh - m.s_iskl

 m.mmy    = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 ArcFName = 'b'+mcod+"."+m.mmy
 m.arcfdate = ''
 IF fso.FileExists(lcPath+'\'+ArcFName)
  poi = fso.GetFile(lcPath+'\'+ArcFName)
  m.arcfdate = TTOC(poi.DateLastModified)
 ENDIF 

 UnzipOpen(lcPath+'\'+ArcFName)
 IsTop = UnzipGotoTopFile()
 FilesInZip = 0
 IF IsTop
  FilesInZip = 1
  DO WHILE UnzipGotoNextFile()=.T.
   FilesInZip = FilesInZip + 1
  ENDDO 
 ENDIF 
 UnzipClose()
 
 m.id_ip = ALLTRIM(cmessage)

 DocName = lcPath + "\Pr" + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 DocToOpen = DocName+'.pdf'
 
 IF fso.FileExists(DocToOpen)
  fso.DeleteFile(DocToOpen)
 ENDIF 

 m.lpuid     = lpuid
 m.mcod      = mcod
 m.lpuname   = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.cokr      = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
 m.cokr_name = IIF(SEEK(m.cokr, 'sprcokr'), ALLTRIM(sprcokr.name), '')

 =OpenFile(lcpath+'\talon', 'talon', 'shar')
 =OpenFile(lcpath+'\people', 'people', 'shar')
 =OpenFile(lcpath+'\e'+m.mcod, 'serror', 'shar', 'rid')
 =OpenFile(lcpath+'\e'+m.mcod, 'rerror', 'shar', 'rrid', 'again')

 SELECT people
 SET RELATION TO RecId INTO rError
 COUNT FOR EMPTY(rError.rid) TO m.PazPrin
 SET RELATION OFF INTO rError
 USE
 USE IN rError

 SELECT Talon 
 SET RELATION TO RecId INTO sError

 m.SchPrin = 0
 m.SchIskl = 0
 
 SCAN 
* IF !INLIST(d_type, 'z', 'h')
  m.SchPrin = m.SchPrin + IIF(EMPTY(sError.rid), 1, 0)
  m.SchIskl = m.SchIskl + IIF(!EMPTY(sError.rid), 1, 0)
* ENDIF
 ENDSCAN  

 m.schprd = RECCOUNT('Talon')

 SET RELATION OFF INTO sError
 USE 
 USE IN sError
 
 USE IN smo 
 USE IN sprcokr
 IF m.WasUsedSprLpu = .f.
  USE IN SprLpu
 ENDIF 
 
 SELECT AisOms

 =OMS6CPDFCREATE(DocName+'.pdf')

 IF IsVisible==.T.
  TRY 
   wshshell.Run("AcroRD32 " + DocToOpen)
  CATCH 
   MESSAGEBOX(CHR(13)+CHR(10)+;
    "Не удалось запустить Acrobat Reader!"+;
    CHR(13)+CHR(10),0+16,'')
  ENDTRY 
 ENDIF 

RETURN  


FUNCTION OMS6CPDFCREATE
 PARAMETERS m.OUTFILENAME
 PRIVATE m.OUTFILENAME, m.CRLF, m.STRPAGES, m.NOPAGES
 PRIVATE XREF_END_CHAR, PDFOBJECT_BEGIN, PDFOBJECT_END, PDFXREFMARKER

 DECLARE ARRXREF(20)

 PDFOBJECT_BEGIN = " 0 obj"
 PDFOBJECT_END = "endobj"
 PDFXREFMARKER	= "PDFXREFMARKER"
 XREF_END_CHAR = " 00000 n"
 
 m.SPW = 277 &&ширина "пробела" 
 m.OBJECTCOUNT = 9  && Начальное количество объектов

 m.CRLF = CHR(13)+CHR(10)

 ARRXREF(1)  = "xref"
 ARRXREF(2)  = "0 13"
 ARRXREF(3)  = "0000000000 65535 f"
 ARRXREF(4)  = PDFXREFMARKER
 ARRXREF(5)  = PDFXREFMARKER
 ARRXREF(6)  = PDFXREFMARKER
 ARRXREF(7)  = PDFXREFMARKER
 ARRXREF(8)  = PDFXREFMARKER
 ARRXREF(9)  = PDFXREFMARKER
 ARRXREF(10) = PDFXREFMARKER
 ARRXREF(11) = PDFXREFMARKER
 ARRXREF(12) = PDFXREFMARKER
 ARRXREF(13) = PDFXREFMARKER

 DECLARE	ARRDATA(1)
 ARRDATA(01) = "%PDF-1.2 "

 PDFCREATEDID() && Создаем первый объект документа - document information dictionary (DID)
 PDFADDFONT()   && Добавляем шрифт
 PDFINITIALISE() && Инициализируем документ

* m.NOPAGES = 2
 m.NOPAGES = 1
 m.STRPAGES = ""

 m.STREAMLENGTH = 0

 PDFBEGINPAGE()

 BT()
  PDFSETTM(1,0,0,1,50,545)
  PDFSETSIZEOFFONT(11)            && Устанавливаем размер шрифта
  PDFSETTEXTLEADING(11*1.2)       && Устанавливаем межстрочное расстояние
  PDFSETCHARSPACING(0)            && Межбуквенное расстояние
  PDFSETWORDSPACING(0)            && Межсловное расстояние
*  PDFSETINITTEXTPOSITION(50,545)  && Координаты первой строчки

  PDFTYPETEXT2('-'+ALLTRIM(STR((m.SPW*130)))+'(Дата приемки заявленного счета: )('+TTOC(Recieved)+')')
*  PDFTYPETEXT('')
  PDFTYPETEXT()
  PDFTYPETEXT(PADC('ПРОТОКОЛ СМО '+ALLTRIM(m.qname)+', '+m.qcod,150))
  PDFTYPETEXT(PADC('ПРИЁМКИ СЧЕТА '+ALLTRIM(m.lpuname)+', '+m.cokr_name+', '+m.mcod,190))
  PDFTYPETEXT('за медицинскую помощь, оказанную по территориальной программе ОМС г.Москвы застрахованным пациентам за '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' года')

  PDFSETSIZEOFFONT(10)            && Устанавливаем размер шрифта
  PDFTYPETEXT()
  PDFTYPETEXT()
  PDFTYPETEXT(SPACE(32)+'Предъявлено'+SPACE(45)+'Принято к рассмотрению'+SPACE(50)+'Исключено')
  PDFTYPETEXT()
  PDFTYPETEXT2('(Пациентов)-2000(Счетов)-2000(Страховая стоимость)-100(Пациентов)-2000(Счетов)-2000(Страховая стоимость)-2000(Счетов)-4500(Страховая стоимость)')
  PDFTYPETEXT()
  PDFTYPETEXT2('-2000(1)-6000(2)-8000(4)-7000(5)-5500(6)-8000(7)-8500(8)-10000(9)')

  PDFSETSIZEOFFONT(11)            && Устанавливаем размер шрифта
  PDFSETTEXTLEADING(11*1.2)       && Устанавливаем межстрочное расстояние
 ET()

* BT()
*  PDFSETINITTEXTPOSITION(60,363)  && Координаты первой строчки
*  PDFTYPETEXT3(PADR(m.paz,6))
  PDFTYPETEXT(PADR(m.paz,6), .F., 60, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(110,363)  && Координаты первой строчки
*  PDFTYPETEXT3(PADR(schprd,6))
  PDFTYPETEXT(PADR(schprd,6), .F., 110, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(185,363)  && Координаты первой строчки
*  PDFTYPETEXT3(TRANSFORM(m.s_pred,'99999999.99'))
  PDFTYPETEXT(TRANSFORM(m.s_pred,'99999999.99'), .F., 185, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(285,363)  && Координаты первой строчки
*  PDFTYPETEXT3(PADR(pazprin,6))
  PDFTYPETEXT(PADR(pazprin,6), .F., 285, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(335,363)  && Координаты первой строчки
*  PDFTYPETEXT3(PADR(schprin,6))
  PDFTYPETEXT(PADR(schprin,6), .F., 335, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(410,363)  && Координаты первой строчки
*  PDFTYPETEXT3(TRANSFORM(m.s_prin,'99999999.99'))
  PDFTYPETEXT(TRANSFORM(m.s_prin,'99999999.99'), .F., 410, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(510,363)  && Координаты первой строчки
*  PDFTYPETEXT3(PADR(schiskl,6))
  PDFTYPETEXT(PADR(schiskl,6), .F., 510, 363)
* ET()
* BT()
*  PDFSETINITTEXTPOSITION(585,363)  && Координаты первой строчки
*  PDFTYPETEXT3(TRANSFORM(m.s_iskl,'99999999.99'))
  PDFTYPETEXT(TRANSFORM(m.s_iskl,'99999999.99'), .F., 585, 363)
* ET()
 
 m.TLB = 50      && Left border - левый край таблицы
 m.TTB = 460     && Top border - верхний край таблицы
 m.TWidth = 725  && ширина таблицы
 m.THeight = 105 && высота таблицы
* PDFADDTABLE(m.TLB, m.TTB, m.TWidth, m.THeight)
 PDFADDTABLE(50, 355, 675, 105)
 PDFDRAWLINE(m.TLB, 430, m.TWidth, 430) && горизонтальные линии
 PDFDRAWLINE(m.TLB, 405, m.TWidth, 405) && горизонтальные линии
 PDFDRAWLINE(m.TLB, 380, m.TWidth, 380) && горизонтальные линии
 
 PDFDRAWLINE(100,430,100,m.TTB-m.THeight)
 PDFDRAWLINE(175,430,175,m.TTB-m.THeight)
 PDFDRAWLINE(275,460,275,m.TTB-m.THeight)
 PDFDRAWLINE(325,430,325,m.TTB-m.THeight)
 PDFDRAWLINE(400,430,400,m.TTB-m.THeight)
 PDFDRAWLINE(500,460,500,m.TTB-m.THeight)
 PDFDRAWLINE(575,430,575,m.TTB-m.THeight)
 
 BT()
  PDFSETINITTEXTPOSITION(50,350)  && Координаты первой строчки
  PDFSETSIZEOFFONT(11)            && Устанавливаем размер шрифта
  PDFSETTEXTLEADING(11*1.2)       && Устанавливаем межстрочное расстояние

  PDFTYPETEXT('Страховая стоимость услуг, заявленных для исполнения в другой МО:')
  PDFTYPETEXT(' по Центру здоровья:' + TRANSFORM(m.sumh, '9999999.99')+ ' руб.коп.')
  PDFTYPETEXT(' пр лечении по МС:' + TRANSFORM(m.sumz, '9999999.99')+ ' руб.коп.')
  PDFTYPETEXT('')
  PDFTYPETEXT2('-31500(СОГЛАСОВАНО:)')
  PDFTYPETEXT('Представитель СМО '+REPLICATE('_',38)+' Руководитель ЛПУ '+REPLICATE('_',38))
  PDFTYPETEXT('')
  PDFTYPETEXT('Дата '+REPLICATE('_',15)+SPACE(11)+'(подпись, фамилия и.о.)'+SPACE(23)+'Дата '+REPLICATE('_',15)+SPACE(11)+'(подпись, фамилия и.о.)')
 ET()
 

 BT()
  PDFSETINITTEXTPOSITION(50,200)  && Координаты первой строчки
  PDFSETSIZEOFFONT(9)            && Устанавливаем размер шрифта
  PDFSETTEXTLEADING(9*1.2)       && Устанавливаем межстрочное расстояние
  
  PDFTYPETEXT('Протокол электронной версии заявленного счета на пациентов:')
  PDFTYPETEXT('Имя архивного файла: '+'b'+m.mcod+m.mmy+'; Дата создания: '+m.arcfdate+'; Контрольное число: ')
  PDFTYPETEXT('Количество вложений в архивном файле, всего: '+STR(FilesInZip,1))
  PDFTYPETEXT('  в т.ч. реестров пациентов: 1'+'; реестров счетов: 2')
  PDFTYPETEXT('Версия тарифа: 055.310112')

  PDFTYPETEXT('Уникальный идентификатор ИП: '+m.id_ip)
  
*  PDFTYPETEXT4('123456789')
 ET()

 PDFENDPAGE()

* m.STREAMLENGTH = 0
* PDFBEGINPAGE()
* BT()
*  PDFSETTM(1,0,0,1,50,545)
*  PDFSETSIZEOFFONT(11)            && Устанавливаем размер шрифта
*  PDFSETTEXTLEADING(11*1.2)       && Устанавливаем межстрочное расстояние
*  PDFSETCHARSPACING(0)            && Межбуквенное расстояние
*  PDFSETWORDSPACING(0)            && Межсловное расстояние
*  PDFTYPETEXT('Ха-ха-ха!')
* ET()
* PDFENDPAGE()

 PDFADDCATALOGDETAILS(842,595) && 595,842 - A4 portrait; 842,595 - A4 ladscape

 ARRXREF(5) = ARRXREF(ALEN(ARRXREF))
 ARRXREF(ALEN(ARRXREF)) = ""

 ARRXREF(6) = PDFXREFMARKER
 m.OBJECTCOUNT = m.OBJECTCOUNT + 1
 ARRXREF(2) = "0 " + ALLTRIM(STR(m.OBJECTCOUNT))
 PDFFOOTER()
 PDFWRITE(m.OUTFILENAME)

RETURN

