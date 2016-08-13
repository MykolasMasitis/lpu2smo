***********************************************************************************************
*Класс для работы с Microsoft Excel и Open Office(Libre Office), Версия 1.2 от 09.09.2013 г.  *
*																							  *
*		Предназначен для работы с файлами xls/xlt											  *
***********************************************************************************************
*
* Пример инициализации класса:
*	Ap=CREATEOBJECT('OOEX')   &&приложение по умолчанию (Microsoft Excel)
*	Ap=CREATEOBJECT('OOEX',1) &&Microsoft Excel
*	Ap=CREATEOBJECT('OOEX',2) &&Open Office (Libre Office)
*
*            *   ПРОЦЕДУРЫ И ФУНКЦИИ   *
*       =======весь документ======
* Ap.NewDoc()								Создание чистого документа
* Ap.LoadDoc(имяфайла)						Загрузить существующий файл (путь к файлу полный)
* Ap.CloseDoc()								Закрытие документа
* Ap.SaveDoc()								Сохранение открытого документа
* Ap.SaveAsDoc(имяфайла)					Сохранение открытого документа по заданному пути и имени файла
*      ======настройки листа========
* Ap.SetBasePageSetup()						Базовые настройки листа (вписывать в лист, без титулов)
*      =======листы========
* Ap.ActivateSheetByIndex(номер)			Активация листа по его номеру (индексу)
* Ap.SetActiveSheetName(новое_имя_листа)	Задает имя активного листа в книге
* Ap.SetSheetName(номер_листа,новое_имя_листа)	Задает имя конкретного листа по его номеру(индексу)
* Ap.UnProtected(пароль)					Снять защиту активного листа
* Ap.Protect(пароль)						Поставить защиту активного листа
* n=Ap.GetCountSheets()						Количество листов в книге
* str=Ap.GetSheetName(номер)				Имя листа по его номеру
* str=Ap.GetActiveSheetName()				Имя активного листа
* f=Ap.IsActiveSheetProtected() boolean		Возвращает истину, если лист защищен паролем
*     ====ячейки=========
* Ap.MergeCells(строка1,колонка1,строка2,колонка2)   Объединить ячейки (верхлев-правнижн углы прямоугольника)
* Ap.CellsWordWrap(строка,столбец)			Включить перенос текста в ячейке по словам
* Ap.SetCellText(строка,столбец,текст)		Запись текста в ячейку по координатам
*str=Ap.GetCellText(строка,столбец)=текст	Чтение текста из ячейки по координатам
* Ap.SetNameFont(строка,столбец,имя_шрифта)	Установить шрифт в ячейке
* Ap.SetFontSize(строка,столбец,размер)		Установить размер шрифта в ячейке
* Ap.SetBold(строка,столбец)				Сделать шрифт жирным в ячейке
* Ap.CellsHoriJustifyLeft(строка,стобец)	Выровнять текст в ячейке по горизонтали по левой стороне
* Ap.CellsHoriJustifyCenter(строка,стобец)	Выровнять текст в ячейке по горизонтали по центру
* Ap.CellsHoriJustifyRight(строка,стобец)	Выровнять текст в ячейке по горизонтали по правой стороне
* Ap.CellsVertJustifyCenter(строка,столбец)	Выровнять текст в ячейке по вертикали по центру
* Ap.CellsVertJustifyTop(строка,столбец)	Выровнять текст в ячейке по вертикали по верхнему краю
*     =====столбцы=======
* Ap.SetColumnWidth(колонка,ширина)			Изменить ширину столбца (100=1мм)
* Ap.InsertColByIndex(столбец,[количество]) Вставить 1(или несколько) столбцов СЛЕВА от указанного столбца
* Ap.DeleteColByIndex(столбец,[количество]) Удалить 1(или несколько) столбцов
* Ap.ColumnHoriJustifyLeft(столбец)			Выровнять ячейки в столбце по левой стороне
* Ap.ColumnHoriJustifyCenter(столбец)		Выровнять ячейки в столбце по центру
* Ap.ColumnHoriJustifyRight(столбец)		Выровнять ячейки в столбце по правой стороне
*     =====линии/рамки=========
* Ap.CellsRightBorder(строка,столбец)		Нарисовать линию справа от ячейки (пока для опена)
* Ap.CellsLeftBorder(строка,столбец)		Нарисовать линию слева от ячейки (пока для опена)
* Ap.CellsTopBorder(строка,столбец)			Нарисовать линию вверху ячейки (пока для опена)
* Ap.CellsBottomBorder(строка,столбец)		Нарисовать линию внизу ячейки (пока для опена)
* Ap.CellsBorder(строка,столбец)			Нарисовать линию вокруг всей ячейки (пока для опена)
*     ======строки=========
* Ap.InsertRowByIndex(строка,[количество])  Вставить 1(или несколько) строк над указанной строкой
* Ap.DeleteRowByIndex(строка,[количество])  Удалить 1(или несколько) строк
* Ap.SetRowHeight(строка,высота)			Установить высоту строки (100=1мм)
*             *  СВОЙСТВА  *
* Ap.BorderSize=x  (целое, от 1 и более)	Установка толщины линии рамки ячейки, по умолч 1(будет использоватся в дальнейшем рисовании рамок)
* Ap.Visible=х     (логический, .T., .F.)	Установка видимости документа (.T.-видим, .F.-невидим)(применяется сразу !)
*==================================================================
Define Class OOEX As Custom

**************************************************
* Определения констант Open Office(Libre Office) *
**************************************************
* enum CellHoriJustify specifies how cell contents are aligned horizontally
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/table/CellHoriJustify.html
	#Define ooCellHoriJustifySTANDARD   0  && default alignment is used (left for numbers, right for text)
	#Define ooCellHoriJustifyLEFT       1  && contents are aligned to the left edge of the cell
	#Define ooCellHoriJustifyCENTER     2  && contents are horizontally centered
	#Define ooCellHoriJustifyRIGHT      3  && contents are aligned to the right edge of the cell
	#Define ooCellHoriJustifyBLOCK      4  && contents are justified to the cell width
	#Define ooCellHoriJustifyREPEAT     5  && contents are repeated to fill the cell

* enum CellVertJustify specifies how cell contents are aligned vertically
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/table/CellVertJustify.html
	#Define ooCellVertJustifySTANDARD   0  && default alignment is used
	#Define ooCellVertJustifyTOP        1  && contents are aligned with the upper edge of the cell
	#Define ooCellVertJustifyCENTER     2  && contents are aligned to the vertical middle of the cell
	#Define ooCellVertJustifyBOTTOM     3  && contents are aligned to the lower edge of the cell

* constants group DocumentZoomType. These constants specify how the document content is zoomed into the document view.
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/view/DocumentZoomType.html
* service ViewSettings: h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/view/ViewSettings.html#ZoomValue provides access to the settings of the controller of an office document.
	#Define ooDocumentZoomTypeOPTIMAL			0  && The page content width (excluding margins) at the current selection is fit into the view.
	#Define ooDocumentZoomTypePAGE_WIDTH		1  && The page width at the current selection is fit into the view.
	#Define ooDocumentZoomTypeENTIRE_PAGE		2  && A complete page of the document is fit into the view.
	#Define ooDocumentZoomTypeBY_VALUE			3  && The zoom is relative and is to be set via the property ViewSettings::ZoomValue.
	#Define ooDocumentZoomTypePAGE_WIDTH_EXACT	4  && The page width at the current selection is fit into the view, with zhe view ends exactly at the end of the page.

* constants group FontWeight. Значения для CharWeight. These values are used to specify whether a font is thin or bold. They may be expanded in future versions.
* h_t_t_p://api.openoffice.org/docs/common/ref/com/sun/star/awt/FontWeight.html
	#Define ooFontWeightDONTKNOW    0.000000  && The font weight is not specified/known.
	#Define ooFontWeightTHIN       50.000000  && specifies a 50% font weight.
	#Define ooFontWeightULTRALIGHT 60.000000  && specifies a 60% font weight.
	#Define ooFontWeightLIGHT      75.000000  && specifies a 75% font weight.
	#Define ooFontWeightSEMILIGHT  90.000000  && specifies a 90% font weight.
	#Define ooFontWeightNORMAL    100.000000  && specifies a normal font weight.
	#Define ooFontWeightSEMIBOLD  110.000000  && specifies a 110% font weight.
	#Define ooFontWeightBOLD      150.000000  && specifies a 150% font weight.
	#Define ooFontWeightULTRABOLD 175.000000  && specifies a 175% font weight.
	#Define ooFontWeightBLACK     200.000000  && specifies a 200% font weight.

*******************************
* Внутренние константы класса *
*******************************
	#Define TypePrgNone 0  &&тип программы: неопределен
	#Define TypePrgExcel 1 &&тип программы: эксель
	#Define TypePrgOO    2 &&тип программы: ОО
	#Define TypePrgDefault TypePrgExcel &&Приложение по умолчанию (Microsoft Excel)
****************************
* Пропертисы класса        *
****************************
	TypeOfCalc=TypePrgNone  &&Содержит тип выбранного приложения
&&на этапе инициализации класса
&&по умолчанию приложение не задано
	Programa=Null 			&&Ссылка на обьект приложения
	Document=Null			&&Ссылка на обьект документ в приложении
	ActiveSheet=Null		&&Ссылка на текущий лист в документе приложения
	oDeskTop=Null			&&Ссылка на приложение калька в опене
&& параметры рамок, можно переустанавливать в процессе работы
	BorderSize=1			&&Толщина линии рамки ячейки
	Visible=.T.				&&Видимость документа на экране
******************************************
*      ВНУТРЕННИЕ МЕТОДЫ КЛАССА          *
******************************************

* Преобразуем путь к формату опена
	Function FileNameToURL(FileName As String) As String
		Return 'file:///'+Chrtran(FileName,'\','/')
	Endfunc  &&FileNameToURL


* Инициализация класса, в качестве значения передаем
* параметр, определяющий, на основе какого офиса он будет создан
* При невозможности создания объекта (например, не установлен офис)
* объект не будет создан.
	Function Init(parTypePrg As Integer)
*если не передан или передан неверный код приложения
*создаем по умолчанию
		If Pcount()=0
			This.TypeOfCalc=TypePrgDefault
		Else
			If Vartype(parTypePrg)<>"N"
				This.TypeOfCalc=TypePrgDefault
			Else
				If parTypePrg<>TypePrgExcel And parTypePrg<>TypePrgOO
					This.TypeOfCalc=TypePrgDefault
				Endif
			Endif
		Endif
		If This.TypeOfCalc=TypePrgNone
			This.TypeOfCalc=parTypePrg
		Endif
*создаем класс приложения на основе выбранного типа
		If This.TypeOfCalc=TypePrgExcel
			Try
&&создаем обьект приложения эксель
				This.Programa=Createobject('Excel.Application')
				This.Programa.DisplayAlerts=.F.
			Catch
				This.Programa=Null
			Endtry
		Endif
		If This.TypeOfCalc=TypePrgOO
			Try
&&обьект приложения ОО манагер? пусть пока так, там видно будет
				This.Programa=Createobject('com.sun.star.ServiceManager')
			Catch
				This.Programa=Null
			Endtry
		Endif
		If Isnull(This.Programa)
			Return .F.
		Endif
	Endfunc  &&Init

* Уничтожение обьекта
	Procedure Destroy
*this.CloseDoc()
*this.CloseProg()
	Endproc  &&Destroy

* Проверим загруженность приложения
	Function ProgLoaded() As Boolean
		Return Not(Isnull(This.Programa))
	Endfunc  &&ProgLoaded

* Проверим загруженность документа
	Function DocLoaded() As Boolean
		Return Not(Isnull(This.Document))
	Endfunc  &&DocLoaded

* Закрытие приложения
	Procedure CloseProg
		If This.DocLoaded()
			This.CloseDoc()
		Endif
		If This.ProgLoaded()
			Try
				If This.TypeOfCalc=TypePrgExcel
					This.Programa.Quit
					This.Programa=Null
				Endif
			Endtry
		Endif
		This.TypeOfCalc=TypePrgNone
	Endproc  &&CloseProg

* Задает+устанавливает видимость документа (код срабатывает при изменении пропертиса Visible класса)
	Procedure Visible_Assign
		Lparameters tAssign
		Local aFileProperties[2]
		If Vartype(tAssign)<>'L'		&&разрешен только логический тип
			Return
		Endif
		If This.Visible=tAssign			&&дальнейший код только если видимость меняется
			Return
		Endif
		This.Visible=tAssign
&& применяем новые настройки
		If This.TypeOfCalc=TypePrgExcel
			Try
				This.Programa.Visible=This.Visible
			Catch
			Endtry
		Endif
		If This.TypeOfCalc=TypePrgOO
			Try
				This.Document.getCurrentController().getFrame().getContainerWindow().SetVisible(This.Visible)
			Catch
			Endtry
		Endif
	Endproc  &&Visible_Assign

******************************************************************************
*  ПРОЦЕДУРЫ И МЕТОДЫ ДЛЯ ИСПОЛЬЗОВАНИЯ В ПРОГРАММЕ                          *
******************************************************************************


****************************************
*   ВЕСЬ ДОКУМЕНТ                      *
****************************************

* Создание чистого документа (существующий закрывается)
	Procedure NewDoc()
		Local aFileProperties[2]
		If !This.ProgLoaded()	&&если приложение не создано, выходим
			Return
		Endif
		This.CloseDoc()			&&Закроем открытый документ
		Desktop=Null
		If This.TypeOfCalc=TypePrgExcel
			This.Programa.WorkBooks.Add()
			This.Programa.Visible=This.Visible
			This.Document=This.Programa.ActiveWorkBook
			This.ActiveSheet=This.Document.ActiveSheet
		Endif
		If This.TypeOfCalc=TypePrgOO
			This.oDeskTop=This.Programa.CreateInstance('com.sun.star.frame.Desktop')
			aFileProperties[1] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[1].Name = "Hidden"
			aFileProperties[1].Value = Not This.Visible
			aFileProperties[2] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[2].Name = "ReadOnly"
			aFileProperties[2].Value = .F.
			This.Document=This.oDeskTop.LoadComponentFromURL('private:factory/scalc','_blank',0,@aFileProperties)
			This.ActivateSheetByIndex(1)
		Endif
	Endproc  &&NewDoc

* Закрытие документа (без всяких сохранений!)
	Procedure CloseDoc
*если документ существует, то закрываем
		If This.DocLoaded()
			Try
				If This.TypeOfCalc=TypePrgExcel
					This.Document.Close
				Endif
				If This.TypeOfCalc=TypePrgOO
					This.Document.Dispose
				Endif
			Catch
			Finally
				This.Document=Null
				This.ActiveSheet=Null
			Endtry
		Endif
	Endproc  &&CloseDoc

* Сохранение документа (имя берется по умолчанию, так же как и путь)
*	нужно в случае, если допустим открыли существующий документ,
*	внесли изменения и изменения надо сохранить)
* в случае успешного сохранения возвращается истина
	Function SaveDoc() As Boolean
		m.Result=.F.
		Try
			If This.DocLoaded()
				If This.TypeOfCalc=TypePrgExcel
					This.Document.Save()
					If This.Document.Saved
						m.Result=.T.
					Endif
				Endif
				If This.TypeOfCalc=TypePrgOO
					This.Document.Store()
					m.Result=.T.
				Endif
			Endif
		Endtry
		Return m.Result
	Endfunc  &&SaveDoc

* Сохранение документа по указанному пути и имени
	Function SaveAsDoc(NameFile As String) As Boolean
		Local aFileProperties[2]
		m.xlExclusive=3
		xlOtherSessionChanges=3
		m.Result=.F.
		Try
			If This.DocLoaded()
				If This.TypeOfCalc=TypePrgExcel
					This.Document.SaveAs(NameFile,,,,,,m.xlExclusive, m.xlOtherSessionChanges,,,,0)
					If This.Document.Saved
						m.Result=.T.
					Endif
				Endif
				If This.TypeOfCalc=TypePrgOO
					aFileProperties[1] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
					aFileProperties[1].Name = "Overwrite"
					aFileProperties[1].Value = .T.
					aFileProperties[2] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
					aFileProperties[2].Name = "FilterName"
					aFileProperties[2].Value = "MS Excel 97"
					This.Document.StoreToUrl(This.FileNameToURL(NameFile),@aFileProperties)
					m.Result=.T.
				Endif
			Endif
		Catch
			MESSAGEBOX("Не удалось сохранить файл '"+NameFile+"'"+CHR(10)+CHR(13)+;
			"Возможные причины:"+CHR(10)+CHR(13)+;
			"1. Не указан полный путь сохраняемого файла."+CHR(10)+CHR(13)+;
			"2. Указан несуществующий путь."+CHR(10)+CHR(13)+;
			"3. В имени пути или файла содержатся недопустимые символы."+CHR(10)+CHR(13)+;
			"4. Файл уже существует и занят другим процессом."+CHR(10)+CHR(13)+;
			"5. Файл существует, но у Вас нет прав на его перезапись.")
		Endtry
		Return m.Result
	Endfunc  &&SaveAsDoc

* Открыть существующий файл
	Procedure LoadDoc(FileName As String)
		Local aFileProperties[2]
		This.CloseDoc()
		If !This.ProgLoaded()
			Return
		Endif
		If Empty(FileName)
			Return
		Endif
		If !File(FileName)
			Return
		Endif
		This.oDeskTop=Null
		If This.TypeOfCalc=TypePrgExcel
			This.Programa.WorkBooks.Add(FileName)
			This.Programa.Visible=This.Visible
			This.Document=This.Programa.ActiveWorkBook
			This.ActiveSheet=This.Document.ActiveSheet
			This.ActivateSheetByIndex(1)
		Endif
		If This.TypeOfCalc=TypePrgOO
			This.oDeskTop=This.Programa.CreateInstance('com.sun.star.frame.Desktop')
			aFileProperties[1] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[1].Name = "Hidden"
			aFileProperties[1].Value = Not This.Visible
			aFileProperties[2] = This.Programa.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
			aFileProperties[2].Name = "ReadOnly"
			aFileProperties[2].Value = .F.
			This.Document=This.oDeskTop.LoadComponentFromURL(This.FileNameToURL(FileName), '_blank', 0,@aFileProperties)
			This.ActivateSheetByIndex(1)
		Endif
	Endproc  &&LoadDoc

****************************************
*   НАСТРОЙКИ ЛИСТА                    *
****************************************

* Задаем некие базовые настройки печати листа
* Без колонтитулов,вписывать в лист по горизонтали
	Procedure SetBasePageSetup()
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				With This.ActiveSheet.PageSetup
					.PrintTitleRows=""
					.PrintTitleColumns=""
				Endwith
				With This.ActiveSheet.PageSetup
					.LeftHeader=""
					.CenterHeader=""
					.RightHeader=""
					.LeftFooter=""
					.CenterFooter=""
					.RightFooter=""
					.PrintHeadings=.F.
					.PrintGridlines=.F.
					.Draft=.F.
					.BlackAndWhite=.F.
					.Zoom=.F.
					.FitToPagesWide=1
					.FitToPagesTall=.F.
				Endwith
			Endif
			If This.TypeOfCalc=TypePrgOO
				oPageStyles = This.Document.StyleFamilies.getByName("PageStyles")
				oPageSetup = oPageStyles.getByName(This.ActiveSheet.PageStyle)
				With oPageSetup
					.ScaleToPagesX=1
					.ScaleToPagesY=1000
					.HeaderOn=.F.
					.FooterOn=.F.
					.FooterShared=.F.
				Endwith
			Endif
		Endif
	Endproc  &&SetBasePageSetup


*******************************
*   ЛИСТЫ                     *
*******************************

* Возвращает количество листов в открытой книге
	Function GetCountSheets() As Integer
		m.Result=0
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.Document.Sheets.Count
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.Document.getSheets.GetCount
			Endif
		Endif
		Return m.Result
	Endfunc  &&GetCountSheets

* Возвращает имя листа по его номеру
	Function GetSheetName(nIndex As Integer) As String
		m.Result=''
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.Document.Sheets.Item(nIndex).Name
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.Document.getSheets.getByIndex(nIndex-1).GetName
			Endif
		Endif
		Return m.Result
	Endfunc  &&GetSheetName

* Возвращает имя активного листа
	Function GetActiveSheetName() As String
		m.Result=''
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.ActiveSheet.Name
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.ActiveSheet.GetName
			Endif
		Endif
		Return m.Result
	Endfunc	&&GetActiveSheetName

* Задать имя активного листа
	Procedure SetActiveSheetName(sName As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Name=sName
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.SetName(sName)
			Endif
		Endif
	Endproc	&&SetActiveSheetName

* Задать имя листа по его номеру
	Procedure SetSheetName(nIndex As Integer,sName As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Document.Sheets.Item(nIndex).Name=sName
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.Document.getSheets.getByIndex(nIndex-1).SetName(sName)
			Endif
		Endif
	Endproc  &&SetSheetName

* Снять защиту текущего листа
	Procedure UnProtected(sPassword As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.UnProtect(sPassword)
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.UnProtect(sPassword)
			Endif
		Endif
	Endproc	&&UnProtected

* Поставить защиту текущего листа
	Procedure Protect(sPassword As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Protect(sPassword,Null,Null,Null,Null,0)
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Protect(sPassword)
			Endif
		Endif
	Endproc	&&UnProtected

* Сделать активным лист по его номеру (индексу)
	Function ActivateSheetByIndex(nIndex As Integer) As Boolean
		m.Result=.F.
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Document.Sheets(nIndex).Activate		&&активируем лист
				This.ActiveSheet=This.Document.ActiveSheet	&&получаем на него ссылку
				m.Result=.T.
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet=This.Document.getSheets.getByIndex(nIndex-1)			&&получаем ссылку на лист
				This.Document.getCurrentController().SetActiveSheet(This.ActiveSheet)	&&активируем его
				m.Result=.T.
			Endif
		Endif
		Return m.Result
	Endfunc  &&ActivateSheetByIndex

* Проверить защищен ли активный(текущий) лист паролем
	Function IsActiveSheetProtected() As Boolean
		m.Result=.F.
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				m.Result=This.ActiveSheet.ProtectContents
			Endif
			If This.TypeOfCalc=TypePrgOO
				m.Result=This.ActiveSheet.IsProtected
			Endif
		Endif
		Return m.Result
	Endfunc	&&IsActiveSheetProtected

**************************************
*           ЯЧЕЙКИ                   *
**************************************

* Объединение ячеек (стр1,кол1,стр2,кол2)
	Procedure MergeCells(row1 As Integer,col1 As Integer,row2 As Integer, col2 As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellRangeByPosition(col1-1,row1-1,col2-1,row2-1).merge(.T.)
			Endif
		Endif
	Endproc  &&MergeCells

* Сделать перенос текста по словам в ячейке (строка, столбец)
	Procedure CellsWordWrap(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
*!*				This.Programa.ActiveSheet.Cells(Row,Col).Font.Bold=.T.
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).IsTextWrapped=.T.
			Endif
		Endif
	Endproc  &&CellsWordWrap

* Записать текст в ячейку (строка,столбец,текст)
	Procedure SetCellText(Row As Integer,Col As Integer,txt As String)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Cells(Row,Col).Formula=txt
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).SetString(txt)
			Endif
		Endif
	Endproc  &&SetCellText

* Прочитать текст из ячейки (строка,столбец)
	Function GetCellText(Row As Integer,Col As Integer) As String
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				Return This.ActiveSheet.Cells(Row,Col).Text  &&Formula
			Endif
			If This.TypeOfCalc=TypePrgOO
				Return This.ActiveSheet.getCellByPosition(Col-1,Row-1).GetString  &&GetFormula
			Endif
		Endif
	Endfunc  &&GetCellText

* Задать имя шрифта в ячейке (строка,столбец,имя_шрифта)
	Procedure SetNameFont(Row As Integer,Col As Integer,NameFont As String)
		If Empty(m.NameFont)
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(Row,Col).Font.Name=m.NameFont
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).getText.createTextCursor.charFontName=m.NameFont
			Endif
		Endif
	Endproc  &&SetNameFont

* Установить размер шрифта в ячейке (строка,столбец,размер)
	Procedure SetFontSize(Row As Integer,Col As Integer,oosize As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(Row,Col).Font.Size=oosize
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).getText.createTextCursor.CharHeight=oosize
			Endif
		Endif
	Endproc  &&SetFontSize

* Сделать шрифт жирным в ячейке (строка,столбец)
	Procedure SetBold(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(Row,Col).Font.Bold=.T.
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).getText.createTextCursor.CharWeight=150
			Endif
		Endif
	Endproc  &&SetFontSize

* Выровнять в ячейке по горизонтали по левой стороне (строка,столбец)
	Procedure CellsHoriJustifyLeft(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).HoriJustify=ooCellHoriJustifyLEFT
			Endif
		Endif
	Endproc	&&CellsHoriJustifyLeft

* Выровнять в ячейке по горизонтали по центру (строка,столбец)
	Procedure CellsHoriJustifyCenter(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).HoriJustify=ooCellHoriJustifyCENTER
			Endif
		Endif
	Endproc	&&CellsHoriJustifyCenter

* Выровнять в ячейке по горизонтали по правой стороне (строка,столбец)
	Procedure CellsHoriJustifyRight(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).HoriJustify=ooCellHoriJustifyRIGHT
			Endif
		Endif
	Endproc	&&CellsHoriJustifyRight

* В ячейке выровнять текст по вертикали по центру
	Procedure CellsVertJustifyCenter(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).VertJustify=ooCellVertJustifyCENTER
			Endif
		Endif
	Endproc  &&CellsVertJustifyCenter

* В ячейке выровнять текст по вертикали по верхнему краю
	Procedure CellsVertJustifyTop(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).VertJustify=ooCellVertJustifyTOP
			Endif
		Endif
	Endproc  &&CellsVertJustifyTop

*********************
*     СТОЛБЦЫ       *
*********************

* Изменить ширину столбца
	Procedure SetColumnWidth(Col As Integer,Width As Integer)    && 1/100мм
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.Programa.ActiveSheet.Cells(1,Col).ColumnWidth=Width/100/3
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).Width=Width
			Endif
		Endif
	Endproc  &&SetColumnWidth

* Вставить столбец (указывается НОМЕР столбца, СЛЕВА от которого будет вставлен, и кол-во столбцов (можно не указывать)
	Procedure InsertColByIndex(Col As Integer,Cnt As Integer)
		Local countCol As Integer
		m.countCol=Iif(Parameters()=2,Cnt,1) &&если не указано кол-во строк, то вставляем один
		If m.countCol<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Columns.InsertByIndex(Col-1,m.countCol)
			Endif
		Endif
	Endproc  &&InsertColByIndex

* Удалить столбец (указывается НОМЕР удаляемого столбца, и кол-во столбцов (можно не указывать))
	Procedure DeleteColByIndex(Col As Integer,Cnt As Integer)
		Local countCol As Integer
		m.countCol=Iif(Parameters()=2,Cnt,1) &&если не указано кол-во столбцов, то удаляем один
		If m.countCol<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Columns.RemoveByIndex(Col-1,m.countCol)
			Endif
		Endif
	Endproc  &&DeleteColByIndex

* В столбце выровнять по горизонтали по левой стороне все ячейки
	Procedure ColumnHoriJustifyLeft(Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).HoriJustify=ooCellHoriJustifyLEFT
			Endif
		Endif
	Endproc  &&ColumnHoriJustifyLeft

* В столбце выровнять по горизонтали по центру все ячейки
	Procedure ColumnHoriJustifyCenter(Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).HoriJustify=ooCellHoriJustifyCENTER
			Endif
		Endif
	Endproc  &&ColumnHoriJustifyCenter

* В столбце выровнять по горизонтали по правой стороне все ячейки
	Procedure ColumnHoriJustifyRIGHT(Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
* код для экселя
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(Col-1,0).getColumns.getByIndex(0).HoriJustify=ooCellHoriJustifyRIGHT
			Endif
		Endif
	Endproc  &&ColumnHoriJustifyRIGHT

******************************
*      ЛИНИИ/РАМКИ           *
******************************

* Нарисовать линию справа от ячейки
	Procedure CellsRightBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&обьект линия для опена
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*обьект линия
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&цвет черный
					.InnerLineWidth = 0  &&толщина внутренней линии
					.OuterLineWidth = This.BorderSize  &&толщина внешней линии
					.LineDistance = 0    &&дистанция между линиями
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).RightBorder=oBL  &&рамка справа
			Endif
		Endif
	Endproc  &&CellsRightBorder

* Нарисовать линию слева от ячейки
	Procedure CellsLeftBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&обьект линия для опена
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*обьект линия
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&цвет черный
					.InnerLineWidth = 0  &&толщина внутренней линии
					.OuterLineWidth = This.BorderSize  &&толщина внешней линии
					.LineDistance = 0    &&дистанция между линиями
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).LeftBorder=oBL  &&рамка справа
			Endif
		Endif
	Endproc  &&CellsLeftBorder

* Нарисовать линию вверху ячейки
	Procedure CellsTopBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&обьект линия для опена
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*обьект линия
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&цвет черный
					.InnerLineWidth = 0  &&толщина внутренней линии
					.OuterLineWidth = This.BorderSize  &&толщина внешней линии
					.LineDistance = 0    &&дистанция между линиями
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).TopBorder=oBL  &&рамка справа
			Endif
		Endif
	Endproc  &&CellsTopBorder

* Нарисовать линию внизу ячейки
	Procedure CellsBottomBorder(Row As Integer,Col As Integer)
		Local oBL As Object   &&обьект линия для опена
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
*обьект линия
				oBL=This.Programa.Bridge_GetStruct("com.sun.star.table.BorderLine")
				With oBL
					.Color = 0  &&цвет черный
					.InnerLineWidth = 0  &&толщина внутренней линии
					.OuterLineWidth = This.BorderSize  &&толщина внешней линии
					.LineDistance = 0    &&дистанция между линиями
				Endwith
				This.ActiveSheet.getCellByPosition(Col-1,Row-1).BottomBorder=oBL  &&рамка справа
			Endif
		Endif
	Endproc  &&CellsBottomBorder

* Нарисовать рамку вокруг ячейки
	Procedure CellsBorder(Row As Integer,Col As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.CellsLeftBorder(Row,Col)
				This.CellsTopBorder(Row,Col)
				This.CellsRightBorder(Row,Col)
				This.CellsBottomBorder(Row,Col)
			Endif
		Endif
	Endproc  &&CellsBorder

**************
*   СТРОКИ   *
**************

* Вставить строку (указывается НОМЕР строки, НАД которой будет вставлена, и кол-во строк (можно не указывать)
	Procedure InsertRowByIndex(Row As Integer,Cnt As Integer)
		Local countRow As Integer
		m.countRow=Iif(Parameters()=2,Cnt,1) &&если не указано кол-во строк, то вставляем одну
		If m.countRow<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Rows.InsertByIndex(Row-1,m.countRow)
			Endif
		Endif
	Endproc  &&InsertRowByIndex

* Удалить строку (указывается НОМЕР строки, и кол-во строк (можно не указывать))
	Procedure DeleteRowByIndex(Row As Integer,Cnt As Integer)
		Local countRow As Integer
		m.countRow=Iif(Parameters()=2,Cnt,1) &&если не указано кол-во строк, то удаляем одну
		If m.countRow<1
			Return
		Endif
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.Rows.RemoveByIndex(Row-1,m.countRow)
			Endif
		Endif
	Endproc  &&DeleteRowByIndex

* Установить высоту строки(указываем номер строки и ее высоту)
	Procedure SetRowHeight(Row As Integer,Height As Integer)
		If This.DocLoaded()
			If This.TypeOfCalc=TypePrgExcel
				This.ActiveSheet.Cells(m.Row,1).RowHeight=m.Height/33
			Endif
			If This.TypeOfCalc=TypePrgOO
				This.ActiveSheet.getCellByPosition(0,m.Row-1).getRows.getByIndex(0).Height=m.Height
			Endif
		Endif
	Endproc 	&&SetRowHeight

Enddefine
