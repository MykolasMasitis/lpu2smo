*-- Определения классов
#INCLUDE	"include\MAIN.H"

DEFINE CLASS SHeader AS Header
	*-- Класс Header для "Разумного" Grid
	Name = "SHeader"
	Caption = "Заголовок"
	Alignment = 2
	FontName = "Arial"
	FontSize = 8
	WordWrap = .T.
	IsSort = .F.
	lnMouseX = 0
	lnMouseY = 0
	ToolTipTextOld = ""
	_memberdata = [<VFPData><memberdata name="issort" type="property" display="IsSort"/>] + ;
				[<memberdata name="lnmousex" type="property" display="lnMouseX"/>] + ;
				[<memberdata name="lnmousey" type="property" display="lnMouseY"/>] + ;
				[<memberdata name="tooltiptextold" type="property" display="ToolTipTextOld"/>] + ;
				[</VFPData>]
	PROCEDURE Init
		WITH THIS
			.ToolTipText = IIF(EMPTY(.ToolTipText), .Caption, .ToolTipText)
			.ToolTipTextOld = NVL(.ToolTipText, '')
		ENDWITH
	ENDPROC
	PROCEDURE MouseDown()
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		WITH THIS
			.lnMouseX = nXCoord
			.lnMouseY = nYCoord
			.IsSort = .T.
		ENDWITH
	ENDPROC
	PROCEDURE Click
		*-- Вызов метода SetOrder объекта Grid
		IF THIS.IsSort
			THIS.IsSort = .F.
			THIS.Parent.Parent.SetOrder(THIS)
		ENDIF
	ENDPROC
	PROCEDURE RightClick
		LOCAL lnMesto, llResult
		*-- Проверям, что щелкнули на Header
		llResult = THIS.Parent.Parent.GridHitTest(THIS.lnMouseX, THIS.lnMouseY, @lnMesto)
		IF llResult AND (lnMesto = 1)
			THIS.Parent.Parent.SetOption(THIS)
		ENDIF
	ENDPROC
ENDDEFINE

*-- Класс колонки для "Разумного" Grid
DEFINE CLASS SColumn AS Column
	FontChangeAllow = .T.
	FontName = "Arial"
	FontSize = 8
	Name = "SColumn"
	HeaderClass = "SHeader"
	HeaderClassLibrary = "class_define.prg"
	Format = 'Z'
	IsResize = .F.
	MinWidth = 5
	Default_IsResize = .F.
	Default_Font = ''
	Default_Visible = .T.
	Default_Width = 0
	Default_Order = 0
	Caption_Original = '' && Хранится в GridStyle - может редактироваться пользователем
	Default_Caption = '' && Заголовок по умолчанию - запоминает заголовок колонки при Init грида
	Grouped = 0
	ViewColumnAggregate = .F.
	AggFunc = ''
	Aggregate = ''
	Filtr_Column = ''
	Filtr_Not = .F.
	Filtr_Value = ''
	Order_Number = 0
	Order_Direct = ''
	Key_Order_Asc = ''
	Key_Order_Desc = ''
	IsNotHide = .F.		&& .T. - колонку нельзя скрыть, .F. - можно
	IsOrdered = .T.		&& .T. - колонка может участвовать в сортировке
	IsFiltred = .T.		&& .T. - колонка может участвовать в фильтрации
	IsAggregate = 1
	*-- IsAggregate = 0 - колонку нельзя агрегировать
	*-- IsAggregate = 1 - сумма
	*-- IsAggregate = 2 - среднее
	*-- IsAggregate = 3 - Кол-во
	*-- IsAggregate = 4 - минимум
	*-- IsAggregate = 5 - максимум
	IsEdit = .F.
	ColumnControl = "SText1"	&& Объект, который будет выводиться в данную колонку
	ColumnSaveVisible = .T.
	_memberdata = [<VFPData><memberdata name="fontchangeallow" type="property" display="FontChangeAllow"/>] + ;
				[<memberdata name="isresize" type="property" display="IsResize"/>] + ;
				[<memberdata name="minwidth" type="property" display="MinWidth"/>] + ;
				[<memberdata name="default_isresize" type="property" display="Default_IsResize"/>] + ;
				[<memberdata name="default_font" type="property" display="Default_Font"/>] + ;
				[<memberdata name="default_visible" type="property" display="Default_Visible"/>] + ;
				[<memberdata name="default_width" type="property" display="Default_Width"/>] + ;
				[<memberdata name="default_order" type="property" display="Default_Order"/>] + ;
				[<memberdata name="caption_original" type="property" display="Caption_Original"/>] + ;
				[<memberdata name="default_caption" type="property" display="Default_Caption"/>] + ;
				[<memberdata name="grouped" type="property" display="Grouped"/>] + ;
				[<memberdata name="viewcolumnaggregate" type="property" display="ViewColumnAggregate"/>] + ;
				[<memberdata name="aggfunc" type="property" display="AggFunc"/>] + ;
				[<memberdata name="aggregate" type="property" display="Aggregate"/>] + ;
				[<memberdata name="filtr_column" type="property" display="Filtr_Column"/>] + ;
				[<memberdata name="filtr_not" type="property" display="Filtr_Not"/>] + ;
				[<memberdata name="filtr_value" type="property" display="Filtr_Value"/>] + ;
				[<memberdata name="order_number" type="property" display="Order_Number"/>] + ;
				[<memberdata name="order_direct" type="property" display="Order_Direct"/>] + ;
				[<memberdata name="key_order_asc" type="property" display="Key_Order_Asc"/>] + ;
				[<memberdata name="key_order_desc" type="property" display="Key_Order_Desc"/>] + ;
				[<memberdata name="isnothide" type="property" display="IsNotHide"/>] + ;
				[<memberdata name="isordered" type="property" display="IsOrdered"/>] + ;
				[<memberdata name="isfiltred" type="property" display="IsFiltred"/>] + ;
				[<memberdata name="isaggregate" type="property" display="IsAggregate"/>] + ;
				[<memberdata name="isedit" type="property" display="IsEdit"/>] + ;
				[<memberdata name="columncontrol" type="property" display="ColumnControl"/>] + ;
				[<memberdata name="columnsavevisible" type="property" display="ColumnSaveVisible"/>] + ;
				[</VFPData>]
	PROCEDURE Init()
		WITH THIS
			SET TALK OFF
			.AddObject("SText1", "Text_For_Smart_Grid")
			IF TYPE(".SHeader1") <> 'O'
				.AddObject("SHeader1", "SHeader")
			ENDIF
			.SText1.Visible = .T.
			.CurrentControl = "SText1"
			.Default_Caption = .SHeader1.Caption
			.Default_Width = .Width
			.Default_Font = ALLTRIM(.FontName) + "," + ALLTRIM(STR(.FontSize)) + "," + ;
				IIF(.FontBold, "B", "") + IIF(.FontItalic, "I", "")
			.Default_Width = .Width
		ENDWITH
	ENDPROC
	HIDDEN PROCEDURE Default_Caption_Assign
		LPARAMETERS m.vNewVal
		m.vNewVal = IIF(VARTYPE(m.vNewVal) = 'C', m.vNewVal, .SHeader1.Caption)
		WITH THIS
			.Default_Caption =  m.vNewVal
			.Caption_Original = m.vNewVal
			.SHeader1.Caption = m.vNewVal
		ENDWITH
	ENDPROC
	HIDDEN PROCEDURE Filtr_Column_Assign
		LPARAMETERS m.vNewVal
		THIS.Filtr_Column = ALLTRIM(m.vNewVal)
		THIS.SHeader1.ToolTipText = ALLTRIM(THIS.SHeader1.ToolTipTextOld) + ;
			IIF(EMPTY(m.vNewVal), "", IIF(EMPTY(THIS.SHeader1.ToolTipTextOld), "", "; ") + "ФИЛЬТР: " + ;
				IIF(THIS.Filtr_Not, 'НЕ ', '') + ALLTRIM(CAST(THIS.Filtr_Column AS C(254))))
	ENDPROC

	PROCEDURE Moved()
		LOCAL i
		THIS.SHeader1.IsSort = .F.
		IF TYPE("THIS.Parent.oAgrGrid")=="O" AND THIS.Parent.oAgrGrid.Visible
			*-- Если включен режим агрегатов, то устанавливаем ColumnOrder
			FOR i = 1 TO THIS.Parent.ColumnCount
				THIS.Parent.oAgrGrid.Columns(i).ColumnOrder = THIS.Parent.Columns(i).ColumnOrder
			ENDFOR
		ENDIF
		IF THIS.Parent.IsMoveColumnEvent
			THIS.Parent.ColumnMove(THIS)
		ENDIF
	ENDPROC
	PROCEDURE RightClick
		THIS.Parent.RightClick()
	ENDPROC
	PROCEDURE UpdateData()
		LPARAMETERS lValue
		*-- Тут помещается код для обновления данных
	ENDPROC
	PROCEDURE GotFocusToTxt()
		*-- Вызывается при попадания фокуса на SText

	ENDPROC
	PROCEDURE Resize()
		LOCAL i
		IF (TYPE("THIS.Parent.oAgrGrid") == "O") AND THIS.Parent.oAgrGrid.Visible
			FOR m.i = 1 TO THIS.Parent.ColumnCount
				IF THIS.Parent.oAgrGrid.Columns[m.i].ColumnOrder = THIS.ColumnOrder
					EXIT
				ENDIF
			ENDFOR
			THIS.Parent.oAgrGrid.Columns[m.i].Width = THIS.Width
		ENDIF
		IF THIS.Parent.IsResizeColumnEvent
			THIS.Parent.ColumnResize(THIS)
		ENDIF
	ENDPROC

	HIDDEN PROCEDURE Caption_Original_Assign
		LPARAMETERS lcCaption
		THIS.Caption_Original = m.lcCaption
		IF THIS.Parent.IsRenameColumnEvent
			THIS.Parent.RenameHeaderCaption(THIS, THIS.Caption_Original)
		ENDIF
	ENDPROC
ENDDEFINE