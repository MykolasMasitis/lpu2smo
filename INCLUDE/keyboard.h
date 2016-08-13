* nKeyCode, nShiftAltCtrl - параметры, передаваемые в KeyPress

****************************************************************************************************************************
#DEFINE KB_ESC				27	&& < Esc >

#DEFINE KB_ENTER				13	&& < Enter >
#DEFINE KB_CTRL_ENTER		10	&& < Ctrl - Enter >
#DEFINE KB_SHIFT_ENTER		13	&& < Shift - Enter >

#DEFINE KB_SPACE				32	&& < Space >
#DEFINE KB_BACK_SPACE		127	&& < Backspace >

#DEFINE KB_SHIFT				1	&& < Shift > KeyPress
#DEFINE KB_CTRL				2	&& < Ctrl >  KeyPress
#DEFINE KB_ALT				4	&& < Alt >   KeyPress

#DEFINE KB_LEFT				19	&& < Left - Array >
#DEFINE KB_RIGHT				4	&& < Right - Array >
#DEFINE KB_UP					5	&& < Up - Array >
#DEFINE KB_DOWN				24	&& < Down - Array >

#DEFINE KB_CTRL_LEFT			26	&& < Ctrl - Left >
#DEFINE KB_CTRL_RIGHT		2	&& < Ctrl - Right >
#DEFINE KB_CTRL_UP			141	&& < Ctrl - Up >
#DEFINE KB_CTRL_DOWN		145	&& < Ctrl - Down >

#DEFINE KB_SHIFT_UP			56	&& < Shift - Up >
#DEFINE KB_SHIFT_DOWN		50	&& < Shift - Down >

#DEFINE KB_HOME				1	&& < Home >
#DEFINE KB_END				6	&& < End >
#DEFINE KB_PGUP				18	&& < PgUp >
#DEFINE KB_PGDOWN			3	&& < PgDown >

#DEFINE KB_CTRL_HOME		29	&& < Ctrl - Home >
#DEFINE KB_CTRL_END			23	&& < Ctrl - End >
#DEFINE KB_CTRL_PGUP			31	&& < Ctrl - PgUp >
#DEFINE KB_CTRL_PGDOWN		30	&& < Ctrl - PgDown >
#DEFINE KB_CTRL_INSERT		146 && < Ctrl - Insert >
#DEFINE KB_CTRL_DELETE		147 && < Ctrl - Delete >

#DEFINE KB_TAB				9	&& < Tab >
#DEFINE KB_SHIFT_TAB			15	&& < Shift - Tab >

#DEFINE KB_INSERT				22	&& < Insert >
#DEFINE KB_DELETE				7	&& < Delete >
#DEFINE KB_SHIFT_INSERT		22	&& < Shift - Insert >
#DEFINE KB_SHIFT_DELETE		7	&& < Shift - Delete >
#DEFINE KB_SHIFT_HOME		55	&& < Shift - Home >
#DEFINE KB_SHIFT_END			49	&& < Shift - End >

#DEFINE KB_F2					-1	&&  < F2 >
#DEFINE KB_F3					-2	&&  < F3 >
#DEFINE KB_F4					-3	&&  < F4 >
#DEFINE KB_F5					-4	&&  < F5 >
#DEFINE KB_F6					-5	&&  < F6 >
#DEFINE KB_F7					-6	&&  < F7 >
#DEFINE KB_F8					-7	&&  < F8 >
#DEFINE KB_F9					-8	&&  < F9 >
#DEFINE KB_F10				-9	&& < F10 >

#DEFINE KB_PLUS				43	&& < + > цифровая клава
#DEFINE KB_MINUS				45	&& < - > цифровая клава
#DEFINE KB_DROB				47	&& < / > цифровая клава
#DEFINE KB_STAR				42	&& < * > цифровая клава
#DEFINE KB_RAVNO				61	&& < = > цифровая клава
#DEFINE KB_CTRL_SLASH		28	&& Ctrl + \

#DEFINE KB_ALT_LEFT			155	&& < Ctrl - Left >
#DEFINE KB_ALT_RIGHT			157	&& < Ctrl - Right >
#DEFINE KB_ALT_UP				152	&& < Ctrl - Up >
#DEFINE KB_ALT_DOWN			160	&& < Ctrl - Down >

****************************************************************************************************************************
#DEFINE KB_ROLLBACK			KB_ESC					&& Откат на один уровень вверх

#DEFINE KB_OPEN				KB_ENTER				&& Открытие
#DEFINE KB_CTRL_OPEN		KB_CTRL_ENTER			&& Альтернативное открытие 1
#DEFINE KB_SHIFT_OPEN		KB_SHIFT_ENTER		&& Альтернативное открытие 2

#DEFINE KB_PREV_FPAGE		KB_CTRL_LEFT			&& Предыдущая страница фрейма
#DEFINE KB_NEXT_FPAGE		KB_CTRL_RIGHT			&& Следующая страница фрейма
#DEFINE KB_NEXT_KOD			KB_CTRL_UP			&& Следущий код
#DEFINE KB_PREV_KOD			KB_CTRL_DOWN			&& Предыдущий код
#DEFINE KB_GRID_SUM			KB_RAVNO				&& Итог по колонке

#DEFINE KB_SIGN				KB_SHIFT_INSERT		&& Поставить подпись
#DEFINE KB_SIGN_DEL			KB_SHIFT_DELETE		&& Снять подпись

#DEFINE KB_LIST_HOME			KB_CTRL_HOME			&& В начало списка
#DEFINE KB_LIST_END			KB_CTRL_END			&& В конец списка
#DEFINE KB_PREV_DOC			KB_CTRL_PGUP			&& Предыдущий документ
#DEFINE KB_NEXT_DOC			KB_CTRL_PGDOWN		&& Следующий документ

#DEFINE KB_NEXT_OBJ			KB_TAB					&& Следующий объект
#DEFINE KB_PREV_OBJ			KB_SHIFT_TAB			&& Предыдущий объект

#DEFINE KB_NEW				KB_INSERT				&& Добавление нового <Ins>
#DEFINE KB_NEW2				14						&& Добавление нового <Ctrl-N>
#DEFINE KB_EDIT				KB_F4					&& Коррекция
#DEFINE KB_DEL				KB_DELETE				&& Удаление старого

#DEFINE KB_COPY				KB_CTRL_INSERT		&& Копировать / Дублировать
#DEFINE KB_COPY1				3						&& Копировать Ctrl-С
#DEFINE KB_PASTE1				22						&& Вставить Ctrl-V
#DEFINE KB_FIND				KB_F2					&& Поиск
#DEFINE KB_FIND2				6						&& Поиск по Ctrl-F
#DEFINE KB_VIEW				KB_F3					&& Просмотр
#DEFINE KB_READ				KB_F5					&& Перечитать
#DEFINE KB_PRINT				KB_F7					&& Печать	<F7>
#DEFINE KB_PRINT2				16						&& Печать	<Ctrl-P>
#DEFINE KB_MENU				KB_F9					&& Локальное Альтернативное меню
#DEFINE KB_MENU_DEFAULT		93						&& Альтернативное меню через стандартную клавишу

#DEFINE KB_SEL				KB_SPACE				&& Пометить строку / Снять отметку со строки
#DEFINE KB_SELALL				KB_PLUS				&& Пометить все строки
#DEFINE KB_SELUN				KB_MINUS				&& Снять отметку со всех строк
#DEFINE KB_SELREV				KB_STAR				&& Инверсия меток всех строк
#DEFINE KB_SELDOWN			KB_SHIFT_DOWN		&& Пометка строки с переходом на следующую строку
#DEFINE KB_SELUP				KB_SHIFT_UP			&& Пометка строки с переходом на предыдущую строку
#DEFINE KB_FILTR				KB_CTRL_SLASH			&& Установка фильтра по значению поля

#DEFINE KB_SORTUP			KB_ALT_UP				&& Сортировка по возрастанию
#DEFINE KB_SORTDOWN			KB_ALT_DOWN			&& Сортировка по убыванию

#DEFINE KB_UNDO				26	&& < Ctrl - Z >		&& Откат назад
#DEFINE KB_REDO				18	&& < Ctrl - R >		&& Закат вперед

*************************************************************************************
#DEFINE KL_LEFT			(nKeyCode = KB_LEFT .and. nShiftAltCtrl = 0)					&& KB_LEFT
#DEFINE KL_RIGHT			(nKeyCode = KB_RIGHT .and. nShiftAltCtrl = 0)				&& KB_RIGHT
#DEFINE KL_UP				(nKeyCode = KB_UP .and. nShiftAltCtrl = 0)					&& KB_UP
#DEFINE KL_DOWN			(nKeyCode = KB_DOWN .and. nShiftAltCtrl = 0)				&& KB_DOWN

#DEFINE KL_PGUP			(nKeyCode = KB_PGUP .and. nShiftAltCtrl = 0)					&& KB_PGUP
#DEFINE KL_PGDOWN		(nKeyCode = KB_PGDOWN .and. nShiftAltCtrl = 0)				&& KB_PGDOWN
#DEFINE KL_CTRL_UP		(nKeyCode = KB_CTRL_UP .and. nShiftAltCtrl = KB_CTRL)		&& KB_CTRL_UP
#DEFINE KL_CTRL_DOWN	(nKeyCode = KB_CTRL_DOWN .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_DOWN

#DEFINE KL_F2				(nKeyCode = KB_F2 .and. nShiftAltCtrl = 0)					&& KB_F2
#DEFINE KL_F3				(nKeyCode = KB_F3 .and. nShiftAltCtrl = 0)					&& KB_F3
#DEFINE KL_F4				(nKeyCode = KB_F4 .and. nShiftAltCtrl = 0)					&& KB_F4
#DEFINE KL_F5				(nKeyCode = KB_F5 .and. nShiftAltCtrl = 0)					&& KB_F5				Пока не занято
#DEFINE KL_F6				(nKeyCode = KB_F6 .and. nShiftAltCtrl = 0)					&& KB_F6				Пока не занято
#DEFINE KL_F7				(nKeyCode = KB_F7 .and. nShiftAltCtrl = 0)					&& KB_F7
#DEFINE KL_F8				(nKeyCode = KB_F8 .and. nShiftAltCtrl = 0)					&& KB_F8				Пока не занято
#DEFINE KL_F9				(nKeyCode = KB_F9 .and. nShiftAltCtrl = 0)					&& KB_F9

#DEFINE KL_ROLLBACK		(nKeyCode = KB_ROLLBACK .and. nShiftAltCtrl = 0)			&& KB_ESC				Откат

#DEFINE KL_OPEN			(nKeyCode = KB_OPEN .and. nShiftAltCtrl = 0)					&& KB_ENTER			Открытие
#DEFINE KL_ALT_OPEN		(nKeyCode = KB_ALT_OPEN .and. nShiftAltCtrl = KB_ALT)		&& KB_ALT_ENTER		Альтернативное открытие
#DEFINE KL_CTRL_OPEN		(nKeyCode = KB_CTRL_OPEN .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_ENTER		Альтернативное открытие
#DEFINE KL_SHIFT_OPEN	(nKeyCode = KB_SHIFT_OPEN .and. nShiftAltCtrl = KB_SHIFT)	&& KB_SHIFT_ENTER		Альтернативное открытие

#DEFINE KL_PREV_FPAGE	(nKeyCode = KB_PREV_FPAGE .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_LEFT		Предыдущая страница фрейма
#DEFINE KL_NEXT_FPAGE	(nKeyCode = KB_NEXT_FPAGE .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_RIGHT		Следующая страница фрейма

#DEFINE KL_PREV_KOD		(nKeyCode = KB_PREV_KOD .and. nShiftAltCtrl = KB_CTRL)		&& KB_CTRL_UP			Следущий код
#DEFINE KL_NEXT_KOD		(nKeyCode = KB_NEXT_KOD .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_DOWN		Предыдущий код
#DEFINE KL_GRID_SUM		(nKeyCode = KB_GRID_SUM .and. nShiftAltCtrl = 0 )			&& KB_RAVNO			Сумма по колонке в GRIDе

#DEFINE KL_SIGN			(nKeyCode = KB_SIGN .and. nShiftAltCtrl = KB_SHIFT)			&& KB_SHIFT_INSERT	Поставить подпись
#DEFINE KL_SIGN_DEL		(nKeyCode = KB_SIGN_DEL .and. nShiftAltCtrl = KB_SHIFT)	&& KB_SHIFT_DELETE	Снять подпись

#DEFINE KL_LIST_HOME		(nKeyCode = KB_LIST_HOME .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_HOME		В начало списка
#DEFINE KL_LIST_END		(nKeyCode = KB_LIST_END .and. nShiftAltCtrl = KB_CTRL	)	&& KB_CTRL_END		В конец списка

#DEFINE KL_PREV_DOC		(nKeyCode = KB_PREV_DOC .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_PGUP		Предыдущий документ
#DEFINE KL_NEXT_DOC		(nKeyCode = KB_NEXT_DOC .and. nShiftAltCtrl = KB_CTRL)	&& KB_CTRL_PGDOWN	Следующий документ

#DEFINE KL_NEXT_OBJ		(nKeyCode = KB_NEXT_OBJ .and. nShiftAltCtrl = 0)			&& KB_TAB				Следующий объект
#DEFINE KL_PREV_OBJ		(nKeyCode = KB_PREV_OBJ .and. nShiftAltCtrl = KB_SHIFT)	&& KB_SHIFT_TAB		Предыдущий объект

#DEFINE KL_NEW2			(nKeyCode = KB_NEW2 .and. nShiftAltCtrl = KB_CTRL)			&& KB_CTRL_N			Добавление нового
#DEFINE KL_NEW			((nKeyCode = KB_NEW .and. nShiftAltCtrl = 0) .or. KL_NEW2)	&& KB_INSERT			Добавление нового
#DEFINE KL_EDIT			(nKeyCode = KB_EDIT .and. nShiftAltCtrl = 0)					&& KB_F4				Коррекция
#DEFINE KL_DEL				(nKeyCode = KB_DEL .and. nShiftAltCtrl = 0)					&& KB_DELETE			Удаление старого

#DEFINE KL_COPY			(nKeyCode = KB_COPY .and. nShiftAltCtrl = KB_CTRL	)		&& KB_CTRL_INSERT		Копировать / Дублировать
#DEFINE KL_COPY1			(nKeyCode = KB_COPY1 .and. nShiftAltCtrl = KB_CTRL)		&& KB_CTRL_C			Копировать Ctrl-С
#DEFINE KL_PASTE1			(nKeyCode = KB_PASTE1 .and. nShiftAltCtrl = KB_CTRL)		&& KB_CTRL_V			Вставить Ctrl-V
#DEFINE KL_FIND2			(nKeyCode = KB_FIND2 .and. nShiftAltCtrl = KB_CTRL)			&& Ctrl-F				Поиск по Ctrl-F
#DEFINE KL_FIND			((nKeyCode = KB_FIND .and. nShiftAltCtrl = 0)	.or. KL_FIND2)	&& KB_F2				Поиск по F2
#DEFINE KL_VIEW			(nKeyCode = KB_VIEW .and. nShiftAltCtrl = 0)					&& KB_F3				Просмотр
#DEFINE KL_READ			(nKeyCode = KB_READ .and. nShiftAltCtrl = 0)					&& KB_F5				Перечитать
#DEFINE KL_PRINT2			(nKeyCode = KB_PRINT2 .and. nShiftAltCtrl = KB_CTRL)		&& KB_F7				Печать
#DEFINE KL_PRINT			((nKeyCode = KB_PRINT .and. nShiftAltCtrl = 0).or. KL_PRINT2)	&& KB_F7				Печать
#DEFINE KL_MENU			((nKeyCode = KB_MENU .and. nShiftAltCtrl = 0) .or. (nKeyCode = KB_MENU_DEFAULT .and. nShiftAltCtrl = 1))
																						&& KB_F9				Локальное альтернативное меню F9 
																						&&						или штатная клавиша контекстного меню
#DEFINE KL_SEL				(nKeyCode = KB_SEL .and. nShiftAltCtrl = 0)					&& KB_SPACE			Пометить строку / Снять отметку со строки
#DEFINE KL_SELALL			(nKeyCode = KB_SELALL .and. nShiftAltCtrl = 0)				&& KB_PLUS				Пометить все строки
#DEFINE KL_SELUN			(nKeyCode = KB_SELUN .and. nShiftAltCtrl = 0)				&& KB_MINUS			Снять отметку со всех строк
#DEFINE KL_SELREV			(nKeyCode = KB_SELREV .and. nShiftAltCtrl = 0)				&& KB_STAR				Инверсия меток всех строк
#DEFINE KL_SELUP			(nKeyCode = KB_SELUP .and. nShiftAltCtrl = KB_SHIFT)		&& KB_SHIFT_UP		Пометка строки с переходом на предыдущую строку
#DEFINE KL_SELDOWN		(nKeyCode = KB_SELDOWN .and. nShiftAltCtrl = KB_SHIFT)	&& KB_SHIFT_DOWN		Пометка строки с переходом на следующую строку
#DEFINE KL_SELHOME		(nKeyCode = KB_SHIFT_HOME .and. nShiftAltCtrl = KB_SHIFT)	&& KB_SHIFT_HOME		Пометка всех строк от текущей вверх
#DEFINE KL_SELEND			(nKeyCode = KB_SHIFT_END .and. nShiftAltCtrl = KB_SHIFT)	&& KB_SHIFT_END		Пометка всех строк от текущей вниз

#DEFINE KL_CALC			(KL_F8)														&& KB_F8				Калькулятор
#DEFINE KL_FILTR			(nKeyCode = KB_FILTR .and. nShiftAltCtrl = KB_CTRL)			&& KB_CTRL_SLASH		Фильтр по значению

#DEFINE KL_SORTUP		(nKeyCode = KB_SORTUP .and. nShiftAltCtrl = KB_ALT)		&& Сортировка по возрастанию
#DEFINE KL_SORTDOWN		(nKeyCode = KB_SORTDOWN .and. nShiftAltCtrl = KB_ALT)		&& Сортировка по убыванию

#DEFINE KL_UNDO			(nKeyCode = KB_UNDO .and. nShiftAltCtrl = KB_CTRL )		&& KB_UNDO			Откат назад			< Ctrl - Z >
#DEFINE KL_REDO			(nKeyCode = KB_REDO .and. nShiftAltCtrl = KB_CTRL )			&& KB_REDO			Закат вперед		< Ctrl - R >

#DEFINE KL_INSERT_LINE	KL_F6
#DEFINE KL_DELETE_LINE 	KL_F7
		
****************************************************************************************************************************