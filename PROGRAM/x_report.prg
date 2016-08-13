**************************************************
*-- Генератор отчетов X-Report 1.0 (12/01/13)
*-- (c) Белан В.И., Харьков, 2013
*--     e-mail: bell_2002@mail.ru
**************************************************

LPARAMETERS m.tcFrmName, m.tcRepName, m.tlDesigner, m.tcUDFName

IF PCOUNT() < 1
	OutError('Не задано имя файла шаблона')
	RETURN .F.
ENDIF
IF VARTYPE(m.tcFrmName) <> 'C' OR EMPTY(m.tcFrmName)
	OutError('Неверное имя файла шаблона')
	RETURN .F.
ENDIF
IF PCOUNT() < 2
	OutError('Не задано имя файла отчета')
	RETURN .F.
ENDIF
IF VARTYPE(m.tcRepName) <> 'C' OR EMPTY(m.tcRepName)
	OutError('Неверное имя файла отчета')
	RETURN .F.
ENDIF
IF PCOUNT() > 2 AND VARTYPE(m.tlDesigner) <> 'L'
	OutError('Неверный признак вызова конструктора')
	RETURN .F.
ENDIF
IF PCOUNT() > 3 AND VARTYPE(m.tcUDFName) <> 'C'
	OutError('Неверное имя функции пользователя')
	RETURN .F.
ENDIF

LOCAL m.lcFrmName, m.lcRepName
m.lcFrmName = FULLPATH(m.tcFrmName)
m.lcRepName = FULLPATH(m.tcRepName)
IF m.lcFrmName == m.lcRepName
	OutError('Имя файла шаблона совпадает с именем файла отчета')
	RETURN .F.
ENDIF
IF NOT FILE(m.lcFrmName)
	OutError('Отсутствует файл шаблона: ' + m.lcFrmName)
	RETURN .F.
ENDIF

LOCAL m.lnFileHandle
m.lnFileHandle = FOPEN(m.lcFrmName,12)
IF m.lnFileHandle = -1
	OutError('Не доступен файл шаблона: ' + m.lcFrmName)
	RETURN .F.
ENDIF
FCLOSE(m.lnFileHandle)

IF FILE(m.lcRepName)
	m.lnFileHandle = FOPEN(m.lcRepName,12)
	IF m.lnFileHandle = -1
		OutError('Не доступен файл отчета: ' + m.lcRepName)
		RETURN .F.
	ENDIF
	FCLOSE(m.lnFileHandle)
	ERASE (m.lcRepName)
ENDIF

PRIVATE m.paRows, m.paCols
DIMENSION m.paRows[1,10], m.paCols[1,10]
* Группы строк (paRows) и колонок (paCols):
* i - номер группы
* [i,1] - имя группы (i=1 - таблицы, i>1 - поля)
* [i,2] - значение поля для текущей записи (i>1)
* [i,3] - начало заголовка группы
* [i,4] - длина заголовка группы
* [i,5] - начало итогов группы
* [i,6] - длина итогов группы
* [i,7] - номер текущей строки (колонки) блока
* [i,8] - счетчик записей в группе
* [i,9] - номер текущей записи в таблице блока
* [i,10] - условие для фильтра группы

PRIVATE m.pcTalk, m.pcSafety, m.pcDeleted, m.pcExact
STORE '' TO m.pcTalk, m.pcSafety, m.pcDeleted, m.pcExact

DO SetStatus IN x_report.prg

LOCAL m.lcError, m.llResult
m.lcError = ON('ERROR')
m.llResult = .F.

ON ERROR DO HandleError WITH ERROR(), MESSAGE(), PROGRAM(), LINENO()

PRIVATE m.poF1

m.poF1 = CREATEOBJECT('cntF1')

IF VARTYPE(m.poF1) = 'O'
	m.llResult = m.poF1.oleF1.GenBook( ;
		m.lcFrmName, m.lcRepName, m.tlDesigner, m.tcUDFName )
ENDIF

RELEASE m.poF1

ON ERROR &lcError

DO ResetStatus IN x_report.prg

RETURN m.llResult

*-------------------------------

PROCEDURE SetStatus

m.pcTalk = SET('TALK')
m.pcSafety = SET('SAFETY')
m.pcDeleted = SET('DELETED')
m.pcExact = SET('EXACT')

SET TALK OFF
SET SAFETY OFF
SET DELETED ON
SET EXACT ON

ENDPROC


PROCEDURE ResetStatus

IF m.pcTalk = 'ON'
	SET TALK ON
ENDIF
IF m.pcSafety = 'ON'
	SET SAFETY ON
ENDIF
IF m.pcDeleted = 'OFF'
	SET DELETED OFF
ENDIF
IF m.pcExact = 'OFF'
	SET EXACT OFF
ENDIF

ENDPROC

*-------------------------------

PROCEDURE HandleError
LPARAMETERS m.tnErr, m.tcMsg, m.tcPrg, m.tnLin
LOCAL m.lcErr, m.llKey
m.lcErr = 'Ошибка ' + ALLTRIM(STR(m.tnErr)) + ' : ' + m.tcMsg
IF VARTYPE(m.poF1) = 'O'
	WITH m.poF1.oleF1
		m.llKey = (.nRow = 0 AND .nRowGr = 0 AND .nColGr = 0)
	ENDWITH
ELSE
	m.llKey = .T.
ENDIF
IF m.llKey
	m.lcErr = m.lcErr + CHR(10) + 'Ошибка в программе ' + m.tcPrg + ;
		', строка ' + ALLTRIM(STR(m.tnLin)) + '.'
ENDIF
OutError(m.lcErr)
ENDPROC


PROCEDURE OutError
LPARAMETERS m.tcError
IF RIGHT(m.tcError,1) <> '.'
	m.tcError = m.tcError + '.'
ENDIF
IF VARTYPE(m.poF1) = 'O'
	WITH m.poF1.oleF1
	DO CASE
	CASE .nRowGr <> 0
		m.tcError = m.tcError + CHR(10) + ;
		'Ошибка в шаблоне, лист ' + LTRIM(STR(.nSheet)) + ;
		', полоса ' + IIF(.nRowGr = ALEN(m.paRows,1), '', ;
		IIF(.nRowGr > 0, 'заголовка ', 'итогов ') ) + ;
		'группы ' + m.paRows[ABS(.nRowGr),1] + '.'
	CASE .nColGr <> 0
		m.tcError = m.tcError + CHR(10) + ;
		'Ошибка в шаблоне, лист ' + LTRIM(STR(.nSheet)) + ;
		', полоса ' + IIF(.nColGr = ALEN(m.paCols,1), '', ;
		IIF(.nColGr > 0, 'заголовка ', 'итогов ') ) + ;
		'группы ' + m.paCols[ABS(.nColGr),1] + '.'
	CASE .nRow > 0
		m.tcError = m.tcError + CHR(10) + ;
		'Ошибка в шаблоне, лист ' + LTRIM(STR(.nSheet)) + ', ячейка (' + ;
		LTRIM(STR(.nRow-.nShift)) + ',' + LTRIM(STR(.nCol)) + ').'
	ENDCASE
	.lStopKey = .T.
	ENDWITH
ENDIF
= MESSAGEBOX(m.tcError, 0+48+0, 'X-Report')
ENDPROC

*-------------------------------

FUNCTION FST
LPARAMETERS tcExpr
LOCAL lnRec, luVal
lnRec = RECNO()
GO TOP
luVal = EVALUATE(tcExpr)
IF lnRec <= RECCOUNT()
	GO lnRec
ELSE
	GO BOTTOM
	SKIP
ENDIF
RETURN luVal


FUNCTION LST
LPARAMETERS tcExpr
LOCAL lnRec, luVal
lnRec = RECNO()
GO BOTTOM
luVal = EVALUATE(tcExpr)
IF lnRec <= RECCOUNT()
	GO lnRec
ELSE
	GO BOTTOM
	SKIP
ENDIF
RETURN luVal

*-------------------------------

DEFINE CLASS cntF1 AS Container
Visible = .F.
ADD OBJECT oleF1 AS oleF1
ENDDEFINE

*-------------------------------

DEFINE CLASS oleF1 AS OLEControl
OLEClass = 'TTF161.TTF1.6'
Visible = .F.

lStopKey = .F.
lBlockKey = .F.
lPivotTable = .F.
lAliasUsed = .F.
cCurrAlias = ''
nCurrRow = 0
nCurrCol = 0
nPrevCol = 0
nRowGr = 0
nColGr = 0
nRow = 0
nCol = 0
nSheet = 0
nShift = 0

nPrintRowBands = 0
nAutoFitRows = 0
nMergeRowTitles = 0
nPrintColBands = 0
nAutoFitCols = 0
nMergeColTitles = 0

FUNCTION TestParams
LOCAL m.lcText, m.lnCnt, m.lnPos, m.lnLen, m.lcName, m.lcValue, m.i

.nPrintRowBands = 2
.nAutoFitRows = 2
.nMergeRowTitles = 1
.nPrintColBands = 1
.nAutoFitCols = 1
.nMergeColTitles = 1

.nRow = m.paRows[1,3]
.nCol = 1
m.lcText = .TextRC(.nRow, .nCol)
m.lnCnt = OCCURS('{', m.lcText)

FOR m.i=1 TO m.lnCnt
	m.lnPos = AT('{', m.lcText, m.i)
	m.lnLen = AT('}', m.lcText, m.i) - m.lnPos - 1
	m.lcName = UPPER(ALLTRIM(SUBSTR(m.lcText, m.lnPos+1, m.lnLen)))
	m.lnPos = AT('=', m.lcName)
	IF LEN(m.lcName) < 2 OR AT('{', m.lcName) > 0 OR m.lnPos < 2
		OutError('Неверный параметр: '+m.lcName)
		RETURN .F.
	ENDIF
	m.lcValue = ALLTRIM(SUBSTR(m.lcName, m.lnPos+1))
	m.lcName = UPPER(ALLTRIM(SUBSTR(m.lcName, 1, m.lnPos-1)))
	IF NOT INLIST(m.lcName, 'PRINTROWBANDS', 'AUTOFITROWS', 'MERGEROWTITLES', ;
							'PRINTCOLBANDS', 'AUTOFITCOLS', 'MERGECOLTITLES')
		OutError('Неверное имя параметра: '+m.lcName)
		RETURN .F.
	ENDIF
	IF NOT INLIST(m.lcValue, '0', '1', '2') OR ;
		LEFT(m.lcName,5) = 'MERGE' AND m.lcValue > '1'
		OutError('Неверное значение параметра '+m.lcName+' = '+m.lcValue)
		RETURN .F.
	ENDIF
	m.lcName = '.n' + m.lcName + '=' + m.lcValue
	= EXECSCRIPT(m.lcName)
ENDFOR

.nRow = 0
.nCol = 0
RETURN .T.
ENDFUNC


FUNCTION TestGroups
LPARAMETERS m.tcTable, m.taGroups
LOCAL m.loFind, m.lnPrevRow, m.lnNextRow, m.llFindNext, m.lnPrevCol, m.lnNextCol
LOCAL m.lcText, m.lcName, m.lnGrCnt, m.lnCnt, m.lnPos, m.lnLen
LOCAL m.llStart, m.i, m.j
m.lnGrCnt = 0
m.lnPrevRow = 1
m.lnPrevCol = 1
m.llStart = .T.

.Sheet = .nSheet

IF m.tcTable == 'x_rows'
	m.loFind = .DefineSearch('[', .nSheet, 1, 1, MAX(.LastRow, 1), 1, 16)
ELSE
	m.i = m.paRows[1,3]
	m.loFind = .DefineSearch('[', .nSheet, m.i, 1, m.i, MAX(.LastCol, 1), 16)
ENDIF
m.llFindNext = m.loFind.FindNext()
m.lnNextRow = m.loFind.Row
m.lnNextCol = m.loFind.Col
.nRow = m.lnNextRow
.nCol = m.lnNextCol

IF NOT m.llFindNext
	.nRow = 0
	.nCol = 0
	.lBlockKey = .F.
	RETURN .F.
ENDIF

IF m.lnNextCol <> 1
	OutError('Не задана таблица блока')
	RETURN .F.
ENDIF

SELECT 0

DO WHILE m.llFindNext
	.nRow = m.lnNextRow
	.nCol = m.lnNextCol
	m.lcText = .TextRC(m.lnNextRow, m.lnNextCol)
	m.lnCnt = OCCURS('[', m.lcText)

	FOR m.i=1 TO m.lnCnt
		m.lnPos = AT('[', m.lcText, m.i)
		m.lnLen = AT(']', m.lcText, m.i) - m.lnPos - 1
		m.lcName = UPPER(ALLTRIM(SUBSTR(m.lcText, m.lnPos+1, m.lnLen)))
		IF LEN(m.lcName) < 2 OR AT('[', m.lcName) > 0
			OutError('Неверный формат поля')
			RETURN .F.
		ENDIF

		DO CASE
		CASE LEFT(m.lcName,1) == '*' OR LEFT(m.lcName,1) == '#'
			IF LEFT(m.lcName,1) == '#'
				.lPivotTable = .T.
			ENDIF
			IF NOT EMPTY(ALIAS())
				OutError('Нарушен баланс групп')
				RETURN .F.
			ENDIF
			m.lcName = ALLTRIM(SUBSTR(m.lcName,2))
			IF m.tcTable == 'x_rows'
				.cCurrAlias = m.lcName
				.lAliasUsed = USED(m.lcName)
				IF NOT .lAliasUsed
					USE (m.lcName) IN 0
					IF NOT USED(m.lcName) OR .lStopKey
						RETURN .F.
					ENDIF
				ENDIF
			ENDIF
			SELECT (m.lcName)
			GOTO TOP
		CASE RIGHT(m.lcName,1) == '*' OR RIGHT(m.lcName,1) == '#'
			IF RIGHT(m.lcName,1) == '*' AND .lPivotTable OR ;
				RIGHT(m.lcName,1) == '#' AND NOT .lPivotTable
				OutError('Нарушен баланс групп')
				RETURN .F.
			ENDIF
			m.lcName = ALLTRIM(LEFT(m.lcName, LEN(m.lcName)-1))
			IF ALIAS() <> m.lcName
				OutError('Нарушен баланс групп')
				RETURN .F.
			ENDIF
			m.llStart = .F.
			SELECT 0
		OTHERWISE
			IF EMPTY(ALIAS())
				OutError('Не задана таблица блока')
				RETURN .F.
			ENDIF
			DO CASE
			CASE LEFT(m.lcName,1) == '<'
				m.lcName = ALLTRIM(SUBSTR(m.lcName, 2))
			CASE RIGHT(m.lcName,1) == '>'
				m.lcName = ALLTRIM(LEFT(m.lcName, LEN(m.lcName)-1))
				m.llStart = .F.
			OTHERWISE
				OutError('Неверный формат поля')
				RETURN .F.
			ENDCASE
			IF LEFT(m.lcName,1) == '&'
				m.lcName = EVALUATE(SUBSTR(m.lcName,2))
			ENDIF
			IF TYPE(ALIAS()+'.'+m.lcName) == 'U'
				OutError('Неверное имя группы: ' + m.lcName)
				RETURN .F.
			ENDIF
		ENDCASE

		IF m.llStart
			IF m.lnGrCnt > 0
				IF m.lnGrCnt < ALEN(m.taGroups,1)
					OutError('Нарушен баланс групп')
					RETURN .F.
				ENDIF
				FOR m.j=1 TO m.lnGrCnt
					IF m.lcName == m.taGroups[m.j,1]
						OutError('Группа с именем '+m.lcName+' уже есть')
						RETURN .F.
					ENDIF
				ENDFOR
				IF m.tcTable == 'x_rows'
					m.taGroups[m.lnGrCnt,4] = m.lnNextRow - m.lnPrevRow
				ELSE
					m.taGroups[m.lnGrCnt,4] = m.lnNextCol - m.lnPrevCol
				ENDIF
			ENDIF
			m.lnGrCnt = m.lnGrCnt + 1
			DIMENSION m.taGroups[m.lnGrCnt,10]
			m.taGroups[m.lnGrCnt,1] = m.lcName
			IF m.lnGrCnt > 1
				m.taGroups[m.lnGrCnt,2] = EVALUATE(ALIAS()+'.'+m.lcName)
			ENDIF
			IF m.tcTable == 'x_rows'
				m.taGroups[m.lnGrCnt,3] = m.lnNextRow
			ELSE
				m.taGroups[m.lnGrCnt,3] = m.lnNextCol
			ENDIF
		ELSE
			IF m.lnGrCnt < 1 OR m.lcName <> m.taGroups[m.lnGrCnt,1]
				OutError('Нарушен баланс групп')
				RETURN .F.
			ENDIF
			IF m.lnGrCnt = ALEN(m.taGroups,1)
				IF m.tcTable == 'x_rows'
					m.taGroups[m.lnGrCnt,4] = m.lnNextRow - m.lnPrevRow
				ELSE
					m.taGroups[m.lnGrCnt,4] = m.lnNextCol - m.lnPrevCol
				ENDIF
			ENDIF
			IF m.tcTable == 'x_rows'
				m.taGroups[m.lnGrCnt,5] = m.lnPrevRow
				m.taGroups[m.lnGrCnt,6] = m.lnNextRow - m.lnPrevRow
			ELSE
				m.taGroups[m.lnGrCnt,5] = m.lnPrevCol
				m.taGroups[m.lnGrCnt,6] = m.lnNextCol - m.lnPrevCol
			ENDIF
			m.lnGrCnt = m.lnGrCnt - 1
			IF m.lnGrCnt = 0
				m.lnPos = AT(']', m.lcText, m.i)
				.TextRC(.nRow, .nCol) = SUBSTR(m.lcText,m.lnPos+1)
				EXIT
			ENDIF
		ENDIF

		m.lnPrevRow = m.lnNextRow
		m.lnPrevCol = m.lnNextCol
	ENDFOR
	IF m.lnGrCnt = 0
		EXIT
	ENDIF

	= m.loFind.FindNext()
	m.lnNextRow = m.loFind.Row
	m.lnNextCol = m.loFind.Col
	m.llFindNext = (m.lnNextRow > m.lnPrevRow) OR (m.lnNextCol > m.lnPrevCol)
ENDDO

LOCAL m.lcAlias, m.lcCond, m.lcList
m.lcAlias = m.taGroups[1,1]
m.lnGrCnt = ALEN(m.taGroups,1)

m.taGroups[1,10] = '.T.'

FOR m.i=2 TO m.lnGrCnt
	m.lcCond = m.taGroups[m.i,1] + '=' + m.tcTable + '.' + m.taGroups[m.i,1]
	IF m.i = 2
		m.taGroups[m.i,10] = m.lcCond
		m.lcList = m.taGroups[m.i,1]
	ELSE
		m.taGroups[m.i,10] = m.taGroups[m.i-1,10] + ' AND ' + m.lcCond
		m.lcList = m.lcList + ',' + m.taGroups[m.i,1]
	ENDIF
ENDFOR

IF .lPivotTable
	SELECT &lcList FROM (m.lcAlias) ;
	GROUP BY &lcList ORDER BY &lcList INTO CURSOR (m.tcTable) READWRITE

	IF m.tcTable == 'x_rows'
		GO TOP IN x_rows
		FOR m.i=2 TO m.lnGrCnt
			m.taGroups[m.i,2] = EVALUATE('x_rows.' + m.taGroups[m.i,1])
		ENDFOR
	ELSE
		GO TOP IN x_cols
		FOR m.i=2 TO m.lnGrCnt
			m.taGroups[m.i,2] = EVALUATE('x_cols.' + m.taGroups[m.i,1])
		ENDFOR
	ENDIF
ELSE
	GO TOP IN (m.lcAlias)
	m.i = RECNO(m.lcAlias)

	SELECT * FROM (m.lcAlias) ;
	WHERE RECNO() = m.i INTO CURSOR x_rows READWRITE

	IF RECCOUNT('x_rows') = 0
		SELECT x_rows
		APPEND BLANK
	ENDIF
ENDIF

RETURN .T.
ENDFUNC


FUNCTION TestFields
LOCAL m.loFind, m.lnPrevRow, m.lnNextRow, m.lnPrevCol, m.lnNextCol, m.llFindNext
LOCAL m.lcText, m.lcExpr, m.lcFunc, m.lnCnt, m.lnPos, m.lnLen, m.llAll, m.luValue
LOCAL m.lnRecNo, m.i1, m.i2
.cCurrAlias = m.paRows[1,1]
m.lnPrevRow = 1
m.lnPrevCol = 2

.Sheet = .nSheet

m.i1 = m.paRows[1,3]
IF .lPivotTable
	m.i1 = m.i1 + 1
ENDIF
m.i2 = m.paRows[1,5] + m.paRows[1,6] - 1

m.loFind = .DefineSearch('[', .nSheet, m.i1, 2, m.i2, MAX(.LastCol,2), 16)
m.llFindNext = m.loFind.FindNext()
m.lnNextRow = m.loFind.Row
m.lnNextCol = m.loFind.Col

SELECT (.cCurrAlias)
GOTO TOP

DO WHILE m.llFindNext
	m.lcText = .TextRC(m.lnNextRow, m.lnNextCol)
	m.lnCnt = 1
	m.lnPos = AT('[', m.lcText, m.lnCnt)
	DO WHILE m.lnPos > 0
		m.lnLen = AT(']', m.lcText, m.lnCnt) - m.lnPos + 1
		m.lcExpr = SUBSTR(m.lcText, m.lnPos, m.lnLen)
		m.llAll = (m.lcText == m.lcExpr)
		m.lcExpr = ALLTRIM(SUBSTR(m.lcExpr, 2, m.lnLen-2))
		m.lcFunc = UPPER(LEFT(m.lcExpr,4))

		.nRow = m.lnNextRow
		.nCol = m.lnNextCol

		IF LEN(m.lcExpr) = 0 OR AT('[', m.lcExpr) > 0
			OutError('Неверный формат поля')
			RETURN .F.
		ENDIF

		m.lnRecNo = RECNO(.cCurrAlias)
		IF INLIST(m.lcFunc, 'SUM(','CNT(','MIN(','MAX(','AVG(','STD(','NPV(','VAR(')
			CALCULATE &lcExpr TO m.luValue FOR RECNO() = m.lnRecNo
		ELSE
			m.luValue = EVALUATE(m.lcExpr)
		ENDIF
		IF m.lnRecNo <> RECNO(.cCurrAlias)
			IF m.lnRecNo <= RECCOUNT(.cCurrAlias)
				GOTO m.lnRecNo IN (.cCurrAlias)
			ELSE
				GO BOTTOM
				SKIP
			ENDIF
		ENDIF

		.nRow = 0
		.nCol = 0
		IF .lStopKey
			RETURN .F.
		ENDIF

		INSERT INTO x_fields ;
			VALUES (m.lnNextRow, m.lnNextCol, 0, 0, m.lcExpr, ;
			m.llAll, .F., .F., .F., .F.)
		m.lnCnt = m.lnCnt + 1
		m.lnPos = AT('[', m.lcText, m.lnCnt)
	ENDDO

	m.lnPrevRow = m.lnNextRow
	m.lnPrevCol = m.lnNextCol
	= m.loFind.FindNext()
	m.lnNextRow = m.loFind.Row
	m.lnNextCol = m.loFind.Col
	m.llFindNext = (m.lnNextRow > m.lnPrevRow OR m.lnNextCol > m.lnPrevCol)
ENDDO

RETURN .T.
ENDFUNC


FUNCTION GenFields
LPARAMETERS m.tnDstR1, m.tnSrcR1, m.tnLen, m.tnRowNo
LOCAL m.luValue, m.lcValue, m.lnRow, m.lnCol, m.lcText, m.lnPos, m.lnLen
LOCAL m.lcAlias, m.lnRecNo, m.lnRowCnt, m.lnColCnt, m.lcCond, m.lnSrcR2
LOCAL ARRAY m.laTemp(1)

m.lcAlias = ALIAS()
m.lnRowCnt = ALEN(m.paRows,1)
m.lnColCnt = ALEN(m.paCols,1)
m.lnSrcR2 = m.tnSrcR1 + m.tnLen - 1

SELECT x_fields
SCAN FOR nRow >= m.tnSrcR1 AND nRow <= m.lnSrcR2
	SELECT (.cCurrAlias)
	.nRow = x_fields.nRow
	.nCol = x_fields.nCol
	m.lcValue = ALLTRIM(x_fields.cExpr)
	m.luValue = .F.
	m.lcFunc = UPPER(LEFT(x_fields.cExpr,4))

	IF INLIST(m.lcFunc, 'SUM(','CNT(','MIN(','MAX(','AVG(','STD(','NPV(','VAR(')
		LOCAL m.lnR, m.lnC
		m.lnR = m.tnRowNo
		m.lnC = x_fields.nGr
		IF .lPivotTable
			IF x_fields.lPivot
				GOTO x_fields.nRec IN x_cols
			ENDIF
			m.lcCond = m.paRows[m.lnR,10] + ' AND ' + m.paCols[m.lnC,10]
			SELECT (.cCurrAlias)
		ELSE
			IF m.lnR < m.lnRowCnt
				m.lcCond = m.paRows[m.lnR,10]
				SCATTER TO m.laTemp MEMO
				SELECT x_rows
				GATHER FROM m.laTemp MEMO
				SELECT (.cCurrAlias)
			ELSE
				m.lcCond = 'RECNO() <= m.lnRecNo'
			ENDIF
		ENDIF

		m.lnRecNo = RECNO(.cCurrAlias)
		CALCULATE &lcValue TO m.luValue FOR &lcCond
		IF m.lnRecNo <= RECCOUNT(.cCurrAlias)
			GOTO m.lnRecNo IN .cCurrAlias
		ELSE
			GO BOTTOM
			SKIP
		ENDIF
	ELSE
		IF .lPivotTable
			DO CASE
			CASE x_fields.lColTitle
				GOTO x_fields.nRec IN x_cols
				m.lcCond = m.paCols[m.lnColCnt,10]
			CASE x_fields.lPivot
				GOTO x_fields.nRec IN x_cols
				m.lcCond = m.paRows[m.lnRowCnt,10] + ' AND ' + ;
							m.paCols[m.lnColCnt,10]
			OTHERWISE
				m.lcCond = m.paRows[m.lnRowCnt,10]
			ENDCASE

			SELECT (.cCurrAlias)
			LOCATE FOR &lcCond
		ENDIF
		m.lnRecNo = RECNO(.cCurrAlias)
		m.luValue = EVALUATE(m.lcValue)
		IF m.lnRecNo <> RECNO(.cCurrAlias)
			IF m.lnRecNo <= RECCOUNT(.cCurrAlias)
				GOTO m.lnRecNo IN .cCurrAlias
			ELSE
				GO BOTTOM
				SKIP
			ENDIF
		ENDIF
	ENDIF

	.nRow = 0
	.nCol = 0
	IF .lStopKey
		RETURN .F.
	ENDIF

	SELECT x_fields
	DO CASE
	CASE VARTYPE(m.luValue) = 'C'
		m.lcValue = TRIM(m.luValue)
		IF x_fields.lRowTitle OR x_fields.lColTitle
			IF LEFT(m.lcValue,1) == '\'
				m.lcValue = SUBSTR(m.lcValue,3)
			ENDIF
			IF RIGHT(m.lcValue,2) == '\*'
				m.lcValue = LEFT(m.lcValue,LEN(m.lcValue)-2)
			ENDIF
		ENDIF
	CASE VARTYPE(m.luValue) = 'N'
		m.lcValue = LTRIM(STR(m.luValue,25,2))
	CASE VARTYPE(m.luValue) = 'D'
		m.lcValue = DTOC(m.luValue)
	OTHERWISE
		m.lcValue = ''
	ENDCASE

	m.lnRow = m.tnDstR1 + x_fields.nRow - m.tnSrcR1
	m.lnCol = x_fields.nCol
	IF x_fields.lAll
		IF VARTYPE(m.luValue) = 'N'
			* Вроде безсмысленно, но иначе в ячейках вместо нулей
			* числа почти равные нулю (порядка 1E-13, 1E-14, ...)
			IF m.luValue = 0
				m.luValue = 0
			ENDIF
			.NumberRC(m.lnRow,m.lnCol) = m.luValue
		ELSE
			.TextRC(m.lnRow,m.lnCol) = m.lcValue
		ENDIF
	ELSE
		m.lcText = .TextRC(m.lnRow,m.lnCol)
		m.lnPos = AT('[', m.lcText, 1)
		m.lnLen = AT(']', m.lcText, 1) - m.lnPos + 1
		m.lcText = STUFF(m.lcText, m.lnPos, m.lnLen, m.lcValue)
		.TextRC(m.lnRow,m.lnCol) = m.lcText
	ENDIF
ENDSCAN

SELECT (m.lcAlias)
RETURN .T.
ENDFUNC


FUNCTION GenGrRows
LOCAL m.i, m.j, m.lnRowCnt, m.lcName, m.lnRecNo, m.lcAlias
LOCAL m.lnR1, m.lnGr1, m.lnLen
LOCAL m.luNext, m.luPrev, m.llGrGen

m.lcAlias = ALIAS()
m.lnRowCnt = ALEN(m.paRows,1)

FOR m.i = m.lnRowCnt-1 TO 2 STEP -1
	m.lcName = m.paRows[m.i,1]
	m.luNext = EVALUATE(m.lcAlias+'.'+m.lcName)
	m.luPrev = m.paRows[m.i,2]

	m.llGrGen = (m.luNext <> m.luPrev)
	m.j = m.i
	DO WHILE NOT m.llGrGen AND m.j > 2
		m.j = m.j - 1
		m.lcName = m.paRows[m.j,1]
		m.llGrGen = (EVALUATE(m.lcAlias+'.'+m.lcName) <> m.paRows[m.j,2])
	ENDDO

	IF m.llGrGen
		m.llGrGen = NOT ( .nPrintRowBands = 2 AND m.paRows[m.i,8] < 2 OR ;
			TYPE("m.luPrev") = 'C' AND RIGHT(TRIM(m.luPrev),2) = '\*' )

		m.lnLen = m.paRows[m.i,4]
		IF m.lnLen > 0 AND m.llGrGen
			m.lnGr1 = m.paRows[m.i,3]
			m.lnR1 = m.paRows[m.i,7]
			.nRowGr = m.i
			.CopyRows(m.lnR1,m.lnGr1,m.lnLen)
			.nRowGr = 0
			IF .lStopKey
				RETURN .F.
			ENDIF

			m.lnRecNo = RECNO(m.lcAlias)
			GOTO m.paRows[m.i,9] IN (m.lcAlias)
			.GenFields(m.lnR1, m.lnGr1, m.lnLen, m.i)
			IF m.lnRecNo <= RECCOUNT(m.lcAlias)
				GOTO m.lnRecNo IN (m.lcAlias)
			ELSE
				GO BOTTOM IN (m.lcAlias)
				SKIP IN (m.lcAlias)
			ENDIF
			IF .lStopKey
				RETURN .F.
			ENDIF
		ENDIF

		m.lnLen = m.paRows[m.i,6]
		IF m.lnLen > 0 AND m.llGrGen
			m.lnGr1 = m.paRows[m.i,5]
			m.lnR1 = .nCurrRow
			.nRowGr = -m.i
			.CopyRows(m.lnR1,m.lnGr1,m.lnLen)
			.nRowGr = 0
			IF .lStopKey
				RETURN .F.
			ENDIF

			SKIP -1 IN (m.lcAlias)
			.GenFields(m.lnR1, m.lnGr1, m.lnLen, m.i)
			SKIP IN (m.lcAlias)
			IF .lStopKey
				RETURN .F.
			ENDIF
		ENDIF

		m.paRows[m.i,2] = m.luNext
		m.paRows[m.i,7] = .nCurrRow
		m.paRows[m.i,8] = 0
		m.paRows[m.i,9] = RECNO(m.lcAlias)
	ENDIF
	m.paRows[m.i,8] = m.paRows[m.i,8] + 1
ENDFOR

SELECT (m.lcAlias)
RETURN .T.
ENDFUNC


FUNCTION GenRows
LOCAL m.i, m.lnRowCnt
LOCAL m.lnR1, m.lnR2, m.lnDet1, m.lnLen, m.lnGr1, m.lnLen1

.Sheet = .nSheet

m.lnRowCnt = ALEN(m.paRows,1)
m.lnDet1 = m.paRows[m.lnRowCnt,3]
m.lnLen  = m.paRows[m.lnRowCnt,4]

.cCurrAlias = m.paRows[1,1]
.nCurrRow = m.paRows[1,5]

IF .lPivotTable
	SELECT x_rows
ELSE
	SELECT (.cCurrAlias)
ENDIF
GOTO TOP

FOR m.i = 1 TO m.lnRowCnt
	m.paRows[m.i,7] = .nCurrRow
	m.paRows[m.i,8] = 0
	m.paRows[m.i,9] = RECNO()
ENDFOR

m.lnGr1 = m.paRows[1,3]
m.lnLen1 = m.paRows[1,4]
IF m.lnLen1 > 0
	.GenFields(m.lnGr1, m.lnGr1, m.lnLen1, 1)
	IF .lStopKey
		RETURN .F.
	ENDIF
ENDIF

IF m.lnRowCnt = 1
	RETURN .T.
ENDIF

m.lnGr1 = m.paRows[1,5]
m.lnLen1 = m.paRows[1,6]
IF m.lnLen1 > 0
	.GenFields(m.lnGr1, m.lnGr1, m.lnLen1, 1)
	IF .lStopKey
		RETURN .F.
	ENDIF
ENDIF

SCAN
	IF .nPrintRowBands > 0
		.GenGrRows()
		IF .lStopKey
			RETURN .F.
		ENDIF
	ENDIF

	IF m.lnLen > 0
		m.lnR1 = .nCurrRow
		.nRowGr = m.lnRowCnt
		.CopyRows(m.lnR1,m.lnDet1,m.lnLen)
		.nRowGr = 0
		IF .lStopKey
			RETURN .F.
		ENDIF

		.GenFields(m.lnR1, m.lnDet1, m.lnLen, m.lnRowCnt)
		IF .lStopKey
			RETURN .F.
		ENDIF
	ENDIF
ENDSCAN

IF .nPrintRowBands > 0
	.GenGrRows()
	IF .lStopKey
		RETURN .F.
	ENDIF
ENDIF

m.lnR1 = m.paRows[1,3] + m.paRows[1,4]
m.lnR2 = m.paRows[1,5]
.DeleteRange(m.lnR1,-1,m.lnR2-1,-1,3)
.nCurrRow = .nCurrRow - (m.lnR2 - m.lnR1)

RETURN .T.
ENDFUNC


FUNCTION CopyRows
LPARAMETERS m.tnR1, m.tnS1, m.tnLen
LOCAL m.lnR2, m.lnS2, m.lnC2, m.lnGr
m.lnR2 = m.tnR1 + m.tnLen - 1
m.lnS2 = m.tnS1 + m.tnLen - 1
m.lnC2 = .LastCol
m.lnGr = .nRowGr
.nRowGr = -1
.InsertRange(m.tnR1,-1,m.lnR2,-1,3)
.nRowGr = m.lnGr
IF .lStopKey
	RETURN .F.
ENDIF
FOR m.i = m.tnR1 TO m.lnR2
	.RowHeight(m.i) = .RowHeight(m.tnS1+m.i-m.tnR1)
ENDFOR
.CopyRange(m.tnR1,1,m.lnR2,m.lnC2,.SS,m.tnS1,1,m.lnS2,m.lnC2)
.nCurrRow = .nCurrRow + m.tnLen
RETURN .T.
ENDFUNC


FUNCTION CopyCols
LPARAMETERS m.tnR1, m.tnR2, m.tnC1, m.tnD1, m.tnLen
LOCAL m.lnC2, m.lnD2, m.lnGr
m.lnC2 = m.tnC1 + m.tnLen - 1
m.lnD2 = m.tnD1 + m.tnLen - 1
m.lnGr = .nColGr
.nColGr = -1
.InsertRange(m.tnR1,m.tnC1,m.tnR2,m.lnC2,1)
.nColGr = m.lnGr
IF .lStopKey
	RETURN .F.
ENDIF
FOR m.i = m.tnC1 TO m.lnC2
	IF m.i >= .nPrevCol
		.ColWidth(m.i) = .ColWidth(m.tnD1+m.i-m.tnC1)
	ENDIF
ENDFOR
.CopyRange(m.tnR1,m.tnC1,m.tnR2,m.lnC2,.SS,m.tnR1,m.tnD1,m.tnR2,m.lnD2)
.nCurrCol = .nCurrCol + m.tnLen
RETURN .T.
ENDFUNC


FUNCTION GenGrCols
LOCAL m.i, m.j, m.k, m.lnColCnt, m.lcName, m.lnRecNo, m.lcAlias
LOCAL m.lnR1, m.lnR2, m.lnC1, m.lnGr1, m.lnLen
LOCAL m.luNext, m.luPrev, m.llGrGen
LOCAL ARRAY laTemp(1,1)

m.lcAlias = ALIAS()
m.lnColCnt = ALEN(m.paCols,1)
m.lnR1 = m.paRows[1,3]
m.lnR2 = m.paRows[1,5] + m.paRows[1,6] - 1

FOR m.i = m.lnColCnt-1 TO 2 STEP -1
	m.lcName = m.paCols[m.i,1]
	m.luNext = EVALUATE(m.lcAlias+'.'+m.lcName)
	m.luPrev = m.paCols[m.i,2]

	m.llGrGen = (m.luNext <> m.luPrev)
	m.j = m.i
	DO WHILE NOT m.llGrGen AND m.j > 2
		m.j = m.j - 1
		m.lcName = m.paCols[m.j,1]
		m.llGrGen = (EVALUATE(m.lcAlias+'.'+m.lcName) <> m.paCols[m.j,2])
	ENDDO

	IF m.llGrGen
		m.llGrGen = NOT ( .nPrintColBands = 2 AND m.paCols[m.i,8] < 2 OR ;
			TYPE("m.luPrev") = 'C' AND RIGHT(TRIM(m.luPrev),2) = '\*' )

		m.lnLen = m.paCols[m.i,4]
		IF m.lnLen > 0 AND m.llGrGen
			m.lnGr1 = m.paCols[m.i,3]
			m.lnC1 = m.paCols[m.i,7]
			.nColGr = m.i
			.CopyCols(m.lnR1,m.lnR2,m.lnC1,m.lnGr1,m.lnLen)
			.nColGr = 0
			IF .lStopKey
				RETURN .F.
			ENDIF

			SELECT x_fields
			REPLACE ALL nCol WITH nCol+m.lnLen FOR lPivot AND nCol >= m.lnC1

			m.k = m.lnC1 - m.lnGr1
			m.j = RECNO('x_cols') - 1
			SELECT nRow, nCol+m.k AS nCol, m.j AS nRec, ;
				nGr, cExpr, lAll, lPivot, lRowTitle, lColTitle, lHeader ;
				FROM x_temp WHERE nGr = m.i AND lHeader ;
			INTO ARRAY laTemp
			INSERT INTO x_fields FROM ARRAY laTemp
		ENDIF

		m.lnLen = m.paCols[m.i,6]
		IF m.lnLen > 0 AND m.llGrGen
			m.lnGr1 = m.paCols[m.i,5]
			m.lnC1 = .nCurrCol
			.nColGr = -m.i
			.CopyCols(m.lnR1,m.lnR2,m.lnC1,m.lnGr1,m.lnLen)
			.nColGr = 0
			IF .lStopKey
				RETURN .F.
			ENDIF

			m.k = m.lnC1 - m.lnGr1
			m.j = RECNO('x_cols') - 1
			SELECT nRow, nCol+m.k AS nCol, m.j AS nRec, ;
				nGr, cExpr, lAll, lPivot, lRowTitle, lColTitle, lHeader ;
				FROM x_temp WHERE nGr = m.i AND NOT lHeader ;
			INTO ARRAY laTemp
			INSERT INTO x_fields FROM ARRAY laTemp
		ENDIF

		m.paCols[m.i,2] = m.luNext
		m.paCols[m.i,7] = .nCurrCol
		m.paCols[m.i,8] = 0
	ENDIF
	m.paCols[m.i,8] = m.paCols[m.i,8] + 1
ENDFOR

SELECT (m.lcAlias)
RETURN .T.
ENDFUNC


FUNCTION GenCols
LOCAL m.i, m.j, m.k, m.lnColCnt
LOCAL m.lnR1, m.lnR2, m.lnC1, m.lnC2, m.lnDet1, m.lnLen
LOCAL ARRAY laTemp(1,1)

.Sheet = .nSheet

m.lnColCnt = ALEN(m.paCols,1)
m.lnC1 = m.paCols[1,3] + m.paCols[1,4]
m.lnC2 = m.paCols[1,5] - 1
m.lnR1 = m.paRows[1,3] + m.paRows[1,4]

SELECT x_fields
REPLACE lPivot WITH .T. FOR nCol >= m.lnC1 AND nCol <= m.lnC2
REPLACE lColTitle WITH .T. FOR lPivot AND nRow < m.lnR1
REPLACE lRowTitle WITH .T. FOR nCol < m.lnC1

m.lnR1 = m.paRows[1,3]
m.lnR2 = m.paRows[1,5] + m.paRows[1,6] - 1

FOR m.i = 1 TO m.lnColCnt
	m.lnC1 = m.paCols[m.i,3]
	m.lnC2 = m.lnC1 + m.paCols[m.i,4] - 1
	REPLACE nGr WITH m.i, lHeader WITH .T. ;
		FOR nCol >= m.lnC1 AND nCol <= m.lnC2
	m.lnC1 = m.paCols[m.i,5]
	m.lnC2 = m.lnC1 + m.paCols[m.i,6] - 1
	REPLACE nGr WITH m.i FOR nCol >= m.lnC1 AND nCol <= m.lnC2
ENDFOR

SELECT * FROM x_fields WHERE lPivot INTO CURSOR x_temp READWRITE

DELETE FOR lPivot IN x_fields

.cCurrAlias = m.paCols[1,1]
.nCurrCol = m.paCols[1,5]

m.lnDet1 = m.paCols[m.lnColCnt,3]
m.lnLen  = m.paCols[m.lnColCnt,4]

SELECT x_cols
GOTO TOP

FOR m.i = 1 TO m.lnColCnt
	m.paCols[m.i,7] = .nCurrCol
	m.paCols[m.i,8] = 0
ENDFOR

IF m.lnColCnt = 1
	RETURN .T.
ENDIF

SCAN
	IF .nPrintColBands > 0
		.GenGrCols()
		IF .lStopKey
			RETURN .F.
		ENDIF
	ENDIF

	IF m.lnLen > 0
		m.lnC1 = .nCurrCol
		.nColGr = m.lnColCnt
		.CopyCols(m.lnR1,m.lnR2,m.lnC1,m.lnDet1,m.lnLen)
		.nColGr = 0
		IF .lStopKey
			RETURN .F.
		ENDIF

		m.k = m.lnC1 - m.lnDet1
		m.j = RECNO('x_cols')
		SELECT nRow, nCol+m.k AS nCol, m.j AS nRec, ;
			nGr, cExpr, lAll, lPivot, lRowTitle, lColTitle, lHeader ;
			FROM x_temp WHERE nGr = m.lnColCnt ;
		INTO ARRAY laTemp
		INSERT INTO x_fields FROM ARRAY laTemp
	ENDIF
ENDSCAN

IF .nPrintColBands > 0
	.GenGrCols()
	IF .lStopKey
		RETURN .F.
	ENDIF
ENDIF

USE IN x_temp

m.lnC1 = m.paCols[1,3] + m.paCols[1,4]
m.lnC2 = m.paCols[1,5]
.DeleteRange(m.lnR1,m.lnC1,m.lnR2,m.lnC2-1,1)

m.k = (m.lnC2 - m.lnC1)
.nCurrCol = .nCurrCol - m.k
SELECT x_fields
REPLACE ALL nCol WITH nCol-m.k FOR lPivot

m.k = .nCurrCol - m.lnC2
SELECT x_fields
REPLACE ALL nCol WITH nCol+m.k FOR NOT lPivot AND nCol >= m.lnC2

RETURN .T.
ENDFUNC


FUNCTION MergeRows
LOCAL m.lnR1, m.lnR2, m.lnRmax, m.lnC1, m.lcText, m.loCellFormat
m.lnRmax = .nCurrRow
.nColGr = 1

SELECT nCol, MIN(nRow) AS nR1 ;
	FROM x_fields WHERE lRowTitle GROUP BY nCol ;
	INTO CURSOR x_temp

SELECT x_temp
SCAN
	m.lnC1 = x_temp.nCol
	m.lnR1 = x_temp.nR1
	m.lcText = .TextRC(m.lnR1, m.lnC1)
	m.lnR2 = m.lnR1
	DO WHILE m.lnR2 < m.lnRmax
		m.lnR2 = m.lnR2 + 1
		IF m.lnR2 = m.lnRmax OR NOT (m.lcText == .TextRC(m.lnR2, m.lnC1))
			IF m.lnR2-1 > m.lnR1
				.SetSelection(m.lnR1, m.lnC1, m.lnR2-1, m.lnC1)
				m.loCellFormat = .GetCellFormat()
				m.loCellFormat.MergeCells = .T.
				.SetCellFormat(m.loCellFormat)
			ENDIF
			m.lnR1 = m.lnR2
			m.lcText = .TextRC(m.lnR1, m.lnC1)
		ENDIF
	ENDDO
ENDSCAN

USE IN x_temp

.nColGr = 0
RETURN .T.
ENDFUNC


FUNCTION MergeCols
LOCAL m.lnC1, m.lnC2, m.lnCmax, m.lnR1, m.lcText, m.loCellFormat
.nRowGr = 1

SELECT nRow, MIN(nCol) AS nC1, MAX(nCol) AS nC2 ;
	FROM x_fields WHERE lColTitle GROUP BY nRow ;
	INTO CURSOR x_temp

SELECT x_temp
SCAN
	m.lnR1 = x_temp.nRow
	m.lnC1 = x_temp.nC1
	m.lnCmax = x_temp.nC2 + 1
	m.lcText = .TextRC(m.lnR1, m.lnC1)
	m.lnC2 = m.lnC1
	DO WHILE m.lnC2 < m.lnCmax
		m.lnC2 = m.lnC2 + 1
		IF m.lnC2 = m.lnCmax OR NOT (m.lcText == .TextRC(m.lnR1, m.lnC2))
			IF m.lnC2-1 > m.lnC1
				.SetSelection(m.lnR1, m.lnC1, m.lnR1, m.lnC2-1)
				m.loCellFormat = .GetCellFormat()
				m.loCellFormat.MergeCells = .T.
				.SetCellFormat(m.loCellFormat)
			ENDIF
			m.lnC1 = m.lnC2
			m.lcText = .TextRC(m.lnR1, m.lnC1)
		ENDIF
	ENDDO
ENDSCAN

USE IN x_temp

.nRowGr = 0
RETURN .T.
ENDFUNC


FUNCTION FitRowsCols
LOCAL m.lnR1, m.lnC1, m.lnR2, m.lnC2, m.loFind

m.loFind = .DefineSearch('{<Auto}', .nSheet, 1, 1, MAX(.LastRow, 1), 1, 16)
IF m.loFind.FindNext()
	m.lnR1 = m.loFind.Row
ELSE
	m.lnR1 = m.paRows[1,3] + m.paRows[1,4]
ENDIF
m.loFind = .DefineSearch('{Auto>}', .nSheet, 1, 1, MAX(.LastRow, 1), 1, 16)
IF m.loFind.FindNext()
	m.lnR2 = m.loFind.Row
ELSE
	m.lnR2 = .nCurrRow + m.paRows[1,6] - 1
ENDIF

m.loFind = .DefineSearch('{<Auto}', .nSheet, 1, 1, 1, MAX(.LastCol, 1), 16)
IF m.loFind.FindNext()
	m.lnC1 = m.loFind.Col
ELSE
	IF .lPivotTable
		m.lnC1 = m.paCols[1,3] + m.paCols[1,4]
	ELSE
		m.lnC1 = 1
	ENDIF
ENDIF
m.loFind = .DefineSearch('{Auto>}', .nSheet, 1, 1, 1, MAX(.LastCol, 1), 16)
IF m.loFind.FindNext()
	m.lnC2 = m.loFind.Col
ELSE
	m.lnC2 = .LastCol
ENDIF

IF .nAutoFitRows > 0 AND m.lnR2 >= m.lnR1
	.SetRowHeightAuto(m.lnR1, m.lnC1, m.lnR2, m.lnC2, (.nAutoFitRows=1))
ENDIF
IF .nAutoFitCols > 0 AND m.lnC2 >= m.lnC1 AND .lPivotTable
	.SetColWidthAuto(m.lnR1, m.lnC1, m.lnR2, m.lnC2, (.nAutoFitCols=1))
ENDIF

RETURN .T.
ENDFUNC


PROCEDURE OpenBlock
DIMENSION m.paRows[1,10], m.paCols[1,10]
.nRowGr = 0
.nColGr = 0
.nRow = 0
.nCol = 0

.lPivotTable = .F.
.lAliasUsed = .F.
.cCurrAlias = ''

CREATE CURSOR x_fields ;
	( nRow N(5), nCol N(3), nRec N(5), nGr N(3), cExpr C(254), ;
	lAll L, lPivot L, lRowTitle L, lColTitle L, lHeader L )
ENDPROC


PROCEDURE CloseBlock
IF TYPE('m.paRows[1,5]') <> 'N'
	m.paRows[1,5] = 0
ENDIF
.nShift = .nShift + .nCurrRow - m.paRows[1,5]
.nPrevCol = .nCurrCol

USE IN x_fields
IF USED('x_rows')
	USE IN x_rows
ENDIF
IF USED('x_cols')
	USE IN x_cols
ENDIF
IF USED('x_temp')
	USE IN x_temp
ENDIF
IF .lBlockKey AND NOT .lAliasUsed AND USED(.cCurrAlias)
	USE IN (.cCurrAlias)
ENDIF
ENDPROC


FUNCTION ClearBlock
LOCAL m.lnR1, m.lnR2

* Очистка управляющей колонки
m.lnR1 = m.paRows[1,3]
m.lnR2 = .nCurrRow + m.paRows[1,6] - 1
.ClearRange(m.lnR1,1,m.lnR2,1,1)

IF .lPivotTable
	* Удаление управляющей строки
	.DeleteRange(m.lnR1,-1,m.lnR1,-1,3)
	IF .FixedRow+.FixedRows > m.lnR1
		.FixedRows = .FixedRows - 1
	ENDIF
	.nCurrRow = .nCurrRow - 1
ENDIF

RETURN .T.
ENDFUNC


FUNCTION TestBlock
	IF .lStopKey OR NOT .TestGroups('x_rows', @m.paRows)
		RETURN .F.
	ENDIF
	IF .lPivotTable
		IF .lStopKey OR NOT .TestGroups('x_cols', @m.paCols)
			RETURN .F.
		ENDIF
	ENDIF
	IF .lStopKey OR NOT .TestParams()
		RETURN .F.
	ENDIF
	IF .lStopKey OR NOT .TestFields()
		RETURN .F.
	ENDIF
RETURN .T.
ENDFUNC


FUNCTION GenBlock

* Формирование колонок
IF .lPivotTable
	.GenCols()
	IF .lStopKey
		RETURN .F.
	ENDIF
ENDIF

*.LaunchDesigner()

* Формирование строк
.GenRows()
IF .lStopKey
	RETURN .F.
ENDIF

*.LaunchDesigner()

* Автоподбор высоты строк и ширины колонок
IF .nAutoFitRows > 0 OR .nAutoFitCols > 0
	.FitRowsCols()
ENDIF

* Объединение одинаковых ячеек в заголовках строк
IF .nMergeRowTitles = 1 AND .lPivotTable
	.MergeRows()
ENDIF

* Объединение одинаковых ячеек в заголовках колонок
IF .nMergeColTitles = 1 AND .lPivotTable
	.MergeCols()
ENDIF

* Очистка текущего блока
.ClearBlock()

*.LaunchDesigner()

SELECT 0

RETURN .T.
ENDFUNC


FUNCTION ReadBook
LPARAMETERS m.tcName
LOCAL m.lnFileType
m.lnFileType = 0
m.lnFileType = .ReadEx(m.tcName)
IF m.lnFileType <> 11
	OutError('Ошибка при открытии шаблона: ' + m.tcName)
	RETURN .F.
ENDIF
RETURN .T.
ENDFUNC


FUNCTION WriteBook
LPARAMETERS m.tcName
.SaveWindowInfo()
.WriteEx(m.tcName,11)
RETURN .T.
ENDFUNC


FUNCTION ClearBook
LOCAL m.i

* Цикл по листам
FOR m.i = 1 TO .NumSheets
	.Sheet = m.i

	* Удаление управляющей колонки
	IF NOT .lStopKey
		.DeleteRange(-1,1,-1,1,4)
		IF .FixedCols > 0
			.FixedCols = .FixedCols - 1
		ENDIF
	ENDIF

	* Установка активной ячейки
	IF .FixedCols > 0
		.Col = .FixedCol + .FixedCols
	ELSE
		.Col = 1
	ENDIF
	IF .FixedRows > 0
		.Row = .FixedRow + .FixedRows
	ELSE
		.Row = 1
	ENDIF
	.ShowActiveCell()
ENDFOR

.Sheet = 1

RETURN .T.
ENDFUNC


FUNCTION GenBook
LPARAMETERS m.tcFrmName, m.tcRepName, m.tlDesigner, m.tcUDFName
LOCAL m.i
WITH THIS
	* Чтение шаблона
	.ReadBook(m.tcFrmName)
	IF .lStopKey
		RETURN .F.
	ENDIF

	* Цикл по листам
	FOR m.i = 1 TO .NumSheets
		.nSheet = m.i
		.nShift = 0
		.nPrevCol = 0
		.lBlockKey = .T.
		* Цикл по блокам на листе
		DO WHILE .lBlockKey AND NOT .lStopKey
			.OpenBlock()
			IF .TestBlock()
				.GenBlock()
			ENDIF
			.CloseBlock()
		ENDDO
		IF .lStopKey
			EXIT
		ENDIF
	ENDFOR

	* Очистка книги отчета
	.ClearBook()

	IF NOT .lStopKey AND NOT EMPTY(m.tcUDFName)
		* Вызов функции пользователя
		DO &tcUDFName WITH THIS
	ENDIF

	IF m.tlDesigner
		* Вызов конструктора F1
		.LaunchDesigner()
		* .AboutBox()
	ENDIF

	* Запись отчета
	.WriteBook(m.tcRepName)

	RETURN NOT .lStopKey
ENDWITH
ENDFUNC

ENDDEFINE
