************************
Procedure ErrorHandler (tnErrorNo, tcMessage, tcErrorLine, tcModule, tnErrorLineNo)
************************
*!* Routine slightly modified from the one shown in
*!* the book 'Special Edition Using Visual FoxPro 6'
*!* http://docs.rinet.ru/GlyadiLisu/ch23/ch23.htm

Local lcError, lnExitMethod
Local lcChkDBC, lcCurDBC, lcErrorFile, lcSuffix
LOCAL lnCnt, lnWhichTrigger
LOCAL ARRAY laErrorArray[1]

= Aerror(laErrorArray)

Set DataSession To

m.lnExitMethod = 0

* Avoid recursive loop if errorlog contains an error
m.lcError = On('ERROR')
On Error *

* Each case in this structure represents one error type
* It handles trivial errors first, followed by recoverable
* errors. Finally, all other errors generate an ASCII text
* file with information about the system and error.
Do Case

*** Check for trivial errors
* Check if beyond end of file, place on last record
Case m.tnErrorNo = 4
	Goto Bottom

* Check if before beginning of file, place on first record
Case m.tnErrorNo = 38
	Goto Top

* Cannot pack a cursor
Case m.tnErrorNo = 1115

* Check for Resume without Suspend
Case m.tnErrorNo = 1236
*!*	*** Unrecoverable Errors
*!*	* Redirect output to a file
Otherwise
	m.lnExitMethod = 2
* Get a file name based on date and time
	m.lcErrorFile = ADDBS(SYS(2023)) + Substr(Dtoc(Date()), 1, 2) + ;
		SUBSTR(Dtoc(Date()), 4, 2) + ;
		SUBSTR(Time(), 1, 2) + ;
		SUBSTR(Time(), 4, 2) + '.ERR'
* Make sure the file name is unique by changing the extension
	m.lcSuffix = '0'
	Do While File(m.lcErrorFile)
		m.lcErrorFile = Stuff(m.lcErrorFile, ;
			LEN(m.lcErrorFile) - Len(m.lcSuffix) + 1, ;
			LEN(m.lcSuffix), m.lcSuffix)
		m.lcSuffix    = Alltrim(Str(Val(m.lcSuffix)+1, 3))
	Enddo
	Set Console Off
	Set Alternate To (m.lcErrorFile)
	Set Alternate On
* Identify error
	? 'DATE:         ' + Ttoc(Datetime())
	? 'COMPUTER AND USERNAME:     ' + Sys(0)
	? 'VERSION:      ' + Version()
	? 'FILE NAME:    ' + m.lcErrorFile

* Next identify the error
	? 'Error:'
*!*		= Aerror(laErrorArray)
	? '    Number: ' + Str(laErrorArray[1], 5)
	? '   Message: ' + laErrorArray[2]

	If !Isnull(laErrorArray[5])
		? ' Parameter: ' + laErrorArray[3]
	Endif

	If !Isnull(laErrorArray[5])
		? ' Work Area: ' + laErrorArray[4]
	Endif

	If !Isnull(laErrorArray[5]) AND VARTYPE(laErrorArray[5])=='N'
		m.lnWhichTrigger = laErrorArray[5]
		Do Case
		Case m.lnWhichTrigger = 1
			? ' Insert Trigger Failed'
		Case m.lnWhichTrigger = 2
			? ' Update Trigger Failed'
		Case m.lnWhichTrigger = 3
			? ' Delete Trigger Failed'
		Endcase
	Endif

	If laErrorArray[1] = m.tnErrorNo
		? '    Module: ' + m.tcModule
		? '      Line: ' + m.tcErrorLine
		? '    Line #: ' + Str(m.tnErrorLineNo)
	Endif
	Release laErrorArray, lnWhichTrigger
	?

* Next identify the basic operating environment
	? 'OP. SYSTEM:     ' + Os()
	? 'PROCESSOR:      ' + Sys(17)
	? 'GRAPHICS:       ' + Left(Sys(2006), At('/', Sys(2006)) - 1)
	? 'MONITOR:        ' + Substr(Sys(2006), At('/', Sys(2006)) + 1)
	? 'RESOURCE FILE:  ' + Sys(2005)
	? 'LAUNCH DIR:     ' + Sys(2004)
	? 'CONFIG.FP:      ' + Sys(2019)
	? 'MEMORY:         ' + Alltrim(Str(Memory())), 'KB OR ' + SYS(12) + 'BYTES'
	? 'CONVENTIONAL:   ' + Sys(12)
	? 'TOTAL MEMORY:   '
	? 'EMS LIMIT:      ' + Sys(24)
	? 'CTRLABLE MEM:   ' + Sys(1016)
	? 'CURRENT CONSOLE:' + Sys(100)
	? 'CURRENT DEVICE: ' + Sys(101)
	? 'CURRENT PRINTER:' + Sys(102)
	? 'CURRENT DIR:    ' + Sys(2003)
	? 'LAST KEY:       ' + Str(Lastkey(),5)
	?

* Next identify the default disk drive and its properties
	? '  DEFAULT DRIVE: ' + Sys(5)
	? '     DRIVE SIZE: ' + Transform(Val(Sys(2020)), '999,999,999')
	? '     FREE SPACE: ' + Transform(Diskspace(), '999,999,999')
	? '    DEFAULT DIR: ' + Curdir()
	? ' TEMP FILES DIR: ' + Sys(2023)
	?

* Available Printers
	? 'PRINTERS:'
	If Aprinters(laPrt) > 0
		For m.lnCnt = 1 To Alen(laPrt, 1)
			? Padr(laPrt[m.lnCnt,1], 50) + ' ON ' + ;
				PADR(laPrt[m.lnCnt,2], 25)
		Endfor
	Else
		? 'No printers currently defined.'
	Endif
	?

* Define Workareas
	? 'WORK AREAS:'
	If Aused(laWrkAreas) > 0
		= Asort(laWrkAreas,2)
		List Memory Like laWrkAreas
		Release laWrkAreas
		? 'Current Database: ' + Alias()
	Else
		? 'No tables currently open in any work areas.'
	Endif
	?

* Begin bulk information dump
* Display memory variables
	? Replicate('-', 78)
	? 'ACTIVE MEMORY VARIABLES'
	List Memory
	?

* Display status
	? Replicate('-', 78)
	? 'CURRENT STATUS AND SET VARIABLES'
	List Status
	?

* Display Information related to databases
	If Adatabase(laDbList) > 0
		lcCurDBC = Juststem(Dbc())
		For m.lnCnt = 1 To Alen(laDbList, 1)
			m.lcChkDBC = laDbList[m.lnCnt, 1]
			Set Database To (m.lcChkDBC)
			List Connections
			?
			List Database
			?
			List Procedures
			?
			List Tables
			?
			List Views
			?
		Endfor
		If !Empty(lcCurDBC)
			Set Database To (lcCurDBC)
		Endif
	Endif

* Close error file and reactivate the screen
	Set Alternate To
	Set Alternate Off
	Set Console On
***CLOSE down open forms
	For Each loForm In _Screen.Forms
		If Type("loform.parent") = "O" &&Handle formsets
			loForm.Parent.Visible = .F.
		Else
			loForm.Visible = .F.
		Endif
	Next
	m.lcErrorInformation = Filetostr(m.lcErrorFile)
	Erase (m.lcErrorFile)
	*!* Call the issues form
	Do Form issues With m.lcErrorInformation
	*!* Now kill this process using Doug's routine
	*!* so no more errors continue to happen
	KillProcess(_Screen.Caption)
	Return
Endcase
On Error &lcError
Endfunc &&showerror

********************************
Function KillProcess
*==============================================================================
* Program:			KillProcess
* Purpose:			Terminate the specified application
* Author:			Doug Hennig
* Copyright:		(c) 2001 Stonefield Systems Group Inc.
* Last revision:	02/02/2001
* Parameters:		tcCaption - the caption for the application to terminate
* Returns:			.T. if it succeeded
* Environment in:	none
* Environment out:	if successful, the application has been terminated
*==============================================================================

Lparameters tcCaption
Local lnhWnd, ;
	llReturn, ;
	lnProcessID, ;
	lnHandle

* Declare the Win32API functions we need.

#Define WM_DESTROY 0x0002
Declare Integer FindWindow In Win32API ;
	string @cClassName, String @cWindowName
Declare Integer SendMessage In Win32API ;
	integer HWnd, Integer uMsg, Integer wParam, Integer Lparam
Declare Sleep In Win32API ;
	integer nMilliseconds
Declare Integer GetWindowThreadProcessId In Win32API ;
	integer HWnd, Integer @lpdwProcessId
Declare Integer OpenProcess In Win32API ;
	integer dwDesiredAccess, Integer bInheritHandle, Integer dwProcessID
Declare Integer TerminateProcess In Win32API ;
	integer hProcess, Integer uExitCode

* Get a handle to the window by its caption.

lnhWnd   = FindWindow(0, tcCaption)
llReturn = lnhWnd = 0

* If we found the window, send a "destroy" message to it, then wait for it to
* be gone. If it didn't, let's use the big hammer: we'll terminate its process.

If Not llReturn
	SendMessage(lnhWnd, WM_DESTROY, 0, 0)
	llReturn = WaitForAppTermination(tcCaption)
	If Not llReturn
		lnProcessID = 0
		GetWindowThreadProcessId(lnhWnd, @lnProcessID)
		lnHandle = OpenProcess(1, 1, lnProcessID)
		llReturn = TerminateProcess(lnHandle, 0) > 0
	Endif Not llReturn
Endif Not llReturn
Return llReturn
Endfunc

* For up to five times, wait for a second, then see if the specified application
* is still running. Return .T. if the application has terminated.

********************************
Function WaitForAppTermination
********************************
Lparameters tcCaption
Local lnCounter, llReturn
m.lnCounter = 0
m.llReturn  = .F.
Do While ! m.llReturn And lnCounter < 5
	Sleep(1000)
	m.lnCounter = lnCounter + 1
	m.llReturn  = FindWindow(0, m.tcCaption) = 0
Enddo
Return m.llReturn
ENDFUNC

*!* Function declarations to avoid compiler errors
*!* because the compiler isn't aware of the functions
*!* that are contained in the FLL.
FUNCTION EMCREATEMESSAGE
FUNCTION EMADDRECIPIENT
FUNCTION EMSEND