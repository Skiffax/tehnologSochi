secondsTip:=1
SetdefaultMouseSpeed 2
stringbCAD:="bCAD"
Stringbcadexe:="b[cC][aA][dD]\.*"
stringbCADtire:="bCAD -"
XImage:=0
YImage:=0
SetTitleMatchMode, 2
MatchMode:=2

waittime:=10
waittimeotladka:=1
waittimealternative:=1000
waittimewindow:=3
waittimewindow10:=10
waittimewindow1:=150

;~ ExcelWindowName:="�����8"
;~ WinActivate, %ExcelWindowName%
;~ oExcel := ComObjGet(ExcelWindowName)
;~ ;��������� ��������� ����������� ������
;~ oExcel.ActiveSheet.Range("A1").Value:="���.��������� ����������"
;~ maxRow:=oExcel.ActiveSheet.UsedRange.Row  + oExcel.ActiveSheet.UsedRange.Rows.Count - 1
;~ MsgBox, maxRow %maxRow%
;~ ExitApp











;WinActivate, ahk_exe %Stringbcadexe%
;~ WinActivate, �����
	 ;~ controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
;~ Smeta:="�����"
;~ IDactive:=WinActive("A")
 ;~ try {
    ;~ ControlSend,%controlNeed%,{Space},ahk_id %IDactive%
	;~ } catch e {
		;~ Msgbox, ErrorRuntime %e%
	;~ }
	
;~ ExitApp00v


 ;~ controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
 ;~ Smeta:="�����"
 ;~ ControlSend,%controlNeed%,{Space},%Smeta%
 ;~ if (ErrorLevel)
	 ;~ MsgBox,%ErrorLevel%
 ;~ ExitApp

SetTitleMatchMode, RegEx
;WinActivate, ahk_exe %Stringbcadexe%

;~ controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
;~ Smeta:="�����"
;~ ControlSend,%controlNeed%,{Space},%Smeta%
;~ if (ErrorLevel)
	;~ MsgBox,%ErrorLevel%
;~ ExitApp

IfWinExist, ahk_exe %Stringbcadexe%
{
	
	;������ ������ ���� Excel (����� ��������� ������������ ������ � ����� "����� ")
	maxNumberExcel:=0
	SetTitleMatchMode, RegEx
	SetTitleMatchMode, %MatchMode%
	SetTitleMatchMode, Slow
	poiskOkno:="Excel.exe"
	WinGet, allwindows, List, ahk_exe %poiskOkno% ;�������� ��� ���� �������� excel.exe
	poiskOknoRaschet:=[] ;� ���� ������� �������� ������� ��� ��������� ������, �� ������� ������ ����� ������������� ����� ���������� ��������
	poiskOknoRaschet[1]:="�����������"
	poiskOknoRaschet[2]:="���������"
	kolvo1:=poiskOknoRaschet.MaxIndex() ;���������� ��������� � �������
	findstring1:="�����"
	Loop, %allwindows%
	{
		;Sleep, 100
		this_id := allwindows%A_Index%
		WinGetTitle, this_title, ahk_id %this_id%
		
		if (InStr(this_title,findstring1))
		{
			position1:=StrLen(findstring1)
			findstring2:=" "
			position2:=InStr(this_title,findstring2,false,position1)
			position1:=position1+1
			
			dlinaStroki:=position2-position1
			
			tempNumberExcel:=SubStr(this_title,position1,dlinaStroki)
			;MsgBox, tempNumberExcel %tempNumberExcel%
			if (tempNumberExcel>maxNumberExcel)
			{
				maxNumberExcel:=tempNumberExcel
			}
		}
		else
		{
			;������� ������ ������������
			Loop, %kolvo1%
			{
				tempPoisk:= % poiskOknoRaschet[A_Index]
				;MsgBox, tempPoisk %tempPoisk%
				;exit
				naideno:=InStr(this_title,tempPoisk)
				if (naideno<>0)
				{
					TrayTip , ������� ����� bCAD, �������� ��������� ����-Excel ������, ����� ���������� ������� �� ����� �����������, %secondsTip%
					raschetExcelTitle:=this_title
				}
			}
		}
	}
	;���������, ��� �� ������� ���� ���� �� �������������. ���� ������-�� ������� ���, �� ���������������� ���� ����� ����� ���� ������
	tempNumberExcel:=0
	;MsgBox, maxNumberExcel %maxNumberExcel%
	Loop, %maxNumberExcel%
	{
		
		tempfindstring:="�����" A_Index
		naideno:=0
		Loop, %allwindows%
		{
			this_id := allwindows%A_Index% ;�������� ID ���� Excel
			WinGetTitle, this_title, ahk_id %this_id% ;�������� ��������� ���� Excel
			;MsgBox, findstring1 %findstring1% this_title %this_title%

			if (InStr(this_title,tempfindstring)=0) ;���� � ��������� ������
			{
				;tempNumberExcel:=A_Index
				;MsgBox, tempNumberExcel %tempNumberExcel%
				;GoTo ZavershenaProverkaPervogoOknaExcel
			}
			else
			{
				Naideno:=1
			}
		}
		If (naideno=0) ;���� ������ �� ���� �������, �� A_Index �������� Loop �������� ������ �����
		{
			maxNumberExcel:=A_Index-1 ;��������� �� 1, ������ ��� ����� �������� �� 1
			Goto ZavershenaProverkaPervogoOknaExcel
		}
	}
	
	
ZavershenaProverkaPervogoOknaExcel:
	;~ if (tempNumberExcel<>0)
		;~ maxNumberExcel:=tempNumberExcel
	maxNumberExcel:=maxNumberExcel+1
	;MsgBox, maxNumberExcel %maxNumberExcel%
	;MsgBox, maxNumberExcel %maxNumberExcel%
	ExcelWindowName:=findstring1 maxNumberExcel
	;ExcelSmetaWindowName:=findstring1 maxNumberExcel
	maxNumberExcelOtchet:=maxNumberExcel+1
	ExcelOtchetKonstruktoraName:=findstring1 maxNumberExcelOtchet
	
	
	
	
	SetTitleMatchMode, RegEx
	WinActivate, ahk_exe %Stringbcadexe%
	;�������� ID ����, ����� ���� ����� �������� � ���
	WinGet, IDbCAD, ID, ahk_exe %Stringbcadexe%
	SetTitleMatchMode, %MatchMode%
	SetTitleMatchMode, Slow
	;������� �����
	errorfirst:=false

	Send, {Alt}
Probuemeshe:
	Sleep,200
	Send,0v
	SetTitleMatchMode, 1
	titleWait:="��� ��������"
	WinWait, %titleWait%,,1
	If (ErrorLevel)
		Send, �
	WinWait, %titleWait%,,%waittimewindow10%
	If (ErrorLevel)
	{
		if (errorfirst=false)
		{
			errorfirst:=true
			Send, 0
;			MsgBox, ololo
			Goto Probuemeshe
		}
		else
		{
			MsgBox, �����-�� �������� � ����� %titleWait%. ��������� ����������
			ExitApp
		}
	}
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d4" ;������ ���������� � ������� ����� ������
	;WindowsForms10.BUTTON.app.0.33c0d9d6 ;������ �������
	;ControlSend,%controlNeed%,{Space},%titleWait%
	ControlFocus,%controlNeed%,%titleWait%
	Send, {Space}
	sleep, waittime
	;ControlGetFocus,currentFocus,A
	;if (currentFocus=controlNeed)
	;MsgBox, eeee currentFocus %currentFocus%
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3" ;������ �� � ������� ����� ������
	;ControlSend,%controlNeed%,{Enter},%titleWait%
	ControlFocus,%controlNeed%,%titleWait%
	Send, {Space}
	;��� ���� �����
	titleWindowWait:="�����"
	SetTitleMatchMode, 1
	ToolTip, ��� ��������� ���� Excel
	WinWaitClose,%titleWait%,,waittimewindow1
	WinWait,%titleWindowWait%,,waittimewindow1
	if ErrorLevel
	{
		MsgBox, ������ �� �������� ���� %titleWindowWait%. ���������� ��������� ���� %titleWindowWait% ������� � ����� ������ ������ �� � ���� ���������
	}
	IDSmeta:=WinExist("A")
	Sleep, 100
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3" ;������ "������� � ������ � ������� �����"
	;SetTitleMatchMode, 2
	WinGetTitle, out1, ahk_id %IDSmeta%
	;MsgBox, nameActive %out1% IDSmeta %IDSmeta%
	;WinActivate, ahk_id %IDactive%
	;controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
	;Smeta:="�����"
	cou1:=0
	SetTitleMatchMode, 1
PovtorControlSend:
	;WinActivate, ahk_id %IDSmeta%
	;WinActivate, ahk_id %IDbCAD% 
	;ControlSend, %controlNeed%, {Space}, ahk_id %IDSmeta%
	ControlSend, %controlNeed%, {Space}, A
	if (ErrorLevel)
	{
		;MsgBox, ErrorControlSend ExportSmetabCADtoExcel
		ControlFocus, %controlNeed%, ahk_id %IDSmeta%
		if (ErrorLevel) ;���� ������, ������� ��� �����
		{
			if (cou1>20)
			{
				ToolTip
				MsgBox, ErrorControlSend ExportSmetabCADtoExcel
				ExitApp
			}
			Sleep, 1000
			cou1++
			WinGetTitle, out1, ahk_id %IDSmeta%
			WinGetTitle, out2, A
			ToolTip, %cou1% Title: %out1% A %out2%
			Goto PovtorControlSend
		}
		Send, {Space}
	}
	ToolTip
	
	SetTitleMatchMode, 2
	WinWait,%ExcelWindowName%,,%waittimewindow1%
	If (ErrorLevel)
	{
		MsgBox, ������ �� �������� ���� %titleWindowWait%. ���������� ��������� ��� � ������������� �������. ����� ����� ������� ������ "��"
	}
	
	;WinWaitActive,%ExcelWindowName%,,waittimewindow1
	WinWaitActive,%ExcelWindowName%,,%waittimewindow1%
	if (ErrorLevel)
	{
		MsgBox, ������ �� �������� ���� Excel %ExcelWindowName% � �� ���� ������������� �� ���� �������������. ��������������, ��� ���� Excel %ExcelWindowName% ���������, ������������� �� ���� � ������� ������ �� � ���� �������.
		WinGetTitle, currentActiveWindow, ahk_id WinActive("A")
		if (InStr(currentActiveWindow,ExcelWindowName)=0)
			WinActivate, %ExcelWindowName%
	}
	
	ToolTip, ��������� �����
	;��������� �����
	WinActivate, ahk_id %IDbCAD% ;������������� �� bCAD
	;WinActivate, %titleWindowWait%
	IDactive:=WinExist("A")
	Sleep, 2000
	
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d4" ;��������� �����
	ControlFocus, %controlNeed%, ahk_id %IDSmeta%
	ControlSend, %controlNeed%, {Space}, ahk_id %IDSmeta%
	Sleep, 1000
	;Send, {Space}
	ControlSend, %controlNeed%, {Enter}, %titleWindowWait%
	Sleep, 1000
	WinWaitClose, %titleWindowWait%,,%waittimewindow1%
	if (ErrorLevel)
	{
		MsgBox, �����-�� �������� � ��������� ���� �����. ���������� ������� ���� ����� �������
		WinWaitClose, %titleWindowWait%,,%waittimewindow1%
	}
	WinActivate, ahk_id %IDbCAD%

	;MsgBox,������� ������
	ToolTip, ������������ �����������������
	;������� ������ ������������
	Send, {Alt}
	Sleep,200
	Send,0w
	SetTitleMatchMode, 1
	titleWait:="��� ��������"
	WinWait, %titleWait%,,1
	If (ErrorLevel)
		Send, �
	WinWait, %titleWait%,,%waittimewindow1%
	If (ErrorLevel)
	{
		MsgBox, �����-�� �������� � ����� %titleWait%. ��������� ����������
		ExitApp
	}
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d5" ;������ ����������
	;WindowsForms10.BUTTON.app.0.33c0d9d6 ;������ �������
	ControlSend,%controlNeed%,{Space},%titleWait%
	sleep, waittime
	ControlGetFocus,currentFocus,A
	;if (currentFocus=controlNeed)
	;MsgBox, eeee currentFocus %currentFocus%
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d4" ;������ ��
	ControlSend,%controlNeed%,{Enter},%titleWait%
	
	;��� ���� ����� ������������
	titleWindowWait:="����� ������������"
	titlebCADOtchet:=titleWindowWait
	SetTitleMatchMode, 1
	waittimewindow1:=20
	ToolTip, ��� ��������� ���� %titleWindowWait%
	WinWait,%titleWindowWait%,,waittimewindow1
	if ErrorLevel
	{
		MsgBox, ������ �� �������� ���� %titleWindowWait%. ���������� ��������� ���� %titleWindowWait% ������� � ����� ������ ������ �� � ���� ���������
	}
	Sleep, waittime
	ToolTip, ��� ������� Excel
	;� ���� �� ���� �������� ������ ������� � ������
	WinActivate, %titleWindowWait%
	
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d6" ;������ "������� � ������"
	ControlFocus, %controlNeed%, %titleWindowWait%
	Sleep, waittime
	if (ErrorLevel=1)
	{
		ControlGetText, OutputVar1,%controlNeed%
		MsgBox, Some problem with: %OutputVar1% (%controlNeed%) in Window %titleWindowWait%
		Exit
	}
	Sleep, waittime
	ControlSend, %controlNeed%, {Enter}, %titleWindowWait%
	;SendInput, {Enter}
	Sleep, 2000 ;sswww
	
	SetTitleMatchMode, 1
	SetTitleMatchMode, Slow
	WinWait,%ExcelOtchetKonstruktoraName%,,180
	if ErrorLevel
	{
		MsgBox, ������ �� �������� ���� %ExcelOtchetKonstruktoraName%. ���������� ��������� ��������� ����, ����� ������������� �� ���� � ����� ������ ������ "��" �����. ������ ��������� ��������
	}
	waittimewindow2:=2
	WinWaitActive, %ExcelOtchetKonstruktoraName%,,waittimewindow2
	if ErrorLevel
	{
		;MsgBox, ������ �� �������� ���� %ExcelOtchetKonstruktoraName%. ���������� ������������ ��� ���� �������
		WinActivate, %ExcelOtchetKonstruktoraName%
	}
	
	Sleep, waittime
	ToolTip, ������� ��������...
	ToolTip, ������������ � Excel
		
	SetTitleMatchMode, RegEx
	;WinActivate, ahk_exe %Stringbcadexe% ; %Stringbcad% 11.01.2022
	WinMinimize, ahk_exe %Stringbcadexe% ; %Stringbcad% ;
	SetTitleMatchMode, 2
	SetTitleMatchMode, Slow
	
	Sleep, waittime
	WinActivate, %ExcelOtchetKonstruktoraName%
	
	
	oExcelKonstruktor := ComObjGet(ExcelOtchetKonstruktoraName) ;������������ � Excel ������ ������������
	kolvoListovExcel:=oExcelKonstruktor.Sheets.Count
	NameFindListExcel:="�������"
	indexFindListExcel:=0
	kolvoRowsExcel:=0
	listProfilivExcelNaiden:=0
	loop, %kolvoListovExcel%
	{
		tempName:=oExcelKonstruktor.Sheets.Item(A_Index).Name
		if (tempName=NameFindListExcel)
		{
			indexFindListExcel:=A_Index
			listProfilivExcelNaiden:=indexFindListExcel
		}
	}
	if (indexFindListExcel) ;���� ���� ��� ������, �� ������������� �� ���� � �������� ������
	{
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Activate
		;������� ��������� �� ������ ������ � ������� "�"
		lRow := 1
		rowsDlyaProverki:=1000
		Loop, %rowsDlyaProverki%  ;�������� �������� �� 2000 ������
		{
			tempProverkaNaPustotu:=oExcelKonstruktor.Sheets.Item(NameFindListExcel).Range("A" A_Index).Value
			if (tempProverkaNaPustotu)
			{
			}
			else
			{
				kolvoRowsExcel:=A_Index-1
				GoTo ZavershenaProverkaPoslednegoRowExcel
			}
			ToolTip, % ExcelOtchetKonstruktoraName ", ������: " A_Index
		}
		ZavershenaProverkaPoslednegoRowExcel:
		SetKeyDelay, 50
		Sleep, 50
		stringRangeExcelProfili:="A1:N" kolvoRowsExcel
		kolvoRowsExcel--
		;MsgBox, kolvoRowsExcel ����� ������������ %kolvoRowsExcel%
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Range(stringRangeExcelProfili).Copy
		;~ Send, ^a
		;~ sleep, 50
		;~ Send, ^c
		;~ sleep, 50
	kolvoRowsExcelOtchetKonstruktora:=kolvoRowsExcel
	}
	
	
	WinActivate, %ExcelWindowName% ;������������ � ��������� ���� Excel
	
	if (listProfilivExcelNaiden)
	{
		ToolTip, ��������� ������� � �����
		oExcel := ComObjGet(ExcelWindowName)
		;��������� ��������� ����������� ������
		maxRow:=oExcel.ActiveSheet.UsedRange.Row  + oExcel.ActiveSheet.UsedRange.Rows.Count - 1
		;MsgBox, maxRowSmeta %maxRow%
		maxRowExcelSmeta:=maxRow
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		Sleep, waittime
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Value:="������� ���������� �������"
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Select
		Sleep, waittime
		oExcel.ActiveSheet.Paste
	}
	
	;oExcel := ComObjGet(ExcelOtchetKonstruktoraName)
	kolvoListovExcel:=oExcelKonstruktor.Sheets.Count
	NameFindListExcel:="������"
	indexFindListExcel:=0
	kolvoRowsExcel:=0
	listProfilivExcelNaiden:=0
	loop, %kolvoListovExcel%
	{
		tempName:=oExcelKonstruktor.Sheets.Item(A_Index).Name
		if (tempName=NameFindListExcel)
		{
			indexFindListExcel:=A_Index
			listProfilivExcelNaiden:=indexFindListExcel
		}
	}
	if (indexFindListExcel) ;���� ���� ��� ������, �� ������������� �� ���� � �������� ������
	{
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Activate
		;������� ��������� �� ������ ������ � ������� "�"
		lRow := 1
		rowsDlyaProverki:=1000
		Loop, %rowsDlyaProverki%  ;�������� �������� �� 2000 ������
		{
			tempProverkaNaPustotu:=oExcelKonstruktor.Sheets.Item(NameFindListExcel).Range("A" A_Index).Value
			if (tempProverkaNaPustotu)
			{
			}
			else
			{
				kolvoRowsExcel:=A_Index-1
				GoTo ZavershenaProverkaPoslednegoRowExcel1
			}
			ToolTip, % ExcelOtchetKonstruktoraName ", ������: " A_Index
		}
		ZavershenaProverkaPoslednegoRowExcel1:
		SetKeyDelay, 50
		Sleep, 50
		stringRangeExcelProfili:="A1:N" kolvoRowsExcel
		kolvoRowsExcel--
		;MsgBox, kolvoRowsExcel ����� ������������ %kolvoRowsExcel%
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Range(stringRangeExcelProfili).Copy
		;~ Send, ^a
		;~ sleep, 50
		;~ Send, ^c
		;~ sleep, 50
	}
	
	WinActivate, %ExcelWindowName% ;������������ � ��������� ���� Excel
	if (listProfilivExcelNaiden)
	{
		ToolTip, ��������� ������� � �����
		oExcel := ComObjGet(ExcelWindowName)
		maxRow:=oExcel.ActiveSheet.UsedRange.Row  + oExcel.ActiveSheet.UsedRange.Rows.Count - 1
		;MsgBox, maxRowSmeta %maxRow%
		maxRowExcelSmeta:=maxRow
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		Sleep, waittime
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Value:="���.��������� ����������"
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Select
		Sleep, waittime
		oExcel.ActiveSheet.Paste
	}
	
	;������������� �� bCAD � ��������� ���� ������ ������������
	;WinActivate, ahk_id %IDbCAD% ;������������� �� bCAD
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d1" ;������ ����� � ���� ������ ������������
	ControlSend, %controlNeed%, {Space}, %titlebCADOtchet%
	WinWaitClose, %titlebCADOtchet%, 4
	WinActivate, %ExcelWindowName%
	Sleep, 100
	Send, ^a
	Send, ^c
}
else
{
	MsgBox, bCAD �� �������. ��������� ���������
}