;������ ������ ���� Excel (����� ��������� ������������ ������ � ����� "����� ")
SetTimer, AvariynoeZavershenie, 15 ;��������� ������� Esc
maxNumberExcel:=0
MatchMode:=1
waitExcelClose:=1.5
;SetTitleMatchMode, RegEx
SetTitleMatchMode, %MatchMode%
SetTitleMatchMode, Slow
poiskOkno:="Excel.exe"
ToolTip, �������� ���������� - Esc
Sleep, 500
SetTimer, RemoveToolTip, -1000
WinGet, allwindows, List, ����� ahk_exe %poiskOkno% ;�������� ��� ���� �������� excel.exe
WinGet, kolvowindows, Count, ����� ahk_exe %poiskOkno%
;~ poiskOknoRaschet:=[] ;� ���� ������� �������� ������� ��� ��������� ������, �� ������� ������ ����� ������������� ����� ���������� ��������
;~ poiskOknoRaschet[1]:="�����������"
;~ poiskOknoRaschet[2]:="���������"
;~ kolvo1:=poiskOknoRaschet.MaxIndex() ;���������� ��������� � �������
;findstring1:="�����"
;string1:=""
if (kolvowindows)
	ToolTip, ���������� ����: %kolvowindows% ��
else
{
	ToolTip, ���������� ���� Excel �� ����������
	SetTimer, RemoveToolTip, -1000
	Sleep, 1000
}

Loop, %allwindows%
{
	;Sleep, 100
	this_id := allwindows%A_Index%
	WinGetTitle, this_title, ahk_id %this_id%
	ToolTip, ��������� %this_title%
	;WinClose, %this_title%
	WinActivate, ahk_id %this_id%
	Sleep, 200
	PostMessage, 0x0112, 0xF060,,, %this_title%
	Sleep, 100
	ControlSend,,{Right},ahk_id %this_id%
	ControlSend,,{Space},ahk_id %this_id%
	WinWaitClose, ahk_id %this_id%,,%waitExcelClose%
	If (ErrorLevel)
	{
		ToolTip, �������������� ����� �������� (� ������� ������� ������ � ������ ������)
		Sleep, 100
		Send, {Right}
		Send, {Space}
		WinWaitClose, ahk_id %this_id%,,%waitExcelClose%
		If (ErrorLevel)
		{
			ToolTip, �� ������� ������� �������������
			Sleep, 500
		}
	}
	Sleep, 100
	;IDChild:=WinExist("ahk_class NetUIHWND")
	;~ MsgBox, IDChild %IDChild%
	;~ hParent := DllCall("GetParent", Ptr, IDChild, Ptr)
	;~ WinGetTitle, Title, ahk_id %hParent%
	;~ MsgBox, %Title%
}

ToolTip, �� �������
Sleep, 500
ToolTip

ExitApp

AvariynoeZavershenie:
GetKeyState, StateKey, Esc
if (StateKey="D")
{
	ToolTip
	BlockInput, Off
	MsgBox, �������� ���� Excel �������� �������������
	SetTitleMatchMode, RegEx
	WinSet, Enable,, ahk_exe %StringbcadexeReg%
	SetTitleMatchMode, 2
	SetTitleMatchMode, Fast	
	ExitApp
}
return

RemoveToolTip:
ToolTip
return