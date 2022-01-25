;начало поиска окон Excel (чтобы вычислить максимальный индекс в имени "Книга ")
SetTimer, AvariynoeZavershenie, 15 ;обработка нажатия Esc
maxNumberExcel:=0
MatchMode:=1
waitExcelClose:=1.5
;SetTitleMatchMode, RegEx
SetTitleMatchMode, %MatchMode%
SetTitleMatchMode, Slow
poiskOkno:="Excel.exe"
ToolTip, Прервать выполнение - Esc
Sleep, 500
SetTimer, RemoveToolTip, -1000
WinGet, allwindows, List, Книга ahk_exe %poiskOkno% ;получаем все окна процесса excel.exe
WinGet, kolvowindows, Count, Книга ahk_exe %poiskOkno%
;~ poiskOknoRaschet:=[] ;в этом массиве написаны шаблоны имён расчётных файлов, на которые скрипт будет переключаться после завершения экспорта
;~ poiskOknoRaschet[1]:="ТестМакроса"
;~ poiskOknoRaschet[2]:="Стоимость"
;~ kolvo1:=poiskOknoRaschet.MaxIndex() ;количество элементов в массиве
;findstring1:="Книга"
;string1:=""
if (kolvowindows)
	ToolTip, Обнаружено окон: %kolvowindows% шт
else
{
	ToolTip, Подходящие окна Excel не обнаружены
	SetTimer, RemoveToolTip, -1000
	Sleep, 1000
}

Loop, %allwindows%
{
	;Sleep, 100
	this_id := allwindows%A_Index%
	WinGetTitle, this_title, ahk_id %this_id%
	ToolTip, Закрываем %this_title%
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
		ToolTip, Альтернативный метод закрытия (с большим объёмом данных в буфере обмена)
		Sleep, 100
		Send, {Right}
		Send, {Space}
		WinWaitClose, ahk_id %this_id%,,%waitExcelClose%
		If (ErrorLevel)
		{
			ToolTip, Не удалось закрыть автоматически
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

ToolTip, Всё закрыли
Sleep, 500
ToolTip

ExitApp

AvariynoeZavershenie:
GetKeyState, StateKey, Esc
if (StateKey="D")
{
	ToolTip
	BlockInput, Off
	MsgBox, Закрытие окон Excel прервано пользователем
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