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

;~ ExcelWindowName:="Книга8"
;~ WinActivate, %ExcelWindowName%
;~ oExcel := ComObjGet(ExcelWindowName)
;~ ;вычисляем последнюю заполненную ячейку
;~ oExcel.ActiveSheet.Range("A1").Value:="ДОП.КРОМОЧНАЯ ИНФОРМАЦИЯ"
;~ maxRow:=oExcel.ActiveSheet.UsedRange.Row  + oExcel.ActiveSheet.UsedRange.Rows.Count - 1
;~ MsgBox, maxRow %maxRow%
;~ ExitApp











;WinActivate, ahk_exe %Stringbcadexe%
;~ WinActivate, Смета
	 ;~ controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
;~ Smeta:="Смета"
;~ IDactive:=WinActive("A")
 ;~ try {
    ;~ ControlSend,%controlNeed%,{Space},ahk_id %IDactive%
	;~ } catch e {
		;~ Msgbox, ErrorRuntime %e%
	;~ }
	
;~ ExitApp00v


 ;~ controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
 ;~ Smeta:="Смета"
 ;~ ControlSend,%controlNeed%,{Space},%Smeta%
 ;~ if (ErrorLevel)
	 ;~ MsgBox,%ErrorLevel%
 ;~ ExitApp

SetTitleMatchMode, RegEx
;WinActivate, ahk_exe %Stringbcadexe%

;~ controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
;~ Smeta:="Смета"
;~ ControlSend,%controlNeed%,{Space},%Smeta%
;~ if (ErrorLevel)
	;~ MsgBox,%ErrorLevel%
;~ ExitApp

IfWinExist, ahk_exe %Stringbcadexe%
{
	
	;начало поиска окон Excel (чтобы вычислить максимальный индекс в имени "Книга ")
	maxNumberExcel:=0
	SetTitleMatchMode, RegEx
	SetTitleMatchMode, %MatchMode%
	SetTitleMatchMode, Slow
	poiskOkno:="Excel.exe"
	WinGet, allwindows, List, ahk_exe %poiskOkno% ;получаем все окна процесса excel.exe
	poiskOknoRaschet:=[] ;в этом массиве написаны шаблоны имён расчётных файлов, на которые скрипт будет переключаться после завершения экспорта
	poiskOknoRaschet[1]:="ТестМакроса"
	poiskOknoRaschet[2]:="Стоимость"
	kolvo1:=poiskOknoRaschet.MaxIndex() ;количество элементов в массиве
	findstring1:="Книга"
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
			;попытка поиска тестаМакроса
			Loop, %kolvo1%
			{
				tempPoisk:= % poiskOknoRaschet[A_Index]
				;MsgBox, tempPoisk %tempPoisk%
				;exit
				naideno:=InStr(this_title,tempPoisk)
				if (naideno<>0)
				{
					TrayTip , Экспорт сметы bCAD, Открытый расчётный файл-Excel найден, после завершения скрипта он будет активирован, %secondsTip%
					raschetExcelTitle:=this_title
				}
			}
		}
	}
	;проверяем, все ли индексы книг есть до максимального. Если какого-то индекса нет, то экспортированный файл будет иметь этот индекс
	tempNumberExcel:=0
	;MsgBox, maxNumberExcel %maxNumberExcel%
	Loop, %maxNumberExcel%
	{
		
		tempfindstring:="Книга" A_Index
		naideno:=0
		Loop, %allwindows%
		{
			this_id := allwindows%A_Index% ;получаем ID окна Excel
			WinGetTitle, this_title, ahk_id %this_id% ;получаем заголовок окна Excel
			;MsgBox, findstring1 %findstring1% this_title %this_title%

			if (InStr(this_title,tempfindstring)=0) ;ищем в заголовке строку
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
		If (naideno=0) ;если ничего не было найдено, то A_Index текущего Loop является первым окном
		{
			maxNumberExcel:=A_Index-1 ;уменьшаем на 1, потому что далее увеличим на 1
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
	;получаем ID окна, чтобы было проще работать с ним
	WinGet, IDbCAD, ID, ahk_exe %Stringbcadexe%
	SetTitleMatchMode, %MatchMode%
	SetTitleMatchMode, Slow
	;экспорт сметы
	errorfirst:=false

	Send, {Alt}
Probuemeshe:
	Sleep,200
	Send,0v
	SetTitleMatchMode, 1
	titleWait:="Тип объектов"
	WinWait, %titleWait%,,1
	If (ErrorLevel)
		Send, м
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
			MsgBox, Какая-то проблема с окном %titleWait%. Прерываем выполнение
			ExitApp
		}
	}
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d4" ;только помеченные в диалоге перед Сметой
	;WindowsForms10.BUTTON.app.0.33c0d9d6 ;только видимые
	;ControlSend,%controlNeed%,{Space},%titleWait%
	ControlFocus,%controlNeed%,%titleWait%
	Send, {Space}
	sleep, waittime
	;ControlGetFocus,currentFocus,A
	;if (currentFocus=controlNeed)
	;MsgBox, eeee currentFocus %currentFocus%
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3" ;Кнопка Ок в диалоге перед Сметой
	;ControlSend,%controlNeed%,{Enter},%titleWait%
	ControlFocus,%controlNeed%,%titleWait%
	Send, {Space}
	;ждём окно Смета
	titleWindowWait:="Смета"
	SetTitleMatchMode, 1
	ToolTip, Ждём появления окна Excel
	WinWaitClose,%titleWait%,,waittimewindow1
	WinWait,%titleWindowWait%,,waittimewindow1
	if ErrorLevel
	{
		MsgBox, Скрипт не дождался окна %titleWindowWait%. Попробуйте запустить окно %titleWindowWait% вручную и затем нажать кнопку ОК в этом сообщении
	}
	IDSmeta:=WinExist("A")
	Sleep, 100
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3" ;кнопка "Экспорт в эксель в диалоге сметы"
	;SetTitleMatchMode, 2
	WinGetTitle, out1, ahk_id %IDSmeta%
	;MsgBox, nameActive %out1% IDSmeta %IDSmeta%
	;WinActivate, ahk_id %IDactive%
	;controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d3"
	;Smeta:="Смета"
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
		if (ErrorLevel) ;если ошибка, пробуем ещё разок
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
		MsgBox, Скрипт не дождался окна %titleWindowWait%. Попробуйте дождаться его и переключиться вручную. После этого нажмите кнопку "Ок"
	}
	
	;WinWaitActive,%ExcelWindowName%,,waittimewindow1
	WinWaitActive,%ExcelWindowName%,,%waittimewindow1%
	if (ErrorLevel)
	{
		MsgBox, Скрипт не дождался окна Excel %ExcelWindowName% и не смог переключиться на него автоматически. Удостоверьтесь, что окно Excel %ExcelWindowName% появилось, переключитесь на него и нажмите кнопку Ок в этом диалоге.
		WinGetTitle, currentActiveWindow, ahk_id WinActive("A")
		if (InStr(currentActiveWindow,ExcelWindowName)=0)
			WinActivate, %ExcelWindowName%
	}
	
	ToolTip, закрываем смету
	;закрываем смету
	WinActivate, ahk_id %IDbCAD% ;переключаемся на bCAD
	;WinActivate, %titleWindowWait%
	IDactive:=WinExist("A")
	Sleep, 2000
	
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d4" ;закрываем смету
	ControlFocus, %controlNeed%, ahk_id %IDSmeta%
	ControlSend, %controlNeed%, {Space}, ahk_id %IDSmeta%
	Sleep, 1000
	;Send, {Space}
	ControlSend, %controlNeed%, {Enter}, %titleWindowWait%
	Sleep, 1000
	WinWaitClose, %titleWindowWait%,,%waittimewindow1%
	if (ErrorLevel)
	{
		MsgBox, Какая-то проблема с закрытием окна сметы. Попробуйте закрыть окно сметы вручную
		WinWaitClose, %titleWindowWait%,,%waittimewindow1%
	}
	WinActivate, ahk_id %IDbCAD%

	;MsgBox,Экспорт отчёта
	ToolTip, Экспортируем отчётКонструктора
	;экспорт отчёта конструктора
	Send, {Alt}
	Sleep,200
	Send,0w
	SetTitleMatchMode, 1
	titleWait:="Тип объектов"
	WinWait, %titleWait%,,1
	If (ErrorLevel)
		Send, ц
	WinWait, %titleWait%,,%waittimewindow1%
	If (ErrorLevel)
	{
		MsgBox, Какая-то проблема с окном %titleWait%. Прерываем выполнение
		ExitApp
	}
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d5" ;только помеченные
	;WindowsForms10.BUTTON.app.0.33c0d9d6 ;только видимые
	ControlSend,%controlNeed%,{Space},%titleWait%
	sleep, waittime
	ControlGetFocus,currentFocus,A
	;if (currentFocus=controlNeed)
	;MsgBox, eeee currentFocus %currentFocus%
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d4" ;Кнопка Ок
	ControlSend,%controlNeed%,{Enter},%titleWait%
	
	;ждём окно Отчёт конструктора
	titleWindowWait:="Отчёт конструктора"
	titlebCADOtchet:=titleWindowWait
	SetTitleMatchMode, 1
	waittimewindow1:=20
	ToolTip, Ждём появления окна %titleWindowWait%
	WinWait,%titleWindowWait%,,waittimewindow1
	if ErrorLevel
	{
		MsgBox, Скрипт не дождался окна %titleWindowWait%. Попробуйте запустить окно %titleWindowWait% вручную и затем нажать кнопку ОК в этом сообщении
	}
	Sleep, waittime
	ToolTip, Ждём экспорт Excel
	;в этом же окне нажимаем кнопку Экспорт в эксель
	WinActivate, %titleWindowWait%
	
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d6" ;кнопка "Экспорт в эксель"
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
		MsgBox, Скрипт не дождался окна %ExcelOtchetKonstruktoraName%. Попробуйте дождаться появления окна, затем переключиться на него и затем нажать кнопку "ОК" здесь. Скрипт продолжит действие
	}
	waittimewindow2:=2
	WinWaitActive, %ExcelOtchetKonstruktoraName%,,waittimewindow2
	if ErrorLevel
	{
		;MsgBox, Скрипт не дождался окна %ExcelOtchetKonstruktoraName%. Попробуйте активировать это окно вручную
		WinActivate, %ExcelOtchetKonstruktoraName%
	}
	
	Sleep, waittime
	ToolTip, Процесс экспорта...
	ToolTip, Подключаемся к Excel
		
	SetTitleMatchMode, RegEx
	;WinActivate, ahk_exe %Stringbcadexe% ; %Stringbcad% 11.01.2022
	WinMinimize, ahk_exe %Stringbcadexe% ; %Stringbcad% ;
	SetTitleMatchMode, 2
	SetTitleMatchMode, Slow
	
	Sleep, waittime
	WinActivate, %ExcelOtchetKonstruktoraName%
	
	
	oExcelKonstruktor := ComObjGet(ExcelOtchetKonstruktoraName) ;подключаемся к Excel отчёту конструктора
	kolvoListovExcel:=oExcelKonstruktor.Sheets.Count
	NameFindListExcel:="Профили"
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
	if (indexFindListExcel) ;если лист был найден, то переключаемся на него и копируем данные
	{
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Activate
		;Находим последнюю не пустую ячейку в столбце "А"
		lRow := 1
		rowsDlyaProverki:=1000
		Loop, %rowsDlyaProverki%  ;проводим проверку до 2000 строки
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
			ToolTip, % ExcelOtchetKonstruktoraName ", Строка: " A_Index
		}
		ZavershenaProverkaPoslednegoRowExcel:
		SetKeyDelay, 50
		Sleep, 50
		stringRangeExcelProfili:="A1:N" kolvoRowsExcel
		kolvoRowsExcel--
		;MsgBox, kolvoRowsExcel отчёт конструктора %kolvoRowsExcel%
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Range(stringRangeExcelProfili).Copy
		;~ Send, ^a
		;~ sleep, 50
		;~ Send, ^c
		;~ sleep, 50
	kolvoRowsExcelOtchetKonstruktora:=kolvoRowsExcel
	}
	
	
	WinActivate, %ExcelWindowName% ;подключаемся к основному окну Excel
	
	if (listProfilivExcelNaiden)
	{
		ToolTip, Вставляем профили в Смету
		oExcel := ComObjGet(ExcelWindowName)
		;вычисляем последнюю заполненную ячейку
		maxRow:=oExcel.ActiveSheet.UsedRange.Row  + oExcel.ActiveSheet.UsedRange.Rows.Count - 1
		;MsgBox, maxRowSmeta %maxRow%
		maxRowExcelSmeta:=maxRow
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		Sleep, waittime
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Value:="РАСКРОЙ ПРОФИЛЬНЫХ ДЕТАЛЕЙ"
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Select
		Sleep, waittime
		oExcel.ActiveSheet.Paste
	}
	
	;oExcel := ComObjGet(ExcelOtchetKonstruktoraName)
	kolvoListovExcel:=oExcelKonstruktor.Sheets.Count
	NameFindListExcel:="Кромки"
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
	if (indexFindListExcel) ;если лист был найден, то переключаемся на него и копируем данные
	{
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Activate
		;Находим последнюю не пустую ячейку в столбце "А"
		lRow := 1
		rowsDlyaProverki:=1000
		Loop, %rowsDlyaProverki%  ;проводим проверку до 2000 строки
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
			ToolTip, % ExcelOtchetKonstruktoraName ", Строка: " A_Index
		}
		ZavershenaProverkaPoslednegoRowExcel1:
		SetKeyDelay, 50
		Sleep, 50
		stringRangeExcelProfili:="A1:N" kolvoRowsExcel
		kolvoRowsExcel--
		;MsgBox, kolvoRowsExcel отчёт конструктора %kolvoRowsExcel%
		oExcelKonstruktor.Sheets.Item(NameFindListExcel).Range(stringRangeExcelProfili).Copy
		;~ Send, ^a
		;~ sleep, 50
		;~ Send, ^c
		;~ sleep, 50
	}
	
	WinActivate, %ExcelWindowName% ;подключаемся к основному окну Excel
	if (listProfilivExcelNaiden)
	{
		ToolTip, Вставляем профили в Смету
		oExcel := ComObjGet(ExcelWindowName)
		maxRow:=oExcel.ActiveSheet.UsedRange.Row  + oExcel.ActiveSheet.UsedRange.Rows.Count - 1
		;MsgBox, maxRowSmeta %maxRow%
		maxRowExcelSmeta:=maxRow
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		Sleep, waittime
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Value:="ДОП.КРОМОЧНАЯ ИНФОРМАЦИЯ"
		maxRowExcelSmeta++
		;MsgBox, maxRowExcelSmeta %maxRowExcelSmeta%
		stringmaxRowExcelSmeta:="A" maxRowExcelSmeta
		oExcel.ActiveSheet.Range(stringmaxRowExcelSmeta).Select
		Sleep, waittime
		oExcel.ActiveSheet.Paste
	}
	
	;переключаемся на bCAD и закрываем окно отчёта конструктора
	;WinActivate, ahk_id %IDbCAD% ;переключаемся на bCAD
	controlNeed:="WindowsForms10.BUTTON.app.0.33c0d9d1" ;кнопка выход в окне отчёта конструктора
	ControlSend, %controlNeed%, {Space}, %titlebCADOtchet%
	WinWaitClose, %titlebCADOtchet%, 4
	WinActivate, %ExcelWindowName%
	Sleep, 100
	Send, ^a
	Send, ^c
}
else
{
	MsgBox, bCAD не запущен. Программа завершена
}