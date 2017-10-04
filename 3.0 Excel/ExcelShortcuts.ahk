#SingleInstance force
#IfWinActive, ahk_class XLMAIN
#MaxHotKeysPerInterval 1000

global lcDict := {}

;F1 Test
CapsLock::
/*
	MsgBox, You pressed CapsLock Key
	capsState := GetKeyState("CapsLock", "T")
	MsgBox, %capsState%
	SetCapsLockState, On
	capsState := GetKeyState("CapsLock", "T")
	MsgBox, %capsState%
	*/
return

;Link Constants Only
^#l::

	xl := ComObjActive("Excel.Application")
	xl.ScreenUpdating := False
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.EnableEvents := False

	try {
		SplashTextOn, , 25, , Working
		rwCount := xl.Selection.Rows.Count
		clmCOunt := xl.Selection.Columns.Count
		Loop, %rwCount% {
			rw := A_Index + 1 -1
			Loop, %clmCount% {
				clm := A_Index + 1 -1
				mySelection := xl.ActiveSheet.Range(xl.Selection.Address)
				selectionAddress := xl.Selection.Address
				myFormula := xl.ActiveWorkbook.ActiveSheet.Range(selectionAddress).Cells(rw, clm).Formula
				;msgBox, Formula of Active CEll is %myFormula%
				isTextForm := xl.IsText(mySelection.Cells(rw,clm))
				;msgBox, Is text function returns %isTextForm%
				If ((isTextForm <> -1) and (myFormula <> "")) {
					;msgBox, Not Text
					If (substr(myFormula,1,1) <> "=") {
						newFormula := "=" . "'" . copyRangeTabName . "'!" . StrReplace(copyRange.Cells(rw, clm).Address,"$","")
						;msgBox, New Formula is %newFormula%
						mySelection.Cells(rw, clm).Formula := newFormula
					}
				}
			}
		}
	}

	xl.EnableEvents := True
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	ObjRelease(xl)
	SplashTextOff

return

^!NumpadDiv Up::
	xl := ComObjActive("Excel.Application")
	SendInput, ^!v
	WinWait, Paste Special,, 1
	
	If (ErrorLevel <> 1) {
		SendInput, v
		SendInput, i
		SendInput, {Enter}
	}
	
	ObjRelease(xl)
Return

^!NumpadMult::
	xl := ComObjActive("Excel.Application")
	SendInput, ^!v
	
	If (ErrorLevel <> 1) {
		SendInput, v
		SendInput, m
		SendInput, {Enter}
	}
	
	ObjRelease(xl)
Return

^!NumpadAdd::
	xl := ComObjActive("Excel.Application")
	SendInput, ^!v
	
	If (ErrorLevel <> 1) {
		SendInput, v
		SendInput, s
		SendInput, {Enter}
	}

	ObjRelease(xl)
Return

^!NumpadSub::
	xl := ComObjActive("Excel.Application")
	SendInput, ^!v
	
	If (ErrorLevel <> 1) {
		SendInput, v
		SendInput, d
		SendInput, {Enter}
	}

	ObjRelease(xl)
Return

;Set Print Area
^!+p::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("PrintAreaSetPrintArea")
return


;Total Shortcut
#t::
	xl := ComObjActive("Excel.Application")
	KeyWait, LWin
	Sleep, 10
	SendInput, {Esc}
	Sleep, 10
	SendInput, ="Total " & 
	Sleep, 10
	SendInput ^{Up}
return

;Calculation Automatic
^#c::
	xl := ComObjActive("Excel.Application")
	xl.Calculation := -4105
	ObjRelease(xl)
return

;Mouse Key
^#F8::
	SendInput ^+=
return

;Mouse Key
^#F11::
	SendInput ^x
return

;Mouse Key
^#F9:: ;G17 Button on Mouse
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F9
    ;Msgbox, you pressed control windows f12
    SendInput ^{PgDn}
	ObjRelease(xl)
return

;Mouse Key
^#F7:: ;G17 Button on Mouse
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F7
    ;Msgbox, you pressed control windows f12
    SendInput ^{PgUp}

	ObjRelease(xl)
return

^ESC::
    SendInput {CtrlBreak}

	ObjRelease(xl)
return

Appskey Up::
	SendInput {AppsKey}
return

;Cell Inspector
AppsKey & i::
	global copyRange
	global copyRangeTabName
	fillColor := copyRange.Interior.Color
	fontColor := copyRange.Font.Color
	styleName := copyRange.Style.Name
	MsgBox, Fill Color: %fillColor% `n Font Color: %fontColor% `n Style Name: %styleName% `n Tab Name %copyRangeTabName%
	ObjRelease(xl)
return


;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL ;GENERAL

;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;NAVIGATION/AUDITING;


;Unhide Column a
^#u::

	try {
		xl := ComObjActive("Excel.Application")
		xl.ScreenUpdating := False
		shtCount := xl.Sheets.Count
		Loop, %shtCount% {
			loopIndex := shtCount - A_Index + 1
			shtVisible := xl.Sheets(loopIndex).Visible
			;MsgBox, %loopIndex% %shtCount% %shtVisible%
			If (shtVisible = -1) {
				xl.Sheets(loopIndex).Activate
				xl.Goto("R1C1", True)
				try xl.CommandBars.ExecuteMso("ColumnsUnhide")
			}
		}
		xl.ScreenUpdating := True
	} catch {
		xl.ScreenUpdating := True
	}

	ObjRelease(xl)
return

;Reset Cells to A1 and Hide Column A
^#Home::

	try {
		xl := ComObjActive("Excel.Application")
		xl.ScreenUpdating := False
		shtCount := xl.Sheets.Count
		Loop, %shtCount% {
			loopIndex := shtCount - A_Index + 1
			shtVisible := xl.Sheets(loopIndex).Visible
			;MsgBox, %loopIndex% %shtCount% %shtVisible%
			If (shtVisible = -1) {
				xl.Sheets(loopIndex).Activate
				try {
					zoomValue := xl.ActiveSheet.Range("A2").Value
					xl.ActiveWindow.Zoom := zoomValue
				}
				xl.Goto("R1C1", True)
				;try xl.CommandBars.ExecuteMso("ColumnsHide")
			}
		}
		xl.ScreenUpdating := True
	} catch {
		xl.ScreenUpdating := True
	}

	ObjRelease(xl)
return

;Jump to KPI Tab
AppsKey & k::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("2.1 KPI").Activate
	}

	ObjRelease(xl)
return

;Jump to Volume Tab
AppsKey & v::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("4.1 Volume").Activate
	}

	ObjRelease(xl)
return

;Jump to Summary Tab
AppsKey & s::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("1.1 Summary").Activate
	}

	ObjRelease(xl)
return

;Jump to Debt Tab
AppsKey & d::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("8.1 Debt").Activate
	}

	ObjRelease(xl)
return

;Jump to FCF Tab
AppsKey & f::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("10.1 FCF").Activate
	}

	ObjRelease(xl)
return

;Jump to Equity Tab
AppsKey & e::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("9.1 Equity").Activate
	}

	ObjRelease(xl)
return

;Jump to Revenue Tab
AppsKey & r::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("4.2 Rev").Activate
	}

	ObjRelease(xl)
return

;Jump to Presentations Tab
AppsKey & p::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("Graphs").Activate
	}

	ObjRelease(xl)
return

;Jump to Output Summary
AppsKey & 1::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Range("B5:B1000").SpecialCells(2,2).Areas(1).Cells(1).Select
		
	}

	ObjRelease(xl)
return

;Jump to dNPV
AppsKey & n::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Sheets("11.0 dNPV").Activate
		
	}

	ObjRelease(xl)
return



;Jump to Assumptions
AppsKey & 2::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Range("B5:B1000").SpecialCells(2,2).Areas(2).Cells(1).Select
	}

	ObjRelease(xl)
return


;Jump to Buildup
AppsKey & 3::

	try {
		xl := ComObjActive("Excel.Application")		
		xl.Range("B5:B10000").SpecialCells(2,2).Areas(3).Cells(1).Select
	}

	ObjRelease(xl)
return


;Jump to RPA tab
AppsKey & a::

	try {
		xl := ComObjActive("Excel.Application")
		xl.Worksheets("RPA Monthly").Activate
	}

	ObjRelease(xl)
return

;Dependents/Precedents
^[::

	xl := ComObjActive("Excel.Application")
	try global sourceCellPrecedent := xl.Selection
	SendInput ^[
	
	ObjRelease(xl)
return

^]::

	xl := ComObjActive("Excel.Application")
	try global sourceCellDependent := xl.Selection
	SendInput ^]

	ObjRelease(xl)
return

![::

	xl := ComObjActive("Excel.Application")
	global sourceCellPrecedent
	try {
		sourceCellPrecedent.Parent.Activate
		sourceCellPrecedent.select
	}
	ObjRelease(xl)
return


!]::
	xl := ComObjActive("Excel.Application")
	global sourceCellDependent
	try {
		sourceCellDependent.Parent.Activate
		sourceCellDependent.Select
	}
	ObjRelease(xl)
return

;TraceDependents
^!]::
	xl := ComObjActive("Excel.Application")
	xl.Commandbars.ExecuteMso("TraceDependents")
	ObjRelease(xl)
return

;TracePrecedents
^![::
	xl := ComObjActive("Excel.Application")
	xl.Commandbars.ExecuteMso("TracePrecedents")
	ObjRelease(xl)
return


;Remove All Arrows
^!\::
	xl := ComObjActive("Excel.Application")
	KeyWait, Control
	KeyWait, Alt
	SendInput {Alt}
	SendInput, m
	SendInput, a
	SendInput, a
	ObjRelease(xl)
return




;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__;Remap Keys__


;Execute Array Formulas
^+Enter::

	Send {F2}
	Send ^+{Enter}

	ObjRelease(xl)


return

;Move Focus Down One Row
NumpadEnter::
+Enter::

	SendInput {Shift up}
	SendInput {Enter}
	SendInput {Down}

	ObjRelease(xl)
return

;ClearAll
>^Del::
>^BS::

	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("ClearAll")

	ObjRelease(xl)
return

;Remap Delete
+BS::

	SendInput {Del}

	ObjRelease(xl)
return



;Application Settings; ApplicationSettings; ApplicationSettings; ApplicationSettings; ApplicationSettings; ApplicationSettings; ApplicationSettings; ApplicationSettings; ApplicationSettings; ApplicationSettings;


;Toggle Gridlines
^!g::
	SendInput {LCtrl Up}
	SendInput {LCtrl Down}
	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("ViewSheetGridlines")

	ObjRelease(xl)
return



;Print Preview
^#Space:: ;Alt-Space
	xl := ComObjActive("Excel.Application")    
		If (xl.ActiveWindow.View = 2) {	
			xl.ActiveWindow.View := 1
		} Else {
			xl.ActiveWindow.View := 2
		}

	ObjRelease(xl)
return


;Toggle Focus after Enter 
>#Enter::	
	xl := ComObjActive("Excel.Application")
	if (xl.MoveAfter = False) {
		try {
			xl.MoveAfter := True
		}
	}
	else {
		try {
			xl.MoveAfter := False
		}
	}

	ObjRelease(xl)
return


;FreezePanesShortcut
!f::!6

	ObjRelease(xl)
return


;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__;Functions__


;SUMIF Macro
^#+s::

	xl := ComObjActive("Excel.Application")
	xl.Run("PERSONAL.XLSB!MODEL_Sumifs")

	ObjRelease(xl)
return


;Update Change Tracker
#u::

	xl := ComObjActive("Excel.Application")

	thisSheet := ""
	notesCell := ""
	filenameCell := ""
	bookName := xl.ActiveWorkbook.Name

	noteField := xl.InputBox(prompt:="Enter Notes for this update")

	try {

		xl.EnableEvents := False
		oldCalcState := xl.Calculation
		xl.Calculation := -4135
		xl.ScreenUpdating := False
	}

	sheetCount := xl.ActiveWorkbook.Worksheets.Count

	Loop, %sheetCount% {
		shtCheck := xl.ActiveWorkbook.Worksheets(A_Index).CodeName
		
		If (shtCheck = "Sheet11") {
			thisSheet := xl.ActiveWorkbook.Worksheets(A_Index)
		}
	}

	cellCount := thisSheet.Range("A5:AB5").Cells.Count
	Loop, %cellCount% {

		cll := thisSheet.Range("A5:AB5").Cells(A_Index)
		cllFormula := thisSheet.Range("A5:AB5").Cells(A_Index).Formula
		
		If (cllFormula = "Notes") {
			notesCell := cll
	}
		
		If (cllFormula = "Filename") {
			filenameCell := cll
		}
	}


	thisSheet.Rows(7).Select
	xl.commandBars.ExecuteMso("SheetRowsInsert")

	thisSheet.Rows(6).Select
	xl.commandBars.ExecuteMso("Copy")
	thisSheet.Rows(7).EntireRow.Select
	xl.CommandBars.ExecuteMso("PasteValues")
	;thisSheet.Rows(6).EntireRow.Select
	;xl.CommandBars.ExecuteMso("PasteFormats")

	helperCell := xl.Workbooks("PERSONAL.XLSB").Sheets(1).Range("A1")
	helperCell.Formula := noteField
	helperCell.copy
	notesCell.Offset(2, 0).Select
	xl.CommandBars.ExecuteMso("PasteValues")

	helperCell.Formula := bookName
	helperCell.copy
	filenameCell.Offset(2, 0).Select
	xl.CommandBars.ExecuteMso("PasteValues")

	xl.EnableEvents := True
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	try{
	} catch {

	try {	
	xl.EnableEvents := True
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	} 
	}

	ObjRelease(xl)
return

;Append Text
#3::

	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False

	appendText := xl.InputBox("Text to append")
	myRange := xl.Selection
	cellCount := myRange.Cells.Count
	
	Loop, %cellCount% {
		Loop, 10 {
			try {
				myRange.Cells(A_Index).Select
				SendInput {F2}
				SendInput %appendText%
				SendInput {Enter}
				SendInput {Enter}
				Break
			} catch {
				sleep, 10
			}
		}
	}
		
	Loop, 4 {
	try {
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	} catch {
	SendInput {Enter}
	sleep, 10
	}
	}


	;try {
		xl := ComObjActive("Excel.Application")
		oldCalcState := xl.Calculation
		xl.Calculation := -4135
		xl.ScreenUpdating := False
		xl.EnableEvents := False

		;appendText := xl.InputBox("Text to append (including space)")
		myselection := xl.Selection
		selectionAddress := myselection.Address(True, True)
		myselection.copy

		rngAddress := "C3:" . xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range("C3").Offset(myselection.Rows.Count-1,myselection.Columns.Count-1).Address
		Range1 := xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range(rngAddress)
		Range1.PasteSpecial(-4122)
		xl.Run("PERSONAL.XLSB!FUNCTIONS_AppendFormula")
		Range1.Copy
		Range1.Clear
		myselection.Select
		xl.CommandBars.ExecuteMso("PasteValues")
		
		
		xl.Calculation := oldCalcState
		xl.ScreenUpdating := True
		xl.EnableEvents := True

	try {

		} catch {
		Loop, 10 {
			try {
				xl.Calculation := oldCalcState
				xl.ScreenUpdating := True
				xl.EnableEvents := True
				Break
				} Catch {
				sleep 10
			}
		}
	}

	ObjRelease(xl)
return

;Secondary Dependent
^+[::

	xl := ComObjActive("Excel.Application")
	
	;xl.EnableEvents := False
	;oldCalcState := xl.Calculation
	;xl.Calculation := -4135
	;xl.ScreenUpdating := False

	;try {
		mySelectionAddress := xl.Selection.Address
		areacount := xl.ActiveSheet.Range(mySelectionAddress).Precedents.Address
		
	;} catch {
		MsgBox, %areacount%
	;}

		ObjRelease(xl)
return

;Multiply Ranges 
#+= Up::

	xl := ComObjActive("Excel.Application")
	;SendInput {# Up}
	

	newformula := xl.Selection.Address(False,False)


	newformula := RTrim(newformula, ",")
	;MsgBox, %newformula%

	newformula := "=" . StrReplace(newformula,",","*")
	;MsgBox, %newformula%

	xl.Selection.Cells(xl.Selection.Cells.Count).Select 
	SendRaw %newformula%


	ObjRelease(xl)
return

;Sum Independent Ranges 
#= Up::

	xl := ComObjActive("Excel.Application")
	;SendInput {# Up}
	

	newformula := xl.Selection.Address(False,False)
	replacestring := xl.Selection.Areas(xl.Selection.Areas.Count).Address(False,False)
	;MsgBox, %newformula%

	newformula := StrReplace(newformula,replacestring)
	;MsgBox, %newformula%

	newformula := RTrim(newformula, ",")
	;MsgBox, %newformula%

	newformula := "=" . StrReplace(newformula,",","+")
	;MsgBox, %newformula%

	xl.Selection.Areas(xl.Selection.Areas.Count).Select 
	SendRaw %newformula%


	ObjRelease(xl)
return

;Append Text
;!Right::

	xl := ComObjActive("Excel.Application")
	appendText := xl.InputBox("Enter String to Add to Formula")
	Send {F2}
	SendRaw %appendText%
	Send {Enter}
	Send {Enter}

	ObjRelease(xl)
return



;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes;Shapes

;Align Shapes Top
AppsKey & t::
	xl := ComObjActive("Excel.Application")
	shapeCount := xl.Selection.ShapeRange.Count
	xl.Selection.ShapeRange.Top := xl.Selection.ShapeRange(shapeCount).Top
	ObjRelease(xl)
return

;Align Shapes Left
AppsKey & l::
	xl := ComObjActive("Excel.Application")
	shapeCount := xl.Selection.ShapeRange.Count
	xl.Selection.ShapeRange.Left := xl.Selection.ShapeRange(shapeCount).Left
	ObjRelease(xl)
return

;Match Shape Widths
Appskey & w::
	xl := ComObjActive("Excel.Application")
	shapeCount := xl.Selection.ShapeRange.Count
	xl.Selection.ShapeRange.Width := xl.Selection.ShapeRange(shapecount).Width
	
	ObjRelease(xl)
return

;Match Shape Heights
Appskey & h::
	xl := ComObjActive("Excel.Application")
	shapeCount := xl.Selection.ShapeRange.Count
	xl.Selection.ShapeRange.Height := xl.Selection.ShapeRange(shapecount).Height
	
	ObjRelease(xl)
return

Appskey & u::
	xl := ComObjActive("Excel.Application")
	shapeCount := xl.Selection.ShapeRange.Count
	xl.Selection.ShapeRange(1).Height := xl.Selection.ShapeRange(shapecount).Height
	
	ObjRelease(xl)
return


;AlignDistributeVerticallyClassic

;ObjectsAlignLeft
;^!l::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectsAlignLeft")

	ObjRelease(xl)
return

;ObjectsAlignCenterHorizontal
^!c::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectsAlignCenterHorizontal")

	ObjRelease(xl)
return

;ObjectsAlignRight
;^!r::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectsAlignRight")

	ObjRelease(xl)
return

;ObjectsAlignTop
;^!t::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectsAlignTop")

	ObjRelease(xl)
return

;ObjectsAlignMiddleVertical
^!m::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectsAlignMiddleVertical")

	ObjRelease(xl)
return

;ObjectsAlignBottom
;^!b::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectsAlignBottom")

	ObjRelease(xl)
return

;AlignDistributeHorizontallyClassic
^+h::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("AlignDistributeHorizontallyClassic")

	ObjRelease(xl)
return




;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS;ROWS COLUMNS


; New Row Insert
+#Space::
	global lcDict

	xl := ComObjActive("Excel.Application")
	;try {	
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	xl.ActiveSheet.DisplayPageBreaks := False
	xl.CommandBars.ExecuteMso("SheetRowsInsert")

	myRange := xl.Selection
	tabName := xl.ActiveSheet.Name
	checkRow := xl.Selection.Row -1
	lastRow := checkRow + xl.Selection.Rows.Count
	lastColumn := 0

	;msgbox, check row %checkRow% last row %lastRow%
	lastColumnBefore := lcDict[tabName]
	;msgbox, %lastColumnBefore%

	if (lcDict[(tabName)] > 1) {
		lastColumn := lcDict[tabName]
		;msgBox, %last%
		GoTo Success
	} else {
		;xl.ActiveSheet.Cells(checkRow, 1).Activate
		debugvar := xl.ActiveSheet.Cells(checkRow, 48).Value
		;msgBox, %debugVar%
		Loop, 100 {
			if (xl.ActiveSheet.Cells(checkRow, A_Index).Value = "|") {
				lcDict[tabName] := A_Index
				lastColumn := lcDict[tabName]
				;msgBox, in the loop last COlumn %lastColumn%
				GoTo Success
			} else {
				xl.ActiveCell.Offset(0,1).Select
				
			}
		}
		
	}
	;msgBox, Go to ending
	GoTo Ending

	Success:
	checkValue := xl.ActiveSheet.Cells(checkRow, lcDict[tabName]).Value
	;msgbox, %checkFormula%

	if (checkValue = "|") {
		xl.activesheet.Range(xl.ActiveSheet.Cells(checkRow, lcDict[tabName]),xl.ActiveSheet.Cells(lastRow, lcDict[tabName])).select
		xl.Commandbars.ExecuteMso("FillDown")
	}
	
	
	Ending:
	myRange.select
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.enableevents := True
	
	
	ObjRelease(xl)

	/*
	xl := ComObjActive("Excel.Application")
	try {	
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	xl.ActiveSheet.DisplayPageBreaks := False
	xl.CommandBars.ExecuteMso("SheetRowsInsert")
	
	myRange := xl.Selection
	rwNumber := xl.Selection.Row + xl.Selection.Rows.Count
	firstRowNumber := xl.Selection.Row
	checkAddress := "AV" . (firstRowNumber -1)
	;msgbox %checkAddress%
	
	If (xl.activesheet.Range(checkAddress).formula = "|") {
		;msgbox triggered
		selectAddress := "AV" . (rwNumber) . ":" . checkAddress
		;msgbox %selectaddress%
		xl.activesheet.Range(selectAddress).select
		xl.Commandbars.ExecuteMso("FillDown")
	}
	
	checkAddress := "AN" . (firstRowNumber -1)
	;msgbox %checkAddress%
	
	If (xl.activesheet.Range(checkAddress).formula = "|") {
		;msgbox triggered
		selectAddress := "AN" . (rwNumber) . ":" . checkAddress
		;msgbox %selectaddress%
		xl.activesheet.Range(selectAddress).select
		xl.Commandbars.ExecuteMso("FillDown")
	}
	
	myrange.select
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.enableevents := True
	
	ObjRelease(xl)
	*/
return

; New Row Delete
+!Space::
	xl := ComObjActive("Excel.Application")
	try {	
		oldCalcState := xl.Calculation
		xl.Calculation := -4135
		xl.ScreenUpdating := False
		;xl.EnableEvents := False
		xl.CommandBars.ExecuteMso("SheetRowsDelete")
	}
	
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	;xl.enableevents := True
	ObjRelease(xl)
return

; Group Columns
^!#Left::

	xl := ComObjActive("Excel.Application")
	xl.ScreenUpdating := False
	try {
		mySelection := xl.Selection
		xl.Selection.EntireColumn.Select
		xl.CommandBars.ExecuteMso("OutlineGroup")
		xl.CommandBars.ExecuteMso("OutlineHideDetail")
		mySelection.Select
	}
	xl.ScreenUpdating := True  
	
	ObjRelease(xl)
return

;Ungroup Columns
^!#Right::

	xl := ComObjActive("Excel.Application")
	xl.ScreenUpdating := False
	try {
		mySelection := xl.Selection
		xl.Selection.EntireColumn.Select
		xl.CommandBars.ExecuteMso("OutlineUngroup")
		mySelection.Select
	}
	xl.ScreenUpdating := True  

	ObjRelease(xl)
return

;Group Rows
^!#Up::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	xl.ActiveSheet.DisplayPageBreaks := False
	
	try {
		selectionAddress := xl.Selection.Address
		
		Loop, Parse, selectionAddress, `, 
		{
		myRange := xl.ActiveSheet.Range(A_LoopField)
			Loop, 10 {
				try {
					myRange.EntireRow.Select
					Break
					} catch {
					sleep, 10
				}
			}
			
		xl.CommandBars.ExecuteMso("OutlineGroup")
		xl.CommandBars.ExecuteMso("OutlineHideDetail")
		}
		mySelection.Select
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.enableevents := True  
	
	ObjRelease(xl)
return


;Ungroup Rows
^!#Down::

	xl := ComObjActive("Excel.Application")
	xl.ScreenUpdating := False
	try {
		mySelection := xl.Selection
		xl.Selection.EntireRow.Select
		xl.CommandBars.ExecuteMso("OutlineUngroup")
		mySelection.Select
	}
	xl.ScreenUpdating := True 
	
	ObjRelease(xl)
return

;Rows Insert
+=::

	if (GetKeyState("Space", "P")) {
		xl := ComObjActive("Excel.Application")
		try {
			oldCalcState := xl.Calculation
			xl.Calculation := -4135
			;xl.ScreenUpdating := False
			;xl.EnableEvents := False
			xl.CommandBars.ExecuteMso("SheetRowsInsert")
			xl.Calculation := oldCalcState
			;xl.ScreenUpdating := True
			;xl.EnableEvents := True
		}
	}
	else{
		Send {+}
	}

	ObjRelease(xl)
return

;Columns Insert
^=::

	if (GetKeyState("Space", "P")) {
		xl := ComObjActive("Excel.Application")
		try {
			oldCalcState := xl.Calculation
			xl.Calculation := -4135
			xl.ScreenUpdating := False
			xl.EnableEvents := False
			xl.CommandBars.ExecuteMso("SheetColumnsInsert")
			xl.Calculation := oldCalcState
		}
		xl.ScreenUpdating := True
		xl.EnableEvents := True
		xl.Calculation := oldCalcState
	}
	else {
		Send {^=}
	}
	
	ObjRelease(xl)
return


;Rows Delete
+-::

	if (GetKeyState("Space", "P")) {
		xl := ComObjActive("Excel.Application")
		try {
			oldCalcState := xl.Calculation
			xl.Calculation := -4135
			xl.CommandBars.ExecuteMso("SheetRowsDelete")
			xl.Calculation := oldCalcState
		}
	}
	else {
		Send {_}
	}

	ObjRelease(xl)
return

;Columns Delete
^-::

	xl := ComObjActive("Excel.Application")
	if (GetKeyState("Space", "P")) {
		try { 
			oldCalcState := xl.Calculation
			xl.Calculation := -4135
			xl.ScreenUpdating := False
			xl.EnableEvents := False
			xl.CommandBars.ExecuteMso("SheetColumnsDelete")
		}
	}
	else {
		try {
			oldCalcState := xl.Calculation
			xl.Calculation := -4135
			xl.ScreenUpdating := False
			xl.EnableEvents := False
			xl.CommandBars.ExecuteMso("CellsDelete")
		}
	}
		xl.ScreenUpdating := True
		xl.EnableEvents := True
		xl.Calculation := oldCalcState

	ObjRelease(xl)
return

;Columns Unhide
^)::

	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ColumnsUnhide")

	ObjRelease(xl)
return

;Groups Collapse
^!9::

	xl := ComObjActive("Excel.Application")
	try {
		xl.ScreenUpdating := False
		oldCalcState := xl.Calculation
		xl.Calculation := 4135
		xl.EnableEvents := False
		xl.ActiveSheet.DisplayPageBreaks := False
		xl.CommandBars.ExecuteMso("OutlineHideDetail")
		}
	;MsgBox %oldCalcState%
	xl.ScreenUpdating := True
	xl.Calculation := oldCalcState
	xl.EnableEvents := True
	xl.ActiveSheet.DisplayPageBreaks := True

	ObjRelease(xl)
return




;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL  ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL  ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL  ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL ;PASTESPECIAL 


;Copy
~^c::


	xl := ComObjActive("Excel.Application")
	try {
		global copyRange := xl.Selection
		global copyRangeAddress := xl.Selection.Address
		global copyRangeAddressExternal := xl.Selection.Address(False,False,1,True)
		global copyRangeTabName :=xl.Selection.Parent.Name
	}
	

	ObjRelease(xl)
return

;Recopy
;^#c::

	xl := ComObjActive("Excel.Application")
	Recopy:

	global copyRange

	xl.ScreenUpdating := False
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	myRange := xl.Selection
	copyRange.Select
	xl.Commandbars.ExecuteMso("Copy")
	myRange.Select
	xl.ScreenUpdating := True
	oldCalcState := xl.Calculation

	ObjRelease(xl)
return

;Paste Address
^!a up::

	;try {
		xl := ComObjActive("Excel.Application")
		global copyRangeAddress
		newClip := StrReplace(copyRangeAddress,"$","")
		KeyWait LAlt
		KeyWait LControl
		KeyWait a
		
		SendRaw %newClip%
	;}

	ObjRelease(xl)
return

;Paste Tab Name
^!b::

	try {
		xl := ComObjActive("Excel.Application")
		global copyRangeTabName
		newClip := copyRangeTabName
		
		KeyWait LAlt
		KeyWait LControl
		KeyWait b
		
		SendRaw '%newClip%'!
	}

	ObjRelease(xl)
return


;Remove $
#4::

	try {
		SendInput ^c
		clipText := %Clipboard%
		clipText := StrReplace(clipText,"$","")
		Clipboard := clipText
		SendInput ^v
	}

	ObjRelease(xl)
return

;Screen Updating is True
^#s::
	xl := ComObjActive("Excel.Application")
	xl.ScreenUpdating := True

	ObjRelease(xl)
return

;Paste Duplicate
^!d::
	xl := ComObjActive("Excel.Application")
	SendInput {Ctrl up}
	SendInput {Ctrl down}
	;GoTo Recopy
	SendInput ^!d

	ObjRelease(xl)
return

;Paste Exact
^!e::
	xl := ComObjActive("Excel.Application")
	SendInput {Ctrl up}
	SendInput {Ctrl down}
	;GoSub Recopy
	SendInput ^!e

	ObjRelease(xl)
return


;Paste Links
^!l::
	xl := ComObjActive("Excel.Application")
	SendInput {Ctrl up}
	SendInput {Ctrl down}
	copyrange.copy
	SendInput ^!l
	ObjRelease(xl)
return

;Drake Special Paste Links
^!+l up::

	KeyWait LWin
	xl := ComObjActive("Excel.Application")
	;try {
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False

	global copyRange
	thissRange := copyRange
	myyRange := xl.Selection
	copyAddress := copyRange.Address
	;MsgBox, %copyAddress%
	Loop, parse, copyAddress, `, 
	{
	areaIndex := A_Index
	formulaText := A_LoopField
		Loop, 100
		{
		Try {
			myyRange.Areas(areaIndex).Select
			Break
			} Catch {
			Sleep 100
			}
		}
			;msgBox, %A_LoopField%
		Send {=}
		SendRaw %FormulaText%
		Send {Enter}
		Send {Enter}
	}
	Loop, 100
	{ 
		try {
			myyRange.Select
			xl.Calculation := oldCalcState
			xl.ScreenUpdating := True
			xl.EnableEvents := True
			Break
		} Catch {
			sleep 10
		}
	}

	ObjRelease(xl)
return

;PasteValues
^!v::
	SendInput {Ctrl up}
	SendInput {Ctrl down}
	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("PasteValues")
	catch {
	global copyrange
	copyrange.copy
	xl.CommandBars.ExecuteMso("PasteValues")
	}

ObjRelease(xl)
return

;NumberFormatting
^!n::

	global copyRange
	SendInput {Ctrl up}
	SendInput {Ctrl down}
	xl := ComObjActive("Excel.Application")
	try {
	myselection := xl.Selection
	myselection.copy
	rngAddress := "A1:" . xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range("A1").Offset(myselection.Rows.Count-1,myselection.Columns.Count-1).Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range(rngAddress)
	style1.PasteSpecial(-4123) ;PasteFormulas
	copyRange.copy
	xl.CommandBars.ExecuteMso("PasteValuesAndNumberFormatting")
	style1.copy
	xl.CommandBars.ExecuteMso("PasteFormulas")
	}

ObjRelease(xl)
return

;PasteFormatting
^!r::

	global copyRange
	SendInput {Ctrl up}
	SendInput {Ctrl down}
	xl := ComObjActive("Excel.Application")
	try {
	xl.CommandBars.ExecuteMso("PasteFormatting")


	}catch {
	global copyRange
	copyRange.copy
	xl.CommandBars.ExecuteMso("PasteFormatting")
	}

ObjRelease(xl)
return

;paste special dialog
PasteSpecialDialog:
^+v::

	SendInput {Ctrl up}
	SendInput {Ctrl down}
	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("PasteSpecialDialog")

ObjRelease(xl)
return 


;Paste Formulas
^!f::

	SendInput {Ctrl up}
	SendInput {Ctrl down}
	xl := ComObjActive("Excel.Application")
	try {
		xl.CommandBars.ExecuteMso("PasteFormulas")
	} catch {
		global copyrange
		copyrange.copy
		xl.CommandBars.ExecuteMso("PasteFormulas")
	}


ObjRelease(xl)
return



;HORIZONTAL SCROLLING ;HORIZONTAL SCROLLING ;HORIZONTAL SCROLLING ;HORIZONTAL SCROLLING ;HORIZONTAL SCROLLING ;HORIZONTAL SCROLLING 


WheelRight::

	try ComObjActive("Excel.Application").ActiveWindow.SmallScroll(0,0,3)  

	ObjRelease(xl)
return

WheelLeft::

	try ComObjActive("Excel.Application").ActiveWindow.SmallScroll(0,0,0,3)  

	ObjRelease(xl)
return




;xlPasteAll
;-4104
;Everything will be pasted.

;xlPasteAllExceptBorders
;7
;Everything except borders will be pasted.

;xlPasteAllMergingConditionalFormats
;14
;Everything will be pasted and conditional formats will be merged.

;xlPasteAllUsingSourceTheme
;13
;Everything will be pasted using the source theme.

;xlPasteColumnWidths
;8
;Copied column width is pasted.

;xlPasteComments
;-4144
;Comments are pasted.

;xlPasteFormats
;-4122
;Copied source format is pasted.

;xlPasteFormulas
;-4123
;Formulas are pasted.

;xlPasteFormulasAndNumberFormats
;11
;Formulas and Number formats are pasted.

;xlPasteValidation
;6
;Validations are pasted.

;xlPasteValues
;-4163
;Values are pasted.

;xlPasteValuesAndNumberFormats
;12
;Values and Number formats are pasted

/*
xlPasteSpecialOperationAdd	2	Copied data will be added to the value in the destination cell.
xlPasteSpecialOperationDivide	5	Copied data will divide the value in the destination cell.
xlPasteSpecialOperationMultiply	4	Copied data will multiply the value in the destination cell.
xlPasteSpecialOperationNone	-4142	No calculation will be done in the paste operation.
xlPasteSpecialOperationSubtract	3	Copied data will be subtracted from the value in the destination cell.
*/
