#SingleInstance force
#IfWinActive, ahk_class XLMAIN
#MaxHotKeysPerInterval 1000
#InputLevel 0

;OTHERS

;Macro Test
AppsKey & Space::
	xl := ComObjActive("Excel.Application")
	colorInteger := xl.RGB(255, 255, 200)
	MsgBox, %colorIndex%

	;xlH := ComObj(9, 10603208)
	;xlH.Visible := True

	;MsgBox, triggered
	ObjRelease(xl)
return


;NumberFormat
!#1::
	xl := ComObjActive("Excel.Application")


	try {
	nfString := xl.Selection.NumberFormat

	;msgBox numberFormatString: %nfString%
	nfArray := StrSplit(nfString,";")

		If (nfArray.MaxIndex() = 4) {

			zeroFormat := nfArray[3]
			;msgbox, zeroString: %zeroFormat%
			
			If (Instr(zeroFormat, """-""") > 0) {
			
				zeroFormat := StrReplace(zeroFormat, """-""", "0")
				;MsgBox, New ZeroString found a dash %zeroFormat%
			} else if (Instr(zeroFormat, "0") > 0) {
			
				zeroFormat := StrReplace(zeroFormat, "0", """-""")
				;MsgBox, New ZeroString found a 0 %zeroFormat%
			}

			nfArray[3] := zeroFormat
			
			Loop, % nfArray._MaxIndex() {
				
				newNFString := newNFString . nfArray[A_Index] . ";"
				
			}
			
			newNFString := RTrim(newNFString, ";")
			
			;MsgBox %newNFString%
			
			SendInput ^1
			sleep 10
			SendInput n
			sleep 10
			SendInput !t
			sleep 100
			SendInput {Del}
			sleep 10
			SendRaw %newNFString%
			SendInput {Enter}
			
			
			;xl.Selection.NumberFormat := newNFString
		}
	}
	ObjRelease(xl)
return


;MillionsFormat
#+m::
	xl := ComObjActive("Excel.Application")

	;use non expression type formatting (does not need to be enclosed in double quotes)
	newFormat = #,##0,"m";(#,##0,"m");0"m";@
	xl.Selection.NumberFormat := newFormat
	ObjRelease(xl)
return

;NumberFormat
!#2::
	xl := ComObjActive("Excel.Application")

	findString := xl.InputBox("String to Find")
	replaceString := xl.InputBox("String to Replace")
	try {
	nfString := xl.Selection.NumberFormat

	;msgBox numberFormatString: %nfString%
	nfArray := StrSplit(nfString,";")

		
	Loop, % nfArray._MaxIndex() - 1 {
		
			newString := RTrim(nfArray[A_Index], findString)
			
			bStringFound := StrLen(nfArray[A_Index]) - StrLen(newString)
			
			If (bStringFound > 0) {
			
				nfArray[A_Index] := newString . replaceString
			}
			
			;MsgBox, New ZeroString found a dash %zeroFormat%
		}

		
		Loop, % nfArray._MaxIndex() {
			
			newNFString := newNFString . nfArray[A_Index] . ";"
			
		}
		
		newNFString := RTrim(newNFString, ";")
		
		;MsgBox %newNFString%
		
		SendInput ^1
		sleep 10
		SendInput n
		sleep 10
		SendInput !t
		sleep 100
		SendInput {Del}
		sleep 10
		SendRaw %newNFString%
		SendInput {Enter}
	}
	ObjRelease(xl)
return


;NumberFormat
!#3::
	xl := ComObjActive("Excel.Application")
	replaceString := xl.InputBox("String to Add")
	try {


		nfString := xl.Selection.NumberFormat

		;msgBox numberFormatString: %nfString%
		nfArray := StrSplit(nfString,";")

			
		Loop, % nfArray._MaxIndex() - 1 {

			nfArray[A_Index] := nfArray[A_Index] . replaceString
			
				;MsgBox, New ZeroString found a dash %zeroFormat%
			}

			
			Loop, % nfArray._MaxIndex() {
				
				newNFString := newNFString . nfArray[A_Index] . ";"
				
			}
			
			newNFString := RTrim(newNFString, ";")
			
			;MsgBox %newNFString%
			
			SendInput ^1
			sleep 10
			SendInput n
			sleep 10
			SendInput !t
			sleep 1000
			SendInput {Del}
			sleep 10
			SendRaw %newNFString%
			SendInput {Enter}
			
		Sleep 10

		}
	ObjRelease(xl)
return


;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES;SHAPES

;Set Global Height
AppsKey & h::

	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterHeight.txt"
	If (keyState = 0) {
		FileRead, newValue, %fileLocation%
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange
			} catch {
				myshape := xl.Selection
			}
			
		myShape.Height := newValue
	ObjRelease(xl)
	return
	} else If (keyState = 1) {
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange(1)
			} catch {
				myshape := xl.Selection
			}
			
		shpText := myShape.Height
		
		FileDelete , %fileLocation%
		FileAppend, %shpTExt%, %fileLocation%
	}	
	ObjRelease(xl)
return

;Global Width
AppsKey & w::
	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterWidth.txt"
	If (GetKeyState("LWin", "P") = 0) {
		FileRead, newValue, %fileLocation%
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange
			} catch {
				myshape := xl.Selection
			}
			
		myShape.Width := NewValue
	} else If (keyState = 1) {
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange(1)
			} catch {
				myshape := xl.Selection
			}
			
		shpText := myShape.Width
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
	ObjRelease(xl)
return

;Apply Global Top
AppsKey & t::
	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterTop.txt"
	If (KeyState = 0) {
		FileRead, newValue, %fileLocation%
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange
			} catch {
				myshape := xl.Selection
			}
			
		myShape.Top := NewValue
	} else If (keyState = 1) {
		xl := ComObjActive("Excel.Application")
		try {
			myShape := xl.ShapeRange(1)
		} catch {
			myshape := xl.Selection
		}		
		shpText := myShape.Top
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
	ObjRelease(xl)
return

;Apply Global Left
AppsKey & l::
	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterLeft.txt"
	If (keyState = 0) {
		FileRead, newValue, %fileLocation%
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange
			} catch {
				myshape := xl.Selection
			}
			
		myShape.Left := NewValue
	} else If (keyState = 1) {
		xl := ComObjActive("Excel.Application")
		try {
			myShape := xl.ShapeRange(1)
		} catch {
			myshape := xl.Selection
		}
		
		shpText := myShape.Left
		
		FileDelete , %fileLocation%
		FileAppend, %shpTExt%, %fileLocation%
	}
	ObjRelease(xl)
return

;Global Aspect Ratio
AppsKey & p::
	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterAspectRatio.txt"
	If (keyState = 0) {
		FileRead, newValue, %fileLocation%
		msgBox, %newValue%
		xl := ComObjActive("Excel.Application")
			try {
				myShape := xl.ShapeRange(1)
			} catch {
				myshape := xl.Selection
			}
			
		myShape.Height := myShape.Width / newValue
	} else If (keyState = 1) {
		xl := ComObjActive("Excel.Application")
		try {
			myShape := xl.ShapeRange(1)
		} catch {
			myshape := xl.Selection
		}
		
		;shpText := myShape.Width / myShape.Height
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
	ObjRelease(xl)
return

;Lock Aspect Ratio
#p::
	
	xl := ComObjActive("Excel.Application")
	try {
		myShape := xl.ShapeRange
	} catch {
		myshape := xl.Selection.Parent.Parent.ShapeRange
	}
		
	myShape.LockAspectRatio := -1
	ObjRelease(xl)
return

;Unlock Aspect Ratio
#!p::
	xl := ComObjActive("Excel.Application")
	try {
		myShape := xl.ShapeRange
	} catch {
		myshape := xl.Selection.Parent.Parent.ShapeRange
	}
		
	myShape.LockAspectRatio := 0
	ObjRelease(xl)
return


;Object Fill
AppsKey & f::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectFillMoreColorsDialog")
	ObjRelease(xl)
return

;Object Outline
AppsKey & o::
	sleep, 100
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ObjectOutlineMoreColorsDialog")
	ObjRelease(xl)
return


;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING;FORMATTING
{

;Fill
#f::
	ObjRelease(xl)
return

;Toggle Blue
^`;::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	colorState := xl.Selection.Font.Color
	if (colorState <> 0) {
	myselection := xl.Selection
	myselection.copy
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Worksheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Font.Color := 0
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	} else {
	myselection := xl.Selection
	myselection.copy
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Worksheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Font.Color := 167116800
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	}
	;style1.clear
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;First Column Fill
^#5::
	global xlHelper
	global xlHelperBook
	xl := ComObjActive("Excel.Application")

	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False




	myselection := xl.Selection
	myselection.copy
	rngAddress := "A1:" . xlHelperBook.Sheets(1).Range("A1").Offset(myselection.Rows.Count,myselection.Columns.Count).Address
	msgbox, %rngAddress%

	style1 := xlHelperSheet.Range(rngAddress)
	style1.PasteSpecial

	style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("#g").Interior.Color
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	style1.Clear


	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True

	ObjRelease(xl)
return

;First Column Fill
^!#5::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	colorState := xl.Selection.Interior.Color
	if (colorState <> xl.Workbooks("PERSONAL.XLSB").Styles("^#5").Interior.Color) {
	myselection := xl.Selection
	myselection.copy
	rngAddress := myselection.Address
	style1 := xlHelperBook.Sheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("^#5").Interior.Color
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	style1.Clear
	} else {
	SendInput {RAlt}
	SendInput hhn
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

; Fill Red
#r::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	colorState := xl.Selection.Interior.Color
	if (colorState <> 10066431) {
	myselection := xl.Selection
	myselection.copy
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Interior.Pattern := 1
	style1.Interior.Color := 10066431
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	} else {
	SendInput {RAlt}
	SendInput hhn
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;Green Fill
#g::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	colorState := xl.Selection.Interior.Color
	if (colorState <> xl.Workbooks("PERSONAL.XLSB").Styles("#g").Interior.Color) {
	myselection := xl.Selection
	myselection.copy
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Sheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("#g").Interior.Color
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	} else {
	SendInput {Alt}
	SendInput 09
	xl.CommandBars.FindControl(,1453).Execute
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;Green Outline
!#g::
	xl := ComObjActive("Excel.Application")

	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False


	myselection := xl.Selection
	myselection.copy

	rngAddress := myselection.Address
	style1 := xlHelperBook.Sheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	;newColorIndex := xl.Workbooks("PERSONAL.XLSB").Styles("GreenOutline").Borders(1).Color
	style1.BorderAround(1, 2, , 1562486)
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")

	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;Yellow Fill
#y::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	colorState := xl.Selection.Interior.Color
	if (colorState <> xl.Workbooks("PERSONAL.XLSB").Styles("#y").Interior.Color) {
	myselection := xl.Selection
	myselection.copy
	rngAddress := myselection.Address
	style1 := xlHelperBook.Sheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("#y").Interior.Color
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	} else {
	SendInput {RAlt}
	SendInput hhn
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;Font Size Increase
^+.::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("FontSizeIncrease")
	ObjRelease(xl)
return

;Font Size Decrease
^+,::
	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("FontSizeDecrease")
	ObjRelease(xl)
return


;Increase Decimals
^.::
	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("DecimalsIncrease")
	ObjRelease(xl)
return

;Decrease Decimals
^,::
	xl := ComObjActive("Excel.Application")
	try xl.CommandBars.ExecuteMso("DecimalsDecrease")
	ObjRelease(xl)
return


;Wrap Text
^+w::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("WrapText")
	ObjRelease(xl)
return

; Clear Formatting
!Backspace::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("ClearFormats")
	ObjRelease(xl)
return

;Align Text Top
^+t::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("AlignTopExcel")
	ObjRelease(xl)
return

;Align Left
^l::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("AlignLeft")
	ObjRelease(xl)
return

;Alignment Cycle
^e::
	xl := ComObjActive("Excel.Application")
	if (xl.ActiveWindow.Selection.HorizontalAlignment = "1" or xl.ActiveWindow.Selection.HorizontalAlignment = "-4131")
	{
	xl.CommandBars.ExecuteMso("AlignCenter")
	}
	else if (xl.ActiveWindow.Selection.HorizontalAlignment = -4108)
	{
	xl.CommandBars.ExecuteMso("AlignRight")
	}
	else 
	{
	xl.CommandBars.ExecuteMso("AlignLeft")
	}
	ObjRelease(xl)
return

;CenterAcrossSelection
^+c::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.Calculation
	;try {
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	myselection := xl.Selection
	myselection.copy
	style1 := xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range("C9")
	style1.PasteSpecial(-4122)
	style1.HorizontalAlignment := 7
	style1.Copy
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	;}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;Increase Indents
!.::
	try {
		xl := ComObjActive("Excel.Application")
		xl.CommandBars.ExecuteMso("IndentIncreaseExcel")
	}
	ObjRelease(xl)
return

;Decrease Indents
!,::
	try {
		xl := ComObjActive("Excel.Application")
		xl.CommandBars.ExecuteMso("IndentDecreaseExcel")
	}
	ObjRelease(xl)
return


;1st Level Formatting
^#1::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	
	myselection := xl.Selection
	
	/*
	myselection.copy
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Sheets(1).Range(rngAddress)
	style1.style := "1st Level"
	;style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("2nd Level").Interior.Color
	;style1.Font.Bold := True
	style1.copy
	*/
	xl.ActiveWorkbook.Sheets("style").Range("B15").Copy
	myselection.select
	If (myselection.Count = 1) {
	rwNumber := myselection.Row
	rangeAddress := "A" . rwNumber . ":BI" . rwNumber
	xl.ActiveSheet.Range(rangeAddress).Select
	selectAddress := "B" . rwNumber
	xl.CommandBars.ExecuteMso("PasteFormatting")
	;style1.Clear
	xl.ActiveSheet.Range(selectAddress).Select
	} else {
	xl.CommandBars.ExecuteMso("PasteFormatting")
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
	return

;2nd Level Formatting
^#2::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	myselection := xl.Selection
	/*
	myselection.copy
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Sheets(1).Range(rngAddress)
	;style1.PasteSpecial(-4122)
	style1.style := "2nd Level"
	;style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("2nd Level").Interior.Color
	;style1.Font.Bold := True
	style1.copy
	*/
	xl.ActiveWorkbook.Sheets("style").Range("B16").Copy
	myselection.select
	If (myselection.Count = 1) {
	rwNumber := myselection.Row
	rangeAddress := "C" . rwNumber . ":BI" . rwNumber
	xl.ActiveSheet.Range(rangeAddress).Select
	selectAddress := "C" . rwNumber
	xl.CommandBars.ExecuteMso("PasteFormatting")
	style1.Clear
	xl.ActiveSheet.Range(selectAddress).Select
	} else {
	xl.CommandBars.ExecuteMso("PasteFormatting")
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;3rd Level Formatting
^#3::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	myselection := xl.Selection
	myselection.copy
	/*
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Sheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.Style := "3rd Level"
	style1.Copy

	;style1.Interior.Color := xl.Workbooks("PERSONAL.XLSB").Styles("3rd Level").Interior.Color
	;style1.Font.Bold := xl.Workbooks("PERSONAL.XLSB").Styles("3rd Level").Font.Bold
	;style1.copy
	*/

	xl.ActiveWorkbook.Sheets("style").Range("B17").Copy
	myselection.select
	If (myselection.Count = 1) {
	rwNumber := myselection.Row
	rangeAddress := "D" . rwNumber . ":BI" . rwNumber
	xl.ActiveSheet.Range(rangeAddress).Select
	selectAddress := "D" . rwNumber
	xl.CommandBars.ExecuteMso("PasteFormatting")

	xl.ActiveSheet.Range(selectAddress).Select
	} else {
	myselection.select
	xl.CommandBars.ExecuteMso("PasteFormatting")
	style1.clear
	}
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

;4th Level Formatting
^#4::
	xl := ComObjActive("Excel.Application")
	oldCalcState := xl.calculation
	xl.Calculation := -4135
	xl.ScreenUpdating := False
	xl.EnableEvents := False
	myselection := xl.Selection
	myselection.copy

	/*
	rngAddress := myselection.Address
	style1 := xl.Workbooks("PERSONAL.XLSB").Sheets(1).Range(rngAddress)
	style1.PasteSpecial(-4122)
	style1.style := "4th Level"
	style1.copy
	*/

	xl.ActiveWorkbook.Sheets("style").Range("B15").Copy
	myselection.select

	If (myselection.Count = 1) {
		rwNumber := myselection.Row
		rangeAddress := "E" . rwNumber . ":BI" . rwNumber
		xl.ActiveSheet.Range(rangeAddress).Select
		selectAddress := "E" . rwNumber
		xl.CommandBars.ExecuteMso("PasteFormatting")
		xl.ActiveSheet.Range(selectAddress).Select
		} else {
		xl.CommandBars.ExecuteMso("PasteFormatting")
	}

	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
	xl.EnableEvents := True
	ObjRelease(xl)
return

}

;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS ;BORDERS 
{

;ThickBorderTop
^!NumPad8::
>!>#Up::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderThickOutside")
	xl.CommandBars.ExecuteMso("BorderBottom")
	xl.CommandBars.ExecuteMso("BorderLeft")
	xl.CommandBars.ExecuteMso("BorderRight")
	ObjRelease(xl)
return

;Thick Left Border
^!NumPad4::
>!>#Left::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderThickOutside")
	xl.CommandBars.ExecuteMso("BorderBottom")
	xl.CommandBars.ExecuteMso("BorderTop")
	xl.CommandBars.ExecuteMso("BorderRight")
	ObjRelease(xl)
return

;Thick Right Border
^!NumPad6::
>!>#Right::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderThickOutside")
	xl.CommandBars.ExecuteMso("BorderBottom")
	xl.CommandBars.ExecuteMso("BorderLeft")
	xl.CommandBars.ExecuteMso("BorderTop")
	ObjRelease(xl)
return

;Thick Bottom Border
^!NumPad2::
>!>#Down::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderThickOutside")
	xl.CommandBars.ExecuteMso("BorderTop")
	xl.CommandBars.ExecuteMso("BorderLeft")
	xl.CommandBars.ExecuteMso("BorderRight")
	ObjRelease(xl)
return

;Left Border
>#Left::
^NumPad4::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderLeftNoToggle")
	ObjRelease(xl)
return

;Right Border
>#Right::
^NumPad6::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderRightNoToggle")
	ObjRelease(xl)
return

;TopBorder
>#Up::
^NumPad8::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderTopNoToggle")
	ObjRelease(xl)
return

;Border Bottom
>#Down::
^NumPad2::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderBottomNoToggle")
	ObjRelease(xl)
return

;BordersALL
^NumPad5::
>#l::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BordersAll")
	ObjRelease(xl)
return

;Remove Borders
^NumPad0::
>#n::
	xl := ComObjActive("Excel.Application")
	xl.CommandBars.ExecuteMso("BorderNone")
	ObjRelease(xl)
return

}


;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION;NOTATION ;NOTATION

;Name
;AddIndent
;Address
;AddressLocal
;AllowEdit
;Application
;Areas
;Borders
;Cells
;Characters
;Column
;Columns
;ColumnWidth
;Comment
;Count
;CountLarge
;Creator
;CurrentArray
;CurrentRegion
;Dependents
;DirectDependents
;DirectPrecedents
;DisplayFormat
;End
;EntireColumn
;EntireRow
;Errors
; Font
; FormatConditions
; Formula
; FormulaArray
; FormulaHidden
; FormulaLocal
; FormulaR1C1
; FormulaR1C1Local
;HasArray
; HasFormula
; Height
; Hidden
; HorizontalAlignment
; Hyperlinks
; ID
;IndentLevel
;Interior
; Item
;Left
; ListHeaderRows
;ListObject
;LocationInTable
; Locked
; MDX
; MergeArea
;MergeCells
; Name
;Next
; NumberFormat
;NumberFormatLocal
; Offset
;Orientation
; OutlineLevel
; PageBreak
; Parent
;Phonetic
;Phonetics
; PivotCell
; PivotField
; PivotItem
; PivotTable
; Precedents
; PrefixCharacter
; Previous
;; QueryTable
 ;Range
 ;ReadingOrder
 ;esize
 ;Row
 ;owHeight
 ;ows
 ;;;howDetail
 ;Sh;rinkToFit
 ;So;;arklineGroups
 ;Sty;le
 ;Sum;mary
 ;Tex;t
 ;To;
 ;Us;eStandardHeight
 ;Us;eStandardWidth
 ;Va;idation
 ;Va;lue
 ;Va;ue2
 ;Ve;rticalAlignment
 ;Wi;dth
 ;Wo;rksheet
 ;Wr;apText
 ;XP;at;;;;;;;;;;;;;;;;;;;;;;;;;;;;