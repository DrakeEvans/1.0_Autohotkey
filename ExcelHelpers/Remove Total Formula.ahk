#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE


xl := ComObjActive("Excel.Application")
oldCalcState := xl.Calculation
;try {
xl.Calculation := -4135
xl.ScreenUpdating := False
xl.EnableEvents := False


replaceString := xl.InputBox("String to Replace")
newString := xl.InputBox("New String")

selectionAddress := xl.Selection.Address 

Loop, parse, selectionAddress, `,
{
myselection := xl.ActiveSheet.Range(A_LoopField)
myselectionAddress := myselection.Address
myselection.copy
style1 := xl.Workbooks("PERSONAL.XLSB").Worksheets("Sheet1").Range(A_LoopField)
style1.PasteSpecial()
;style1.select
 cellCount := style1.Cells.Count
Loop, %cellCount% {
	cllformula := style1.cells(A_Index).formula
	cllformula := StrReplace(cllformula,"=","")
	cllformula := "=SUBSTITUTE(" . cllformula . "," . """" . replaceString . """" . "," . """" . newString . """)"

	style1.cells(A_Index).formula := cllformula
}

style1.copy
myselection.select
xl.CommandBars.ExecuteMso("Paste")
}


myselection.Calculate
xl.Calculation := oldCalcState
xl.ScreenUpdating := True
xl.EnableEvents := True
return
return
