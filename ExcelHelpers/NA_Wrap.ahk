#IfWinActive, ahk_exe EXCEL.EXE
#SingleInstance Force



xl := ComObjActive("Excel.Application")

mySelection := xl.Selection

cellCount := mySelection.Cells.Count

Loop, %cellCount% {

	cellFormula := LTrim(mySelection.Cells(A_Index).formula, "=")
	
	cellFormula := "=IF(" . cellFormula . "<>0, " . cellFormula . ", NA())"
	
	mySelection.Cells(A_Index).Select
	
	Sleep 100
	
	SendRaw %cellFormula%
	
	
	SendInput {Enter}
	SendInput {Enter}
	sleep 100
	
}