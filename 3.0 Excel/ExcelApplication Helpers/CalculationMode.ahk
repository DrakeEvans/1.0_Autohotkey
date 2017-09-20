#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE

try {

xl := ComObjActive("Excel.Application")

If (xl.Calculation = 2) {
	xl.Calculation := -4135
	MsgBox, Manual Calculation
} else {
	xl.Calculation := 2 
	MsgBox, Calculation SemiAutomatic
}
}
return