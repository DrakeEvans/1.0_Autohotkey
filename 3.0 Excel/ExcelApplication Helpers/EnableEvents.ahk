#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE

try {
MsgBox, %xl.EnableEvents%
xl := ComObjActive("Excel.Application")

If (xl.EnableEvents = 1) {
	xl.EnableEvents := 0
	MsgBox, Events Disabled
} else {
	xl.EnableEvents := True 
	MsgBox, Events Enabled
}
}
return