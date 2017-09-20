#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE

try {

xl := ComObjActive("Excel.Application")

If (xl.ScreenUpdating = True) {
	xl.ScreenUpdating := False
} else {
	xl.ScreenUpdating := True 
}
}
return
