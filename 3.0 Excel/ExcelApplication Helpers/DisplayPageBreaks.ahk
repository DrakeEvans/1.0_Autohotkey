#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE

try {
xl := ComObjActive("Excel.Application")

If (xl.ActiveSheet.DisplayPageBreaks = -1) {
	xl.ActiveSheet.DisplayPageBreaks := False
	msgBox Page Breaks Disabled
} else {
	xl.ActiveSheet.DisplayPageBreaks := True 
	msgBox page breaks enabled
}	
} catch {
	msgbox fail
}
return