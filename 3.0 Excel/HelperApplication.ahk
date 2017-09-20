#SingleInstance force
;#IfWinActive, ahk_exe EXCEL.EXE
#MaxHotKeysPerInterval 1000
#Persistent

global xlHelper := ComObjCreate("Excel.Application")
;xlHelper.Visible := True
xlHelperBook := xlHelper.Workbooks.Open("C:\Users\adrak\Documents\HelperBook.xlsx")
objectType := ComObjType(xlHelper)
objectName := ComObjType(xlHelper, "Name")
objectIID := ComObjType(xlHelper, "IID")
textBox := ComObjValue(xlHelper)
;FileAppend, %textBox%, Z:\Home\Documents\Scripts\Autohotkey\newfile.txt
MsgBox % DllCall("MulDiv", int, &xlHelper, int, 1, int, 1, str)


global xlHelper
xlHelper.Visible := True
return

^#F2::
global xlHelper
xlHelper.Visible := False
return