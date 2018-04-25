#SingleInstance Force
DetectHiddenWindows, On
SetTitleMatchMode, 2

launchScript(myProcess, myScript)
{
    Process, Exist, %myProcess%
    If (ErrorLevel <> 0) {
        ;msgBox, first if statement
        If !WinExist(myScript) {
           ; MsgBox, Success
            Run, %myScript%
        }
    } else {
        If WinExist(myScript) {
            WinClose, %myScript%
        }
    }
}

Loop, {

    launchScript("POWERPNT.EXE", "C:\Users\MBP\Documents\1.0_Autohotkey\4.0 Powerpoint\PowerpointShortcuts.ahk")
    launchScript("EXCEL.EXE", "C:\Users\MBP\Documents\1.0_Autohotkey\3.0 Excel\ExcelFormatting.ahk")
    launchScript("EXCEL.EXE", "C:\Users\MBP\Documents\1.0_Autohotkey\3.0 Excel\ExcelShortcuts.ahk")
    Sleep, 10000
}
