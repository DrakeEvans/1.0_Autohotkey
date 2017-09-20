#SingleInstance Force


Loop, {
    Sleep, 10000
    Run, "C:\Users\adrak\Documents\1.0_Autohotkey\3.0 Excel\ExcelFormatting.ahk"
    Run, "C:\Users\adrak\Documents\1.0_Autohotkey\3.0 Excel\ExcelShortcuts.ahk"

}

Loop, 2 {
    MsgBox, message
    sleep 100
}