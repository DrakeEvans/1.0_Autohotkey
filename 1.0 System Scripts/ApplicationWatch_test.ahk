#SingleInstance Force
DetectHiddenWindows, On
SetTitleMatchMode, 2

launchScript(myProcess, myScripts)
{
    global
    myArray := StrSplit(myScripts, ",")
    StringReplace, uniqueName, myProcess, .EXE ;We use the process name to create unique global variables and the "." causes issues in variable names

    Process, Exist, %myProcess%
    myvar := state_%uniqueName%
    ;MsgBox, state_%uniqueName% is %myvar%
    myPID := ErrorLevel
    If (ErrorLevel <> 0) {
        if state_%uniqueName% <> On
        { ;If the script runs for the first time the uniquename variable will be unset but still trigger the condition
            ;msgBox, first if statement
            state_%uniqueName% = On ;Set the unique global variable to On which indicates that the Excel is on, this is used so that the WinExist functions are only called if there has been a state change
            ;msgBox, % "myArray.count: " . myArray.MaxIndex()
            Loop, % myArray.MaxIndex() {
                If !WinExist(myArray[A_Index]) {
                    ;MsgBox, % "Running Script: " . myArray[A_Index]
                    Run, % myArray[A_Index]
                }
            }
        }
        else if state_%uniqueName% = On
        {
            xl := ComObjActive("Excel.Application")

                wbCount := xl.Workbooks.Count
                xlVisible := xl.Visible
                
                if (wbCount = 0) and (xlVisible = 0) {
                    Process, Close, %myPID%
                    state_%uniqueName% = Off
                }

            ObjRelease(xl)
        }
    }
    else if state_%uniqueName% = On
    { ;if the previous run found an excel process the variable would be "On" therefore we assume that a change of state has occured
        ;MsgBox, % "myArraycount: " myArray.Count
        Loop, % myArray.MaxIndex() {
            If WinExist(myArray[A_Index]) { ;now we can run the close ahk script thing
                ;MsgBox, closing Script %myArray%
                state_%uniqueName% = Off
                WinClose, % myArray[A_Index]
            }        
        }
    }
}

Loop, {

    launchScript("POWERPNT.EXE", "C:\Users\MBP\Documents\1.0_Autohotkey\4.0 Powerpoint\PowerpointShortcuts.ahk")
    launchScript("EXCEL.EXE", "C:\Users\MBP\Documents\1.0_Autohotkey\3.0 Excel\ExcelFormatting.ahk,C:\Users\MBP\Documents\1.0_Autohotkey\3.0 Excel\ExcelShortcuts.ahk,C:\Users\MBP\Documents\1.0_Autohotkey\3.0 Excel\ExcelTabSwittcher2.ahk")
    Sleep, 8000
}
