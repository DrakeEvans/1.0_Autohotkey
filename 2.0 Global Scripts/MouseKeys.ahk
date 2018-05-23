#MaxHotKeysPerInterval 1000
#SingleInstance Force


ScrollLock:: ;Disable Scroll Lock
return

#If, GetKeyState("ScrollLock", "P") ;Your CapsLock hotkeys go below

LButton:: 
    SendInput, ^c
return

^#F7::
    global lastValue
    SendInput, ^c
    Sleep, 150
    textToAppend := clipboard

    textToAppend := ";" . textToAppend

    FileAppend, %textToAppend%, C:\Users\MBP\Desktop\flashcards.txt
return

^#F9::
    global lastValue
    SendInput, ^c
    Sleep, 150
    textToAppend := clipboard

    textToAppend := "," . textToAppend

    FileAppend, %textToAppend%, C:\Users\MBP\Desktop\flashcards.txt
return

RButton::
    SendInput, ^v
return

#If


#IfWinActive, ahk_class XLMAIN
^#F3::
^!F3::
    KeyWait, F3
    xl := ComObjActive("Excel.Application")

    loopCount := xl.Selection.Count

    InputBox, appendText, , Text to Append
    Clipboard := appendText
    Loop, %loopCount% {
        SendInput, {F2}
        Sleep, 100
        SendInput, ^v
        Sleep, 100
        SendInput, {Enter}
        Sleep, 200
    }

    ObjRelease(xl)
return

^#F1::
^!F1::

xl := ComObjActive("Excel.Application")

loopCount := xl.Selection.Cells.Count
;msgbox, %loopCount%
 Loop, %loopCount% {
     text := xl.Selection.Cells(A_Index).Formula
    ; msgbox, %text%
    url1 := StrReplace(text, "," , "%2C")
    url2 := StrReplace(url1, " ", "+")
    url3 := StrReplace(url2, "&", "%26")
    url4 := "https://scholar.google.com/scholar?q=" . url3
    run, chrome.exe %url4%
 }
    /*KeyWait, F1
    xl := ComObjActive("Excel.Application")

    loopCount := xl.Selection.Count

    InputBox, prependText, , Text to PrePend
    Clipboard := prependText
    Loop, %loopCount% {
        SendInput, {F2}
        Sleep, 100
        SendInput, {Home}
        Sleep, 100
        SendInput, ^v
        Sleep, 100
        SendInput, {Enter}
        Sleep, 200
    }

    ObjRelease(xl)
    */
return

^!F9::
^#F9::
    KeyWait, F9
    SendInput, ^{PgDn}
return

^!F7::
^#F7::
    KeyWait, F7
    SendInput, ^{PgUp}
return

#IfWinActive

^#F4::
^!F4::
Click
SendInput, 115
return

^#F5::
^!F5::
Clipboard = The file you uploaded cannot be read by the grading system or myself.  Please reupload as a .doc or try the "Print to PDF" option in the program used to prepare the original document.
SendInput, ^v
return


#IfWinActive, ahk_exe chrome.exe
~+LButton::
KeyWait, Lshift
;SendInput, #+{Left}
KeyWait, LButton, D
KeyWait, LButton
SendInput, {F6}
Sleep, 200
SendInput, ^c
Sleep, 200
;MsgBox, %Clipboard%
strNum := InStr(Clipboard, "/", , -1)
;msgBox, %strNum%
studentNumber := SubStr(Clipboard, strNum)
;msgBox, %studentNumber%

newurl := "ficheck.org/financial-goals" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/monthly-budget" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/revolving-savings" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/net-worth-statement" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/income-and-expense-statement" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/financial-ratios" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/retirement-needs" . studentNumber
Run, chrome.exe %newurl%
newurl := "ficheck.org/life-insurance" . studentNumber
Run, chrome.exe %newurl%
return