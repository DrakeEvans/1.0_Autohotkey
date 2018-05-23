#Hotstring EndChars ()[]{}:;'"/\,.?!`n `t
#MaxHotKeysPerInterval 1000
#SingleInstance Force

;* No end character required
;? No preceding space required
;B0 (B followed by a zero): Automatic backspacing is not done to erase the abbreviation you type.
;C: Case sensitive: When you type an abbreviation, it must exactly match the case defined in the script. Use C0 to turn case sensitivity back off.
;C1: Do not conform to typed case. Use this option to make auto-replace hotstrings case insensitive and prevent them from conforming to the case of the characters you actually type.
;O: Omit the ending character of auto-replace hotstrings when the replacement is produced. Use O0 (the letter O followed by a zero) to turn this option back off.
;R: Send the replacement text raw; that is, exactly as it appears
;Z: This rarely-used option resets the hotstring recognizer after each triggering of the hotstring. Use Z0 to turn this option back off.


;WIP WIP Copy word where cursor sits and repast
/*
CapsLock & 1::
    SendInput ^{Left}
    SendInput ^+{Right}
    SendInput ^c
    StringUpper, %clipboard%, %clipboard%
    SendInput ^v
    MsgBox, success
return
*/

;WIP WIP GUI Work
;^+1::
;   Gui, Add, Text,, Please enter your name:
;  Gui, Add, Edit, vName
; Gui, Show
;return

^#Left::
    SendInput, {Home}
return

^#Right::
    SendInput, {End}
return

^#F9::
SplashTextOn, 100, 100, Title, Scroll Left
    /*SendInput, {CtrlUp}
    SendInput, {AltUp}
    SendInput, {LWinUp}
    SendInput, {RWinUp}
    */
    SendInput, {WheelLeft}
    SplashTextOff
return


#InputLevel, 0
~WheelLeft::
    
    Loop, 2 {
        SendInput {WheelLeft}
    }
return

~WheelRight::
    Loop, 2 {
        SendInput {WheelRight}
    }
return

/*
;Prevent accidental tooltip triggers in Office Applications 
~^LAlt::
    SendInput {Ctrl up}
    SendInput {Ctrl down}
return
*/
+Esc::
    SendInput {CtrlBreak}
return

#Include C:\Users\MBP\Documents\1.0_Autohotkey\1.0 System Scripts\PersonalInfo.ahk

Appskey::
    Send {Appskey}
return 

; μ
:*?:;u::{U+03BC} ;μ
return

; ⇾
:*?:-->::{U+21FE} ;⇾
return

; ≈
:*?:;=::{U+2248}
return

; ±
:*?:+_::{U+00B1}
return

; ⇽
:*?:<--::{U+21FD}
return

; ¢
:?*:;;c::{U+00A2}
return

;checkmark in windows sometimes
:*:;v::{U+00FC}
return


:*?:__::{U+2013}
return

; —
:*?:---::{U+2014}
return

; ÷
:*?:;/::{U+00F7}
return

; ∆
:?*:;Delta::{U+2206}
return

; ·
:?*:;.::{U+00B7}
return

; (¢/lb)
:*:;cpp::({U+00A2}/lb)
return

; &
:?*:;and::&
return


:?*:;pptobject::
    SendInput, ppt := ComObjActive("Powerpoint.Application")
return

; ($, '000s)
:?*:;tho::($, '000s)
return

; ($/T)
:?*:;dpt::($/T)
return

:?*:;xlobject::
    SendInput, xl := ComObjActive("Excel.Application")
    SendInput {Enter}
    SendInput {Enter}
    SendInput {Enter}
    SendInput {Enter}
    SendInput, ObjRelease(xl)
    SendInput {Up}
    SendInput {Up}
return

; €
:*?:;euro::{U+20AC}
return

:*?:;capex::Capital Expenditure
return

:*?:;ttho::(T, '000s)
return

; $
:*?:;dol::$
return

; %
:*?:;;p::%
return


:*?:;asc::Adhesives, Sealants, & Caulks
return

CapsLock::
return

/*
~^v up::
Keywait, LControl
Sleep, 100
SendInput, {Enter}
return


~^c::
Sleep, 200
WinActivate, ahk_exe WINWORD.EXE
WinWaitActive, ahk_exe WINWORD.EXE
SendInput, ^v
SendInput, {Enter}
Sleep, 200
WinActivate, ahk_exe chrome.exe
return

*/


^#Space::
SplashImage, C:\Users\MBP\Documents\1.0_Autohotkey\2.0 Global Scripts\Calendar\Calendar.jpg
KeyWait, Space
SplashImage, Off
return

#IfWinActive, , Todoist

#a::
;tempFile := clipboard
;clipboard =
clipboard = @afternoon @evening @morning
KeyWait, ^
;ClipWait
SendInput, ^v
;clipboard := tempFile

#IfWinExist, ahk_exe Wox.exe

Esc::
    WinActivate, ahk_exe Wox.exe
    WinWait, ahk_exe Wox.exe
    SendInput, {Esc}
return

#IfWinActive, ahk_exe Code.exe

:*?B0:f'::'
return

:*?B0:r'::'
return