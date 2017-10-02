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

AppsKey & Left::
    SendInput #{Left}
return

AppsKey & Right::
    SendInput #{Right}
return

AppsKey & Up::
    SendInput #{Up}
return

AppsKey & Down::
    SendInput #{Down}
return


#InputLevel, 1
WheelUp::
    ;MouseGetPos, ,, currentWindowID
    ;WinActivate, currentWindowID
    SendInput {WheelDown}
return

WheelDown::
    SendInput {WheelUp}
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


;Prevent accidental tooltip triggers in Office Applications 
~^ & LAlt::
    SendInput {Ctrl up}
    SendInput {Ctrl down}
return



+Esc::
    SendInput {CtrlBreak}
return

:*R:;cc::CaCO3
return

:*?R:@adrake::a.drake.evans@gmail.com
return

:*?R:@dwh::drake.evans@whitehavenltd.com
return

:*?R:@email::jennifersgilday@gmail.com
return

:*?R::;phone::9546924650
return

:*?R::@ecg::ecg.drake@gmail.com
return

:*?R:;ktn::TT11KTH6F
return

Appskey::
    Send {Appskey}
return 

:*?:;u::{U+03BC}
return

:*?:-->::{U+21FE}
return

:*?:;=::{U+2248}
return

:*?:+_::{U+00B1}
return

:*?:<--::{U+21FD}
return

:?*:;;c::{U+00A2}
return

:*:;v::{U+00FC}
return

:*?:__::{U+2013}
return

:*?:---::{U+2014}
return

:*?:/;::{U+00F7}
return

:?*:;Delta::{U+2206}
return

:?*:;.::{U+00B7}
return

:*:;cpp::({U+00A2}/lb)
return

:?*:;and::&
return

:?*:;pptobject::ppt := ComObjActive("Powerpoint.Application")
return

:?*:;tho::($, '000s)
return

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

:*?:;euro::{U+20AC}
return

:*?:;capex::Capital Expenditure
return

:*?:;ttho::(T, '000s)
return

:*?:;dol::$
return

:*?:;zoom::https://zoom.us/j/9094275712
return

:*?:;;p::%
return

:*?:;asc::Adhesives, Sealants, & Caulks
return

CapsLock::
return

