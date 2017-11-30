#SingleInstance force
#IfWinActive, ahk_class XLMAIN
#MaxHotKeysPerInterval 1000


;Mouse Key
^#F1::
return

;Mouse Key
^#F2::
return

;Mouse Key
^#F3::  
return

;Mouse Key
^#F4::
return

;Mouse Key
^#F5::
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F5
    SendInput {F4}
return

;Mouse Key
^#F6::
return

;Mouse Key
^#F7:: ;G17 Button on Mouse
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F7
    ;Msgbox, you pressed control windows f12
    SendInput ^{PgUp}
return

;Mouse Key
^#F8::
	SendInput ^x
return

;Mouse Key
^#F9:: ;G17 Button on Mouse
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F9
    ;Msgbox, you pressed control windows f12
    SendInput ^{PgDn}
return

;Mouse Key
^#F10::
return

;Mouse Key
^#F11::
	SendInput ^+=
return

;Mouse Key
^#F12::
return
