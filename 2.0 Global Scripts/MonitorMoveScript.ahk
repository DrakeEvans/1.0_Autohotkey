#Hotstring EndChars ()[]{}:;'"/\,.?!`n `t
#MaxHotKeysPerInterval %mySleep%
#SingleInstance Force

^+x::
    mySleep := 50
    WinGetActiveTitle, myWindow
    Sleep, %mySleep%
    KeyWait, x
    KeyWait, LControl
    Sleep, %mySleep%
    WinMaximize, % myWindow ;Maximize Active Window
    Sleep, %mySleep%
    SendInput {LWin down}
    Sleep, %mySleep%
    SendInput {Right}
    Sleep, %mySleep%
    SendInput {Right}
    SendInput {Up}
    SendInput {Up}
    Sleep, %mySleep%
    SendInput {LWin up}
    Sleep, %mySleep%
    WinMaximize, % myWindow ;Maximize Active Window
    SplashTextOff
return