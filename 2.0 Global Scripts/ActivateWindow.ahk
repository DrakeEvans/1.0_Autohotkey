#SingleInstance Force


/*
Loop, {
    Sleep 100
    MouseGetPos,,, myWindow, OutputVarControl, 1
    if (WinActive("ahk_id" . myWindow) = 0) {
        WinActivate, ahk_id %myWIndow%
    }
}
*/
Esc::

hwnd := WinActive("A")
Loop
{
hwnd := DllCall("GetWindow",uint,hwnd,int,2) ; 2 = GW_HWNDNEXT
; GetWindow() returns a decimal value, so we have to convert it to hex
SetFormat,integer,hex
hwnd += 0
SetFormat,integer,d
; GetWindow() processes even hidden windows, so we move down the z oder until the next visible window is found
if (DllCall("IsWindowVisible",uint,hwnd) = 1)
break
}
WinActivate,ahk_id %hwnd%

return