#NoEnv
#SingleInstance, Force

SendMode Input
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;keyhistory
;#InstallMouseHook
#installkeybdhook
SetNumLockState,On
SetNumLockState,AlwaysOn

SetScrollLockState,Off
SetScrollLockState,AlwaysOff
return
;
/*
CapsLock::
	KeyWait, CapsLock
	If (A_PriorKey="CapsLock")
		SetCapsLockState, % GetKeyState("CapsLock","T") ? "Off" : "On"
Return
*/

#If, GetKeyState("CapsLock", "P") ;Your CapsLock hotkeys go below



;i::+Tab
i::Up
j::Left
;k::Tab
k::Down
l::Right ;u::Send {Home}

t::SendInput (Tons)


[::Send {End}
]::Send {PgDn}
=::Send {PgUp}
`;::Send {BS}

;Percent
p:: SendInput `%

;Cents
c::SendInput {U+00A2}

;Mu symbol
u::SendInput {U+03BC}

Enter::SendInput {Esc}


