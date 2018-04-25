#Persistent
#NoTrayIcon
#SingleInstance, force

Loop, {
    IfWinExist, ahk_exe msiexec.exe
    {
        WinKill , ahk_exe msiexec.exe
        MsgBox, , msiexec.exe, Installation Error
    }
    Random, timetosleep, 500, 4000
    Sleep, %timetosleep%
}