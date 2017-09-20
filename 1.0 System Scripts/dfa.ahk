#NoEnv
#NoTrayIcon
#SingleInstance, Force

Loop, 10 {
    try SoundSet, 0,,, %A_Index%
}
