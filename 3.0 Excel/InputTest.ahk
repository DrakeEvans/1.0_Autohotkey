#Persistent
#SingleInstance Force


^`::
    SplashTextOn, 100, 100, Title, "Started"
    Sleep, 500
    SplashTextOff
    while (GetKeyState("``", "P") = 0 and GetKeyState("LCtrl", "P")) {
        sleep, 100
    }
    
    while (GetKeyState("LCtrl", "P")) {
     
            SplashTextOn, 100, 100, Title, "Fired"
            Sleep, 1000
            SplashTextOff
    }