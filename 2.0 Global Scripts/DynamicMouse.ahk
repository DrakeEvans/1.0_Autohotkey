#Hotstring EndChars ()[]{}:;'"/\,.?!`n `t
#MaxHotKeysPerInterval 100000
#SingleInstance Force

WheelUp::
;SplashTextOn, 100, 25, , % A_TimeSincePriorHotkey
    if (A_TimeSincePriorHotkey) > 200 {  
        SendInput {WheelUp}
    } else if (A_TimeSincePriorHotkey) > 45 {
        Loop, 3 {
            SendInput {WheelUp}
        }
    } else {
        Loop, 6 { ;less than 60
            SendInput {WheelUp}
        }        
    }
return


WheelDown::
;SplashTextOn, 125, 25, , % A_TimeSincePriorHotkey
    if (A_TimeSincePriorHotkey) > 200 {  
        SendInput {WheelDown}
    } else if (A_TimeSincePriorHotkey) > 45 {
        Loop, 3 {
            SendInput {WheelDown}
        }
    } else { ;less than 60
        Loop, 6 {
            SendInput {WheelDown}
        }        
    }
return


