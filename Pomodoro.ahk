#SingleInstance Force

filename := "C:\Users\MBP\Documents\12.0 Work Tracking" . A_DDDD . " " . A_MM . "/" . A_DD . "txt"
Loop, 100 {
    Loop, 3 {
        Sleep, (45*1000*60) ;45 minutes
        InputBox, note, Short Break, What have you been doing?
        noteString := A_DDDD . " " . A_MM . "/" . A_DD . " " . A_Hour . ":" . A_Min . ": " . note . "`n"
        FileAppend, %noteString%, %filenam%
        MsgBox, Take a short break %A_Hour% : %A_Min%
    }
    Sleep, (45*1000*60) ;45 minutes
    InputBox, note, Short Break, What have you been doing?
    noteString := A_DDDD . " " . A_MM . "/" . A_DD . " " . A_Hour . ":" . A_Min . ": " . note . "`n"
    FileAppend, %noteString%, %filename%
    MsgBox, Take a long break %A_Hour% : %A_Min%
}