#SingleInstance Force

Loop, 100 {
    Loop, 3 {
        Sleep, 1500000
        InputBox, note, Short Break, What have you been doing?
        noteString := A_DDDD . " " . A_MM . "/" . A_DD . " " . A_Hour . ":" . A_Min . ": " . note . "`n"
        FileAppend, %noteString%, C:\Users\MBP\Documents\1.0_Autohotkey\Pomodoro.txt
        MsgBox, Take a short break %A_Hour% : %A_Min%
    }
    Sleep, 1500000
    InputBox, note, Short Break, What have you been doing?
    noteString := A_DDDD . " " . A_MM . "/" . A_DD . " " . A_Hour . ":" . A_Min . ": " . note . "`n"
    FileAppend, %noteString%, C:\Users\MBP\Documents\1.0_Autohotkey\Pomodoro.txt
    MsgBox, Take a long break %A_Hour% : %A_Min%
}