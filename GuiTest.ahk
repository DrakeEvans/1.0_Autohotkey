
    Loop, 3 {
        ;Sleep, 1000*60*45 ;Sleep for 45 minutes
        InputBox, noteText, Pomodoro, What have you been doing for the last 45 minutes
        noteText = %A_DDDD%, %A_MMMM% %A_DD%, %A_YYYY% %A_hour%:%A_Min%: %noteText% `n`n
        FileAppend, %noteText%, C:\Users\MBP\Documents\Pomodoro.txt
    }
