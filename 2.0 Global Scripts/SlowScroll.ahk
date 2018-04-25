#MaxHotkeysPerInterval, 1000


global myUCount
global myDCount
WheelUp::
    global myUCount
    if (myUCount) == 1 {
        SendInput {WheelUp}
        myUCount := 0
    } else {
        myUCount := 1
    }
return

WheelDown::
    global myDCount
    if (myDCount) == 1 {
        SendInput {WheelDown}
        myDCount := 0
    } else {
        myDCount := 1
    }
return