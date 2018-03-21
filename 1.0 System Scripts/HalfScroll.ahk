#MaxHotkeysPerInterval, 1000
#SingleInstance, force


global myDown := 1
global myUp := 1
global myNumerator := 1
global myDenominator := 4
WheelDown::
    global myNumerator
    global myDenominator
    global myDown
    test := Mod(myDown,myDenominator)
    if (test = 1) {
        SendInput, {WheelDown}
    }
    myDown := myDown*myNumerator + 1
return

WheelUp::
    global myNumerator
    global myDenominator
    global myUp
    test := Mod(myUp,myDenominator)
    if (test = 1) {
        SendInput, {WheelUp}
    }
    myUp := myUp*myNumerator + 1
return

