#SingleInstance force
#IfWinActive, ahk_class XLMAIN
#MaxHotKeysPerInterval 1000

/*
F3::
    xl := ComObjActive("Excel.Application")
    formulaString := xl.Selection.Formula
    formulaArray := StrSplit(formulaString , ["+","-","*","/",")","(","=",",",";","^"])
    text := ""
    loopCount := formulaArray
    for index, element in formulaArray {
        if (element="") {

        } else {
        text := "   " . formulaArray[index] . text
    }
    msgBox, %text%
    ObjRelease(xl)
return
*/

F3::
    xl := ComObjActive("Excel.Application")
    formulaString = =SUMIFS('4.2 m_Rev'!T434:T$545,'4.2 m_Rev'!$D$434:$DE$545,$E345)
    
    references := Object()
    startPos := 1
    while (startPos > 0)
    {
        matchPos := RegExMatch(formulaString, "O)(?<sheet>['\s].*?!)?(?<cellref>\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)" , matchObj, startPos)
        startPos := matchPos + matchObj.Len[0]
        text := MatchObj.Value[0]
        msgBox, %text%
        Loop, % matchObj.Count {
            message := matchObj.Value[A_Index]
            msgBox, subpattern: %message%
        }

    }

/*
    Loop, % matchObj.Count() {
        text := matchObj.Value[0]
        MsgBox, %text%
    }
*/

    ObjRelease(xl)
return

F6::
    xl := ComObjActive("Excel.Application")
    xl.DisplayAlerts := False
    cellCount := xl.Selection.Count

    selectionAddress := xl.Selection.Address(1,0)

    newRange := xl.InputBox("Select New Range",,,,,,,8)

    newCount := newRange.Count

    newAddress := newRange.Address(1,0)

    newAddressArray := StrSplit(newAddress, ",")

    LoopCount := 0
    Loop, Parse, selectionAddress, `,
    {
        LoopCount := LoopCount +1
        ;MsgBox, %A_LoopField%
        ;MsgBox, % newAddressArray[LoopCount]
        xl.ActiveSheet.Cells.Replace(A_LoopField, newAddressArray[LoopCount],2,,True)
    }

    selectionAddress := StrReplace(selectionAddress, "$" , "")
    newAddress := StrReplace(newAddress, "$", "")
    newAddressArray := StrSplit(newAddress, ",")
    
    LoopCount := 0
    Loop, Parse, selectionAddress, `,
    {
        LoopCount := LoopCount +1
        ;MsgBox, %A_LoopField%
        ;MsgBox, % newAddressArray[LoopCount]
        xl.ActiveSheet.Cells.Replace(A_LoopField, newAddressArray[LoopCount],2,,True)
    }
    xl.DisplayAlerts := True
    ObjRelease(xl)