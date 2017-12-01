#SingleInstance force
#IfWinActive, ahk_class XLMAIN
#MaxHotKeysPerInterval 1000


F3::
    xl := ComObjActive("Excel.Application")
    formulaString := xl.Selection.Formula
    formulaArray := StrSplit(formulaString , ["+","-","*","/",")","(","=",",",";","^"])
    text := ""
    Loop, 10 {
        text := "   " . formulaArray[A_Index] . text
    }
    msgBox, %text%
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