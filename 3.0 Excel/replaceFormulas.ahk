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

^[::
    xl := ComObjActive("Excel.Application")
    formulaString := xl.Selection.Cells(1).Formula
    
    while (GetKeyState("LCtrl", "P")) {

        startPos := 1
        while True
        {
            matchPos := RegExMatch(formulaString, "O)(?<sheet>['\s].*?!)?(?<cellref>\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)" , matchObj, startPos)
            startPos := matchPos + matchObj.Len[0]
            ;text := MatchObj.Value[0]
            ;msgBox, Full MatchObj.Value[0]: %text%

            if (startPos > 0) {

                sheetMatch := matchObj.Value["sheet"]
                sheetMatch := StrReplace(sheetMatch, "'", "", OutputVarCount, Limit := -1)
                sheetMatch := StrReplace(sheetMatch,"!","")
                
                ;msgBox, Sheet: %sheetMatch%
                cellMatch := matchObj.Value["cellref"]
                ;msgBox, CellReference: %cellMatch%
                ;msgBox, StartPos: %startPos%  MatchPos: %matchPos%

                if (sheetMatch = "") {
                    xl.ActiveSheet.Range(cellMatch).Select
                } else {
                    xl.ActiveWorkBook.Sheets(sheetMatch).Activate
                    xl.activeSheet.Range(cellMatch).Select
                }
            } else {
                break
            }
        }
        KeyWait, [
        while (GetKeyState("[", "P") = 0 and GetKeyState("LCtrl", "P")) {
            sleep, 100
        }
    }

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