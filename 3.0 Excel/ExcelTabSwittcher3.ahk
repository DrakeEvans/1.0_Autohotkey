#Persistent
#IfWinActive, ahk_class XLMAIN
#SingleInstance Force
global xl := ComObjActive("Excel.Application")   ; ID active Excel Application
ComObjConnect(xl, "xl_")	; connect oWord events to corresponding script functions with the prefix "xl_".
;xl.Visible := 1	 ; make xl Visible
global onOffSwitch := True



global cellIndexOffset := 0
global cellHistory := Object()

global tabIndexOffset := 0
global tabHistory := Object()






xl_SheetActivate() {	; event that fires when new sheet is activated

	global xl
	
	global tabIndexOffset
	global tabHistory
	global thisSheetIndex
	
	global cellIndexOffset
	global cellHistory
	
	;msgbox, fired

	try {

		If (tabIndexOffset = 0) {	
			
			;Add sheet index from the most recent sheet to the array of indices
			tabHistory.InsertAt(1, thisSheetIndex)
			
            ;replace thisSheetIndex value for future use to be appended when new sheet is activated
			thisSheetIndex := xl.ActiveSheet.Index
			
            loopCount := tabHistory._MaxIndex()
			Loop, %loopCount% {
                ;A_Index begins at 1 and loops though to maxIndex
				If (tabHistory[A_Index] = thisSheetIndex) { ;remove any history of the current tab, current tab index will be added to list once we activate a new tab
					tabHistory.RemoveAt(A_Index)
				}
			}

            ;Reset Cell History	
			cellHistory := []
			cellIndexOffset := 0
			
			;msgBox, %indexDifference%
			;MsgBox, %previousSheet%
			;MsgBox,  %thisIndex% 
		}
		
	}
}
return


;Previous Sheet
^`::

	global xl
	global tabIndexOffset
	global tabHistory
    
    while (GetKeyState("LCtrl", "P")) {

        ;try {
            tabIndexOffset := tabIndexOffset + 1
            selectTabIndex := tabHistory[tabIndexOffset]
            xl.ActiveWorkbook.Worksheets(selectTabIndex).Activate
            KeyWait, ``
        ;}
        while (GetKeyState("``", "P") = 0 and GetKeyState("LCtrl", "P")) {
            sleep, 100
        }
    }


    tabIndexOffset := 0
    try {

        If (tabIndexOffset = 0) {	
            
            ;Add sheet index from the most recent sheet to the array of indices
            tabHistory.InsertAt(1, thisSheetIndex)
            
            ;replace thisSheetIndex value for future use to be appended when new sheet is activated
            thisSheetIndex := xl.ActiveSheet.Index
            
            loopCount := tabHistory._MaxIndex()
            Loop, %loopCount% {
                ;A_Index begins at 1 and loops though to maxIndex
                If (tabHistory[A_Index] = thisSheetIndex) { ;remove any history of the current tab, current tab index will be added to list once we activate a new tab
                    tabHistory.RemoveAt(A_Index)
                }
            }

            ;Reset Cell History	
            cellHistory := []
            cellIndexOffset := 0
            
            ;msgBox, %indexDifference%
            ;MsgBox, %previousSheet%
            ;MsgBox,  %thisIndex% 
        }
    }

return


^F1::
	global tabHistory
	loopCount := tabHistory._MaxIndex()
	msgText :=""
	Loop %loopCount% {
		msgText := msgText . `A_Index . ":" . tabHistory[A_Index] . "  "
	}
msgBox, %msgText%
msgText2 := tabHistory[0]
msgBox, %msgText2%
return


xl_SheetSelectionChange() {
	
	global xl
	global cellHistory
	global thisCellAddress
	global cellIndexOffset
	
	If (cellIndexOffset = 0) {
	
		cellHistory.Insert(thisCellAddress)
		;thisCell := xl.Selection
		thisCellAddress := xl.Selection.Address
		
		loopCount := cellHistory._MaxIndex()
		Loop, %loopCount% {
		
			If (cellHistory[A_Index] = thisCellAddress) {
				
				cellHistory.RemoveAt(A_Index)
			}
		}
		
	}
	
}
return

;Previous Cell
<!`::
{
global xl
global cellIndexOffset
global cellHistory

try {
	cellIndexOffset := cellIndexOffset + 1

	cellIndex := cellHistory._MaxIndex() - cellIndexOffset + 1

	selectCellAddress := cellHistory[cellIndex]
	;msgbox %selectCellAddress%
	xl.ActiveSheet.Range(selectCellAddress).Select
}
return
}

;Reset Cell Index Offset

~<!` up::
{
Keywait LAlt
global xl
global cellHistory
global cellIndexOffset := 0
global thisCellAddress

cellIndexOffset := 0

try {
	cellHistory.Insert(thisCellAddress)
	thisCellAddress := xl.Selection.Address

	loopCount := cellHistory._MaxIndex()
			Loop, %loopCount% {
			
				If (cellHistory[A_Index] = thisCellAddress) {
					
					cellHistory.RemoveAt(A_Index)
				}
			}
	;msgBox, %cellIndexOffset%
}
return
}


;LWin up::return
!#F1::
global tabHistory

textBox := " "

loopCount := tabHistory._MaxIndex()

Loop, %loopCount% {

textBox := textBox . A_Index . ":[" . tabHistory[A_Index] . "] "

}

MsgBox, %textBox% `n TabIndexOffset: %tabIndexOffset%
return


; F1::ComObjConnect(oWord) ; disconnect the object (stop handling events)
; F2::ComObjConnect(oWord, "oWord_")	; connect again