#Persistent
#IfWinActive, ahk_class XLMAIN
#SingleInstance Force
global xl := ComObjActive("Excel.Application")   ; ID active Excel Application
ComObjConnect(xl, "xl_")	; connect oWord events to corresponding script functions with the prefix "xl_".
;xl.Visible := 1	 ; make xl Visible

/*
wbCount := xl.Workbooks(1).Name
msgbox, % wbCount
if !(wbCount > 0) {
	MsgBox, % wbCount . ": " . A_ScriptFullPath
	run, %A_ScriptFullPath%
}
*/

global onOffSwitch := True



global cellIndexOffset := 0
global cellHistory := Object()

global tabIndexOffset := 0
global tabHistory := Object()






xl_SheetActivate(){	; event that fires when new sheet is activated
	
	global xl
	
	global tabIndexOffset
	global tabHistory
	global thisSheetIndex
	
	global cellIndexOffset
	global cellHistory
	
	;msgbox, fired

	try {
		

		If (tabIndexOffset = 0) {	
			
			tabHistory.Insert(thisSheetIndex)
			;msgBox, Sheet Added %thisSheetIndex%
			
			loopCount := tabHistory._MaxIndex()
				
			thisSheetIndex := xl.ActiveSheet.Index

			Loop, %loopCount% {
			
				If (tabHistory[A_Index] = thisSheetIndex) {
					tabHistory.RemoveAt(A_Index)
				}
			}
				
			cellHistory := []
			cellIndexOffset := 0
			
			;msgBox, %indexDifference%
			;MsgBox, %previousSheet%
			;MsgBox,  %thisIndex% 
		}
		
	}
return
}

;Reset Tab Index Offset when Lwin is released
<#`::

	Keywait LWin
		;msgbox, reset tab index
	global xl
	global tabIndexOffset
	global tabHistory
	global thisSheetIndex

	tabIndexOffset := 0
	try {
		tabIndexOffset := 0	
		
		tabHistory.Insert(thisSheetIndex)
		
		loopCount := tabHistory._MaxIndex()
		
		thisSheetIndex := xl.ActiveSheet.Index
		
		Loop, %loopCount% {
			
				If (tabHistory[A_Index] = thisSheetIndex) {
					
					tabHistory.RemoveAt(A_Index)
				}
		}
	}
	;msgbox, reset
return


;Previous Sheet
<#` up::

	global xl
	global tabIndexOffset
	global tabHistory

	try {
		Keywait, ``
		Next_Sheet:

		tabIndexOffset := tabIndexOffset + 1
		
		tabIndex := tabHistory._MaxIndex() - tabIndexOffset + 1
		
		selectTabIndex := tabHistory[tabIndex]
		
		;MsgBox, %selectTabIndex%
		
		xl.ActiveWorkbook.Worksheets(selectTabIndex).Activate
		KeyWait:
		Keywait, ``, D T0.1
		
		If (ErrorLevel = 0) {
		
		GoTo Next_Sheet
		} Else {
			GoTo KeyWait
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
msgText2 := tabHistory[0]
msgBox, %msgText2%
return


; F1::ComObjConnect(oWord) ; disconnect the object (stop handling events)
; F2::ComObjConnect(oWord, "oWord_")	; connect again