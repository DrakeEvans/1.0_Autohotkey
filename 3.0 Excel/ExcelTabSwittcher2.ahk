#Persistent
#IfWinActive, ahk_exe EXCEL.EXE
#SingleInstance Force
global xl := ComObjActive("Excel.Application")   ; ID active Excel Application
ComObjConnect(xl, "xl_")	; connect oWord events to corresponding script functions with the prefix "xl_".
;xl.Visible := 1	 ; make xl Visible
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


;Previous Sheet
<#`::
{
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
	
	Keywait, ``, D T4
	
	If (ErrorLevel = 0) {
	
	GoTo Next_Sheet
	}
	
}
return
}

;Reset Tab Index Offset
~<#` up::
{
Keywait LWin
	;msgbox, reset tab index
global xl
global tabIndexOffset := 0
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
return
}



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