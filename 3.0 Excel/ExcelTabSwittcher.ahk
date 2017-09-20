#Persistent
#IfWinActive, ahk_exe EXCEL.EXE
#SingleInstance Force
global xl := ComObjActive("Excel.Application")   ; ID active Excel Application
ComObjConnect(xl, "xl_")	; connect oWord events to corresponding script functions with the prefix "xl_".
;xl.Visible := 1	 ; make xl Visible
global onOffSwitch := True


global cellSwitch
global cellIndexOffset := 0
global cellHistory := Object()
global tabIndexOffset := 0
global tabHistory := Object()
global thisCellAddress

/*
xl_SheetChange() { ;log all changes
	
	global xl

	try {
		xl.EnableEvents := False
		appendText := A_MM . "/" . A_MM . "/" . A_YYYY . " " . A_Hour . ":" . A_Min . " : " . xl.Selection.Address(False, False,1,True)
		fileLocation := "Z:\Home\Documents\xlDocumentChanges\" . xl.ActiveWorkbook.Name . ".txt"
		;MsgBox triggered
		FileAppend, %appendText%`n, %FileLocation%
		xl.EnableEvents := True
	} catch { 
		xl.EnableEvents := True
	}

}
*/


xl_SheetActivate() {	; event that fires when new sheet is activated
	
	global xl
	
	global tabIndexOffset
	global tabHistory
	
	global cellIndexOffset
	global cellHistory
	
	global previousSheet
	global thisSheet

try {
	
		previousSheet := thisSheet
		thisSheet := xl.ActiveSheet
	
		;MsgBox, %previousSheet%
		;MsgBox,  %thisIndex% 
	
	
}
return
}


xl_SheetSelectionChange() {
	
	
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


;Previous Sheet
<#`::
{
global xl
global tabIndexOffset
global tabHistory
global previousSheet
;try {
;MsgBox, triggered	
previousSheet.Activate	
;}
return
}




; F1::ComObjConnect(oWord) ; disconnect the object (stop handling events)
; F2::ComObjConnect(oWord, "oWord_")	; connect again