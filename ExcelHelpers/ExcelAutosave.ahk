#SingleInstance force
#IfWinActive, ahk_xl EXCEL.xl


oExcel := ComObjActive("Excel.Application")
;try {
		;try {
		
		currentDate := A_MM . A_DD . " " . A_Hour . "," . A_Min . "." . A_Sec
		;MsgBox, %currentDate%
		oExcel.CalculateBeforeSave := False
		workbookCount := oExcel.Workbooks.Count
		;MsgBox, %workbookCount%
		
		Loop %workbookCount% {
			;fileName := oExcel.Workbooks(A_Index).Name
			;bVisible := oExcel.Windows(filename).Visible
			;fullPath := oExcel.Workbooks(A_Index).FullName
			
			if (bVisible = -1) {
				;fileType := SubStr(fileName,InStr(fileName, "." , False, -5, 1))
				;fileName := "C:\Users\adrak\Documents\ADE Autosave\" . RTrim(fileName,fileType) . currentDate . fileType
				
				;MsgBox, %fileName%
				;FileCopy, %fullPath%, %fileName%
				oExcel.Workbooks(A_Index).Save
			}
		}
	;}
	oExcel.CalculateBeforeSave := True
;}
return
