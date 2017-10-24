

xl := ComObjActive("Excel.Application")
try {
    xl.Visible := False
    xl.ScreenUpdating := False
    xl.Calculation := -4135
	xl.EnableEvents := False
    mainBook := xl.ActiveWorkbook
    Loop, Files, C:\Users\adrak\Documents\Mining_Logs\*
    {
        ;MsgBox, %A_LoopFileLongPath%
        sourceWorkbook := xl.Workbooks.Open(A_LoopFileLongPath)
        xl.Visible := False
        sourceWorkbook.Sheets(1).Copy(, mainBook.Sheets(1))
        sourceWorkbook.Close(False)
    }
} catch {
    xl.Visible := True
    xl.EnableEvents := True
	xl.Calculation := oldCalcState
	xl.ScreenUpdating := True
}
xl.Visible := True
xl.EnableEvents := True
xl.Calculation := oldCalcState
xl.ScreenUpdating := True