#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE

	xl := ComObjActive("Excel.Application")
	axesCount := xl.ActiveChart.Axes.Count

	Loop %axesCount% {
		try {
			defValue := xl.ActiveChart.Axes(A_Index).TickLabels.NumberFormat
			newValue := xl.InputBox("Enter the number format: " . A_Index, "Axis Format: " . A_Index, defValue)
			xl.ActiveChart.Axes(A_Index).TickLabels.NumberFormatLinked := 0
			xl.ActiveChart.Axes(A_Index).TickLabels.NumberFormat := newValue
		}
}

return



