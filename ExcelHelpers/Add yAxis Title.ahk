#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE


xl := ComObjActive("Excel.Application")
axisTitle := xl.Inputbox("Enter Y Axis Value")
xl.screenupdating := False
xl.enableevents := False
SendInput {RAlt Down}
SendInput {RAlt Up}
Send jcaav
;sleep 500
Loop, 100 {
	try {
		xl.activechart.axes(2).axistitle.select
		Break
		} catch {
			sleep 10 
	}
}

Send {Enter}
SendRaw %axisTitle%
xl.screenupdating := True
return
