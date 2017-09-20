#SingleInstance force
#IfWinActive, ahk_exe EXCEL.EXE

!#c up::
xl := ComObjActive("Excel.Application")
axisTitle := xl.Inputbox("Enter X Axis Value")
xl.screenupdating := False
xl.enableevents := False
SendInput {RAlt Down}
SendInput {RAlt Up}
Send jcaah
;sleep 500
Loop, 100 {
	try {
		xl.activechart.axes(1).axistitle.select
		Break
		} catch {
			sleep 10 
	}
}
Send {Enter}
SendRaw %axisTitle%
xl.screenupdating := True
return
