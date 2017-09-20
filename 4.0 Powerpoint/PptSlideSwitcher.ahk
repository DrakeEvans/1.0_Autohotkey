#Persistent
#IfWinActive, ahk_exe POWERPNT.EXE
#SingleInstance Force
global ppt := ComObjActive("Powerpoint.Application")   ; ID active Excel Application
ComObjConnect(ppt, "ppt_")	;Connect events

global tabHistory := Object()
global previousSlide
global thisSlide

ppt_SlideSelectionChanged() {
	
	ppt := ComObjActive("Powerpoint.Application")
	try {
		previousSlide := thisSlide
		thisSlide := ppt.ActiveWindow.View.Slide
	}
	
return
}

;Activate Previous Slide
#`::
{
	try {
		ppt := ComObjActive("Powerpoint.Application")
		
		previousSlide.Select
	}
return
}

