#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
#IfWinActive, ahk_exe POWERPNT.EXE
#MaxHotKeysPerInterval 1000


global copyRange

global pointsToCentimeters := 28.35

F1::
	ppt := ComObjActive("Powerpoint.Application")
	myTable := ppt.ActiveWindow.Selection.ShapeRange(1).Table
	ppt.CommandBars.ExecuteMso("BorderNone")
	rwCount := myTable.Rows.Count
	Loop, 3 {
		;msgBox, %A_index%
		activeColumn := A_Index
		myTable.Cell(1, activeColumn).Shape.Fill.Forecolor.RGB := 6250335
		myTable.Cell(rwCount, A_Index).Shape.Fill.Forecolor.RGB := 6250335
		myTable.Cell(1, activeColumn).Shape.TextFrame.TextRange.Font.Color.RGB := 16777215
		myTable.Cell(rwCount, A_Index).Shape.TextFrame.TextRange.Font.Color.RGB := 16777215
	}
	
return
	



^#F9:: ;Button on Mouse
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F9
    ;Msgbox, you pressed control windows f12
    SendInput {PgUp}
return

^#F7:: ;Button on Mouse
    KeyWait, LControl
    KeyWait, LWin
    KeyWait, F7
    ;Msgbox, you pressed control windows f12
    SendInput {PgDn}
return


#p::
	ppt := ComObjActive("Powerpoint.Application")
	Input, inputVar, L1, {Esc}, "p,c,a"
	if (inputVar = "p") {
		GoTo input_p
	} else if (inputVar = "a") {
		GoTo input_a
	} else if (inputVar = "c") {
		GoTo input_c
	} else {
		Goto exitScript
	}

	input_p:
		;MsgBox, input p
		previousPrinter := ppt.ActivePresentation.PrintOptions.ActivePrinter
		ppt.ActivePresentation.PrintOptions.ActivePrinter := "Microsoft Print to PDF"
		
		Input, inputVar, L1, {Esc}, "c,a"
		
		if (inputVar = "a") {
			GoTo input_a
		} else if (inputVar = "c") {
			GoTo input_c
		} else {
			Goto exitScript
		}
		Goto exitScript

	input_a:
		;MsgBox, input a
		ppt.ActivePresentation.PrintOptions.RangeType := 1 ;Print All
		ppt.ActivePresentation.printout
		Goto exitScript

	input_c:
		;msgbox, input c
		ppt.ActivePresentation.PrintOptions.RangeType := 3 ;Print Current
		ppt.ActivePresentation.printout

		Goto exitScript

	exitScript:
		if (ppt.ActivePresentation.PrintOptions.ActivePrinter = "Microsoft Print to PDF") {
			ppt.ActivePresentation.PrintOptions.ActivePrinter := previousPrinter
		}
return

+Enter::
	SendInput {Esc}
return

MButton::
	SendInput {Esc}
return

;Insert Note Updated
#+n::
	try {
		ppt := ComObjActive("Powerpoint.Application")
		;(1,605,39,150,40)
		myShape := ppt.ActiveWindow.View.Slide.Shapes.AddShape(1,605,39,150,40)
		myShape.Fill.Forecolor.RGB := 5263615
		myShape.TextFrame.TextRange.Text := "Updated"
		myShape.Select
		SendInput {F2}
	}
return


;Unlock Aspect Ratio
#!p::

	try {
		ppt := ComObjActive("Powerpoint.Application")
		ppt.ActiveWindow.Selection.ShapeRange.LockAspectRatio := 0
	}
return


;copyRange
~^c::

try {
	ppt := ComObjActive("Powerpoint.Application")
	copyRange := ppt.ActiveWindow.Selection.ShapeRange(1)
}
return


;Copy New Table text into individual cells in a table

^!p:: ;Paste Table

	ppt := ComObjActive("Powerpoint.Application")

		originalTable := ppt.ActiveWindow.Selection.ShapeRange(1).Table

		newTable := ppt.ActiveWindow.View.Slide.Shapes.PasteSpecial(0)
		newTable := newTable.Table
		
		clmCount := newTable.Columns.Count
		rwCount := newTable.Rows.Count

		newRows := rwCount - originalTable.Rows.Count
		newColumns := clmCount - originalTable.Columns.Count

		if (newRows > 0) {
			Loop, %newRows% {
				originalTable.Rows.Add(3)
			} 
		} else if (newRows < 0) {
			delRows := newRows * -1
			;MsgBox, %delRows%
			Loop, %delRows% {
				originalTable.Rows(2).Delete
				
			}
		}

		if (newColumns > 0) {
			Loop, %newColumns% {
				originalTable.Columns.Add(3)
			} 
		} else if (newColumns < 0) {
			delColumns := newColumns * -1
			Loop, %delColumns% {
				originalTable.Columns(2).Delete
			}
		}
		
		Loop, %clmCount% {
				activeColumn := (A_Index)
			Loop, %rwCount% {
				activeRow := A_Index
				originalTable.Cell(activeRow, activeColumn).Shape.TextFrame.TextRange.Text := newTable.Cell(activeRow, activeColumn).Shape.TextFrame.TextRange.Text
			}
		}



	SendInput {Del}
return


;Copy as Enhanced Metafile and copy the dimensions/positions of selected shape
^!e:: ;paste Enhanced metafile
	ppt := ComObjActive("Powerpoint.Application")
	try {
		originalShape := ppt.ActiveWindow.Selection.ShapeRange(1)

		;Template Shape characteristics to copy
		tempHeight := originalShape.Height
		tempWidth := originalShape.Width
		tempLeft := originalShape.Left
		tempTop := originalShape.Top

		originalShape.Delete

		newShape := ppt.ActiveWindow.View.Slide.Shapes.PasteSpecial(2)

		newShape.LockAspectRatio := 0
		newShape.Height := tempHeight
		newShape.Left := tempLeft
		newShape.Width := tempWidth
		newShape.Top := tempTop

	} catch {

		newShape := ppt.ActiveWindow.View.Slide.Shapes.PasteSpecial(2)
	}

	;newShape.PictureFormat.CropRight := 240
	;newShape.PictureFormat.CropBottom := 120

	ppt.CommandBars.ExecuteMso("ObjectSendToBack")
return

;Copy the dimensions/positions of selected shape
^!s:: 
	KeyWait, LControl
	KeyWait, LAlt
	ppt := ComObjActive("Powerpoint.Application")
	try {
		originalShape := ppt.ActiveWindow.Selection.ShapeRange(1)

		;Template Shape characteristics to copy
		tempHeight := originalShape.Height
		tempWidth := originalShape.Width
		tempLeft := originalShape.Left
		tempTop := originalShape.Top

		;originalShape.Delete

		newShape := ppt.ActiveWindow.View.Slide.Shapes.PasteSpecial(0)
		try newShape.LinkFormat.BreakLink
		newShape.LockAspectRatio := 0
		newShape.Height := tempHeight
		newShape.Left := tempLeft
		newShape.Width := tempWidth
		newShape.Top := tempTop

	} 

	;newShape.PictureFormat.CropRight := 240
	;newShape.PictureFormat.CropBottom := 120
	Sleep, 500
	originalShape.Select
	Sleep, 500
	ppt.CommandBars.ExecuteMso("FormatPainter")
	;ppt.CommandBars.ExecuteMso("FormatPainter")
	;ppt.ActiveWindow.View.Slide.Shapes(originalShape.Name).PickUp
	sleep, 500
	originalShape.Visible := False
	xl.CommandBars.ExecuteMso("ObjectSendToBack")
	
return

;Group Shapes
#g::

	ppt := ComObjActive("Powerpoint.Application")
	ppt.CommandBars.ExecuteMso("ObjectsGroup")
return


;Ungroup Shapes
!#g::

	ppt := ComObjActive("Powerpoint.Application")
	ppt.CommandBars.ExecuteMso("ObjectsUngroup")
return


;Align Middle
^m::
	ppt := ComObjActive("Powerpoint.Application")
	txtFrame := ppt.ActiveWindow.Selection.TextRange.Parent

	If (txtFrame.VerticalAnchor = 4) {
		txtFrame.VerticalAnchor := 1
	} else if (txtFrame.VerticalAnchor = 3) {
		txtFrame.VerticalAnchor := 4 
	} else {
		txtFrame.VerticalAnchor := 3
	}
return


;Set Master Top
!#t::

	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	ppt := ComObjActive("Powerpoint.Application")
	masterTop := ppt.ActiveWindow.Selection.ShapeRange(1).Top
return


;Apply Master Top
^#t::
	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	try {
		ppt := ComObjActive("Powerpoint.Application")
		ppt.ActiveWindow.Selection.ShapeRange(1).Top := masterTop
	}
return


;Set Master Left
!#l::
	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	ppt := ComObjActive("Powerpoint.Application")
	masterLeft := ppt.ActiveWindow.Selection.ShapeRange(1).Left
return


;Apply Master Left
^#l::
	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	global masterCenter
	try {
		ppt := ComObjActive("Powerpoint.Application")
		ppt.ActiveWindow.Selection.ShapeRange(1).Left := masterLeft
	}
return


;Set Master Height
!#h::

	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	ppt := ComObjActive("Powerpoint.Application")
	masterHeight := ppt.ActiveWindow.Selection.ShapeRange(1).Height
return


;Apply Master Height
^#h::

	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	try {
		ppt := ComObjActive("Powerpoint.Application")
		ppt.ActiveWindow.Selection.ShapeRange(1).Height := masterHeight
	}
return


;Set Master Width
!#w::

	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	ppt := ComObjActive("Powerpoint.Application")
	masterWidth := ppt.ActiveWindow.Selection.ShapeRange(1).Width
return


;Apply Master Width
^#w::

	global masterHeight
	global masterWidth
	global masterLeft
	global masterTop
	try {
		ppt := ComObjActive("Powerpoint.Application")
		ppt.ActiveWindow.Selection.ShapeRange(1).Width := masterWidth
	}
return


;HideShapes
#h::

	ppt := ComObjActive("Powerpoint.Application")
	ppt.ActiveWindow.Selection.ShapeRange(1).Visible := false
return


;Unhide Shapes
+#h::

	ppt := ComObjActive("Powerpoint.Application")
	shapeCount := ppt.ActiveWindow.View.Slide.Shapes.Count
	SendInput {Esc}{Esc}
	Loop, %shapeCount% {
		If ppt.ActiveWindow.View.Slide.Shapes(A_Index).Visible = False {
			ppt.ActiveWindow.View.Slide.Shapes(A_Index).Visible := True
			ppt.ActiveWindow.View.Slide.Shapes(A_Index).Select(0)
		}
	}
return


;zero Margins
!0::

	ppt := ComObjActive("Powerpoint.Application")
	myShp := ppt.ActiveWindow.Selection.ShapeRange(1)
	myShp.textFrame2.MarginBottom := 0
	myShp.textFrame2.MarginTop := 0
	myShp.textFrame2.MarginLeft := 0
	myShp.textFrame2.MarginRight := 0
return




;Increase Margins by 0.05cm
!=::
	
	ppt := ComObjActive("Powerpoint.Application")
	myShp := ppt.ActiveWindow.Selection.ShapeRange(1)

	bottomMargin := myShp.textFrame2.MarginBottom
	topMargin := myShp.textFrame2.MarginTop
	leftMargin := myShp.textFrame2.MarginLeft
	rightMargin := myShp.textFrame2.MarginRight

	myShp.textFrame2.MarginBottom := bottomMargin + (0.05*28.5)
	myShp.textFrame2.MarginTop := topMargin + (0.05*28.5)
	myShp.textFrame2.MarginLeft := leftMargin + (0.05*28.5)
	myShp.textFrame2.MarginRight := rightMargin + (0.05*28.5)
return


;Decrease Margins by 0.05cm
!-::
	
	ppt := ComObjActive("Powerpoint.Application")
	myShp := ppt.ActiveWindow.Selection.ShapeRange(1)

	bottomMargin := myShp.textFrame2.MarginBottom
	topMargin := myShp.textFrame2.MarginTop
	leftMargin := myShp.textFrame2.MarginLeft
	rightMargin := myShp.textFrame2.MarginRight

	myShp.textFrame2.MarginBottom := bottomMargin - (0.05*28.5)
	myShp.textFrame2.MarginTop := topMargin - (0.05*28.5)
	myShp.textFrame2.MarginLeft := leftMargin - (0.05*28.5)
	myShp.textFrame2.MarginRight := rightMargin - (0.05*28.5)
return


;Copy Shape Formatting
^!+r::

	If (shp.HasTextFrame) {
		textFrame2.Autosize
		textFrame2.HorizontalAnchor
		textFrame2.MarginBottom
		textFrame2.MarginTop
		textFrame2.MarginLeft
		textFrame2.MarginRight
		textFrame2.NoTextRotation
		textFrame2.Orientation
		textFrame2.VerticalAnchor
		textFrame2.WordWrap
	}

	shp.AlternativeText
	shp.AutoShapeType
	shp.BackgroundStyle
	shp.BlackWhiteMode
	shp.Height
	shp.Left
	shp.LockAspectRatio
	shp.Rotation
	shp.ShapeStyle
	shp.Top
	shp.Visible
	shp.Width


	shp.LineFormat.BackColor
	shp.LineFormat.BeginArrowheadLength
	shp.LineFormat.BeginArrowheadStyle
	shp.LineFormat.BeginArrowheadWidth
	shp.LineFormat.DashStyle
	shp.LineFormat.EndArrowheadLength
	shp.LineFormat.EndArrowheadStyle
	shp.LineFormat.EndArrowheadWidth
	shp.LineFormat.ForeColor
	shp.LineFormat.InsetPen
	shp.LineFormat.Pattern
	shp.LineFormat.Style
	shp.LineFormat.Transparency
	shp.LineFormat.Visible
	shp.LineFormat.Weight

	shp.FillFormat.BackColor
	shp.FillFormat.ForeColor
	shp.FillFormat.GradientAngle
	shp.FillFormat.RotateWithObject
	shp.FillFormat.TextureAlignment
	shp.FillFormat.TextureHorizontalScale
	shp.FillFormat.TextureOffsetX
	shp.FillFormat.TextureOffsetY
	shp.FillFormat.TextureTile
	shp.FillFormat.TextureVerticalScale
	shp.FillFormat.Transparency
	shp.FillFormat.Visible

	shp.Adjustments.Count
	shp.Adjustments.Items()
return


;Apply Global Center Position
AppsKey & c::

	keyState := GetKeyState("RAlt", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterCenter.txt"

	If (keyState = 0) {
		ppt := ComObjActive("Powerpoint.Application")

		myShape := ppt.ActiveWindow.Selection.ShapeRange(1)
			
		shpText := myShape.Left - myShape.Width * .5
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
	MsgBox, success
return


;Set Global Height
AppsKey & h::

	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterHeight.txt"

	If (keyState = 1) {
		ppt := ComObjActive("Powerpoint.Application")

		myShape := ppt.ActiveWindow.Selection.ShapeRange(1)
			
		shpText := myShape.Height
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
return


;Set Global Width
AppsKey & w::

	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterWidth.txt"

	If (keyState = 1) {
		ppt := ComObjActive("Powerpoint.Application")

		myShape := ppt.ActiveWindow.Selection.ShapeRange(1)
			
		shpText := myShape.Width
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
return


;Set Global Top
AppsKey & t::

	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterTop.txt"

	If (keyState = 1) {
		ppt := ComObjActive("Powerpoint.Application")

		myShape := ppt.ActiveWindow.Selection.ShapeRange(1)
			
		shpText := myShape.Top
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}
return


;Set Global Left
AppsKey & l::

	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterLeft.txt"

	If (keyState = 1) {
		ppt := ComObjActive("Powerpoint.Application")
		
		myShape := ppt.ActiveWindow.Selection.ShapeRange(1)
			
		shpText := myShape.Left
		
		FileDelete , %fileLocation%
		FileAppend, %shpTExt%, %fileLocation%
	}
return


;Set Global Aspect Ratio
AppsKey & p::

	keyState := GetKeyState("LWin", "P")
	fileLocation := "C:\Users\adrak\Documents\AutoHotkey\myVariables\masterAspectRatio.txt"

	If (keyState = 1) {
		ppt := ComObjActive("Powerpoint.Application")
		
		myShape := ppt.ActiveWindow.Selection.ShapeRange(1)
			
		shpText := myShape.Width / myShape.Height
		
		FileDelete , %fileLocation%
		FileAppend, %shpText%, %fileLocation%
	}

return


AppsKey::
	SendInput {AppsKey}
return

;msoAnchorBottom	4	Aligns text to bottom of text frame.
;msoAnchorBottomBaseLine	5	Anchors bottom of text string to current position, regardless of text resizing. When you resize text without baseline anchoring, text centers itself on previous position.
;msoAnchorMiddle	3	Centers text vertically.
;msoAnchorTop	1