

xl := ComObjActive("Excel.Application")
xl.Visible := 1
msgBox, %isVis%
xl.Quit
;xl.Open(%A_Args[1]%)

ObjRelease(xl)