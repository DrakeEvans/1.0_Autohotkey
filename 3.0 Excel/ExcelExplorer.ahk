


xl := ComObjActive("Excel.Application")

    myWb := xl.Workbooks.Count
    xlVisible := xl.Visible
    msgBox, % myWb . " " . xlVisible
ObjRelease(xl)