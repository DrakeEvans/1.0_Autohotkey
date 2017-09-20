GetClipboardData(ByRef data)
{
	local cfFormat, hData, pData, dataL

	CBID := DllCall("RegisterClipboardFormat", Str, "HTML FORMAT", UInt)
		
	If (DllCall("IsClipboardFormatAvailable", UInt, CBID) <> 0)
		If (DllCall("OpenClipboard", UInt, 0) <> 0)
			If hData := DllCall("GetClipboardData", UInt, CBID, UInt)
				dataL := DllCall("GlobalSize", UInt, hData, UInt)
				, pData := DllCall("GlobalLock", UInt, hData, UInt)
				, VarSetCapacity(data, dataL * ( A_IsUnicode ? 2 : 1 ) ), StrGet := "StrGet"
				, A_IsUnicode ? Data := %StrGet%( pData, DataL, 0 )
								: DllCall( "lstrcpyn", Str, data, UInt, pData, UInt, dataL)
				, DllCall( "GlobalUnlock", UInt, hData)
DllCall("CloseClipboard")
Return dataL ? dataL : 0
}

GetClipboardData(str)

;MsgBox, %str%
;sHTMLFragment := str

FileDelete, Z:\Home\Documents\FullHTML.txt
FileAppend, %str%, Z:\Home\Documents\FullHTML.txt
x := RegExMatch(str, "s)<!--StartFragment-->.*<!--EndFragment-->", tableHTML)
FileDelete, Z:\Home\Documents\HTMLClipboardData.txt
FileAppend, %tableHTML%, Z:\Home\Documents\HTMLClipboardData.txt

;clipboard := tableHTML