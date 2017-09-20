#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
#MaxHotKeysPerInterval 1000

;Insert Default Numbered List
^/::
{
oWord := ComObjActive("Word.Application")
oWord.CommandBars.ExecuteMso("Numbering")
return
}

;Insert Default Bulleted List
^.::
{
oWord := ComObjActive("Word.Application")
oWord.CommandBars.ExecuteMso("Bullets")
return
}