#SingleInstance Force
#MaxHotKeysPerInterval 1000




^.::
oNN := ComObjActive("OneNote.Application")
oNN.CommandBars.ExecuteMso("ChangeToBullet")
return

