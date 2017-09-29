#SingleInstance

cmdWindow := ComObjCreate("WScript.Shell")

execScript := "cmd.exe /k cd .."

cmdWindow.Exec("cmd.exe /k cd ..")

sleep 100

cmdWindow.Exec("cmd.exe /k cd ..")


