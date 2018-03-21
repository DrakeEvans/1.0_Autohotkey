#SingleInstance force
#MaxHotKeysPerInterval 1000
#InputLevel 0

;Sponored Session
;ahk_class #322770
;ahk_exe TeamViewer.exe


While (True) {
     WinClose, ahk_exe TeamViewer.exe,, 2
}