' vbs
' execute_elevated.vbs
' github.com/tomwisniewskiprv/scripts/vbs

' -------------------------------------------------------'
'TODO'

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "wscript.exe", "", "", "runas", 1
