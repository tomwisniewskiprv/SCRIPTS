
Function getInput(text)
	' Displays input box for diffrent parametrs to change
	' getInput = "return value" & chr(32) & data & chr(32) & text
	
	Do While infinite = 0
	strNewIP = InputBox("Type new " & text &" adress:" & vbNewLine & "format:" & vbNewLine & "xxx.xxx.xxx.xxx" , "New " & text & " for " & strAdapterName)

	If strNewIP = "" Then 		
		Wscript.Echo "Quiting!"
		Wscript.Quit
	End If
	
	If re.Test( strNewIP ) Then
		Wscript.echo (strNewIP & " is a valid " & text & " adress.")
		Exit Do
	Else
		Wscript.echo (strNewIP & " is a NOT valid " & text & " adress. Please try again.")
	End If
	Loop
	
End Function

newIP  = getInput("IP")
wscript.echo m
