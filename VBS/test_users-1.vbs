strComputer = "."
Dim objGroup
Set objWMISerrvice = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
	Set objGroup = GetObject("WinNT://" & strComputer & "/U¿ytkownicy,group")
Set colItems = objWMISerrvice.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = true")

'Potrzebujemy tego by uruchomiÄ‡ cmd jako administrator
Set WshShell = WScript.CreateObject("WScript.Shell")
strUser = WshShell.ExpandEnvironmentStrings( "%USERNAME%" )
If WScript.Arguments.Length = 0 Then
	Set ObjShell = CreateObject("Shell.Application")
	ObjShell.ShellExecute "wscript.exe" _
    , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
	WScript.Quit
End if

Wscript.Echo "Actual user loged in: " & strUser

Set myArray = CreateObject("System.Collections.ArrayList")
index = 0

For Each user in objGroup.Members

if user.Name <> vbEmpty then
myArray.add user.Name
strMsg = strMsg & index & " " & myArray(index) & vbNewLine
index = index + 1
end if

Next
index = 0

correctUser = false
Do While correctUser = false
	sInput = InputBox("Choose User " & _
	vbNewLine & strMsg, ,"Choose one option")
	
If IsEmpty(sInput) then 
  WScript.Quit
End if

if (isNumeric(sInput)) then
	sInput = CInt(sInput)

if(sInput < myArray.Count) then
	correctUser = true
	Wscript.Echo "Correct!"
end if

end if
	if correctUser = false then
		Wscript.Echo "Try Again"
	end if
loop

strFolder = "C:\Users\" & strUser & "\Desktop\test"
WScript.Echo strFolder

set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists(strFolder) = false then
	objFSO.CreateFolder strFolder
	wscript.echo "Folder Created"
else
	wscript.echo "Folder already exists"
end if

result = MsgBox ("Should the User: " & myArray(sInput) & " had the only access?", vbYesNo + vbQuestion, "Yes No Example")
If IsEmpty(result) then 
  WScript.Quit
End if

Select Case result
Case vbYes
    MsgBox("User: " & myArray(sInput)& " will have the only access")
	

For Each user in objGroup.Members

	if user.Name <> myArray(sInput) then

		WshShell.run "cmd /R icacls.exe C:\Users\" & strUser & "\Desktop\test /deny " & user.Name & ":D"
		WScript.Echo "cmd /R icacls.exe C:\Users\" & strUser & "\Desktop\test /deny " & user.Name & ":D"

	end if
	
	if user.Name = myArray(sInput) then

		WshShell.run "cmd /R icacls.exe C:\Users\" & strUser & "\Desktop\test /grant " & user.Name & ":f"
		WScript.Echo "cmd /R icacls.exe C:\Users\" & strUser & "\Desktop\test /grant " & user.Name & ":f"
	end if

Next
Case vbNo
    MsgBox("User: " & myArray(sInput)& " won't have the only access")
	
For Each user in objGroup.Members

	WshShell.run "cmd /R icacls.exe C:\Users\" & strUser & "\Desktop\test /grant " & user.Name & ":f"
	WScript.Echo "cmd /R icacls.exe C:\Users\" & strUser & "\Desktop\test /grant " & user.Name & ":f"

Next
End Select