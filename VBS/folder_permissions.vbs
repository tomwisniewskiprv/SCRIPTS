' vbs
' folder_permissions.vbs
' github.com/tomwisniewskiprv/scripts/vbs
'
' Create folder for User and modify it permissions with icacls.exe
' (Polish version) #cookbook
' ----------------------------------------------------------------'

'********************************************************************
'* Main Script
'********************************************************************
' Script requires elevated privileges
' -----------------------------------'
Set WshShellA = WScript.CreateObject("WScript.Shell")	' command shell

If WScript.Arguments.Length = 0 Then
	Set ObjShell = CreateObject("Shell.Application")
	ObjShell.ShellExecute "wscript.exe" , """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
	WScript.Quit
End if

' Variables
'----------
strComputer = "."
strWQLQueryUserList = "SELECT * FROM Win32_UserAccount WHERE LocalAccount = True" 
Dim arrUsersList
arrUsersList = array()                   ' array with user names
iUserNumber = 1
strUserList = ""                         ' message containing all user names

strTestDir = "TEST_FOLDER"               ' name of test folder

' Get Users
' ---------
Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery(strWQLQueryUserList)

For Each objItem In colItems

	' Polish version of Default Account and Guest Account, skipp these
    If objItem.name <> "Konto domyślne" and objItem.name <> "Gość" and objItem.name <> "Administrator" Then
        strUserList = strUserList & iUserNumber & ") " & " " & objItem.Name & vbCr
        iUserNumber = iUserNumber + 1

        ' Dynamic array with Users
        ReDim Preserve arrUsersList(UBound(arrUsersList) + 1)
        arrUsersList(UBound(arrUsersList)) = objItem.Name
    End If
Next

' Get Input - Choose which User
' -----------------------------
iUser = InputBox(strUserList, "Users:")

' variables for MsgBox
Dim msg 
Dim style
Dim title
Dim choice

If isNumeric(iUser) Then
    iUser = CInt(iUser)

    If ((iUser > 0) and (iUser < iUserNumber)) Then
        strUserName = arrUsersList(iUser - 1)        
        
        msg = "Please confirm new permissions for User : " & chr(32) & strUserName 
        style = 4 ' YesNo
        title = "Ask for confirmation"
        choice = MsgBox(msg, style, title)

        If choice = 6 Then ' 6 - means yes
            testFolder = createFolderForUser(strUserName, strTestDir)            ' Create folder
            result = changePermissionsForFolder(testFolder, strUserName, arrUsersList) ' Change permissions
        Else
            Wscript.Echo "Abort"
            Wscript.Quit
        End If
    Else
        Wscript.Echo "Wrong choice. Please try again."
        Wscript.Quit
    End If
Else
    Wscript.Echo iUser & " is not number. Please try again."
End If

If result <> 0 Then
    Wscript.Echo "Operation finished with " & result & " errors." 
Else
    Wscript.Echo "Operation succesful."
End If

'********************************************************************
'* End of Main Script
'********************************************************************

'********************************************************************
'*
'* Function: createFolderForUser(user)
'* Purpose: creates folder for 'user' in %USERPROFILE% directory
'* Input:   user - User name
'* Output:  Path to created folder
'* Notes:   
'*
'********************************************************************
Function createFolderForUser(user, strTestDir)
    Dim objShell
    Dim objEnv
    Dim objFSO
    Dim objFolder
    Dim strDirectory
	Dim result	
    
    Set objShell = Wscript.CreateObject("Wscript.Shell")
    Set objEnv = objShell.Environment("Process")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strPath = objEnv("USERPROFILE") ' get drive and path to home folder

	tmp = split(strPath, "\")
	Dim strHomePath
	
	For i = 0 To UBound(tmp) - 1
		strHomePath = strHomePath + tmp(i) + "\"
	Next
	
	strCreateThisDirectory = strHomePath + user + "\" + strTestDir

    If Not objFSO.FolderExists(strCreateThisDirectory) Then
        Wscript.Echo "Directory does not exist. Creating directory." & chr(10) & strCreateThisDirectory
		Set result = objFSO.CreateFolder(strCreateThisDirectory)
        createFolderForUser = result.path ' returns path to folder
    Else 
        Wscript.Echo "Directory already exist."
        createFolderForUser = strCreateThisDirectory 
    End If


End Function

'********************************************************************
'*
'* Function: changePermissionsForFolder(strFolder, strUser)
'* Purpose:  Changes directory permissions with cacls or icacls command
'* Input:   strFolder, strUser - path to folder, User name
'* Output:  Number of errors. 
'* Notes:   Uses cacls.exe with elevated privileges.
'*
'********************************************************************
Function changePermissionsForFolder(strFolder , strSelectedUser, arrUsers)
    Dim strCommand
    Dim intRunError
    Dim strError
	Set WshShellA = WScript.CreateObject("WScript.Shell")	' command shell
    Set arrErrors = CreateObject("System.Collections.ArrayList") ' Error log

    For Each User in arrUsers
        If User <> strSelectedUser Then
            strCommand = "%COMSPEC% /r icacls " & strFolder & " /deny " & User & ":D" ' deny all right 
            intRunError = WshShellA.Run(strCommand, 2, True)
            strError = intRunError & chr(32) & User & chr(32) & strFolder
            If intRunError <> 0 Then
                arrErrors.add(strError)
            End If
        End If
    Next

    strCommand = "%COMSPEC% /r icacls " & strFolder & " /grant " & strSelectedUser & ":F" ' grant full access to selected User
    intRunError = WshShellA.Run(strCommand, 2, True)
    If intRunError <> 0 Then
        strError = intRunError & chr(32) & User & chr(32) & strFolder
        arrErrors.add(strError)
    End If

	If arrErrors.Count <> 0 Then
		Wscript.Echo join(arrErrors.toArray(), " ; ")
	End If

    changePermissionsForFolder = arrErrors.Count

End Function

