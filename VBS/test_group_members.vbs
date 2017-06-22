strComputer = "."

Dim objGroup
Set objWMISerrvice = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set objGroupU = GetObject("WinNT://" & strComputer & "/U¿ytkownicy,group")
Set objGroupA = GetObject("WinNT://" & strComputer & "/Administratorzy,group")
Set colItems = objWMISerrvice.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = true")
Set arrUsers = CreateObject("System.Collections.ArrayList")
Set arrAdmins = CreateObject("System.Collections.ArrayList")

For Each User in objGroupU.Members
    arrUsers.add(User.Name)
    Wscript.Echo User.Name 
Next

For Each Admin in objGroupA.Members
    arrAdmins.add(Admin.Name)
    Wscript.Echo Admin.Name
Next


