' vbs
' change_adapters_ip.vbs 
' github.com/tomwisniewskiprv/scripts/vbs
'
' Script can be used to change IP adress of network adapters
' ----------------------------------------------------------'

Dim strNewIP, strNewMask, strNewGateway ' new values

Dim strNetConnectionList	' for inputbox
iAdaptersCount  = 0

Dim arrAdaptersList			' dynamic array with named connections
arrAdaptersList = array()


' Get all network adapters
' ------------------------'
On Error Resume Next 
 
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter") 

' List all named connections 
' --------------------------'
For Each objItem in colItems 
	
	If objItem.NetConnectionID <> 0 Then

		iAdaptersCount  = iAdaptersCount + 1
		strNetConnectionList = strNetConnectionList & iAdaptersCount & chr(32) & objItem.NetConnectionID & vbNewLine
				
		' Dynamic array. Don't use with big data !
		ReDim Preserve arrAdaptersList(UBound(arrAdaptersList) + 1)
		arrAdaptersList(UBound(arrAdaptersList)) = objItem.NetConnectionID
		
	End If    
Next


' Get User input 
' --------------'
Dim iUserInput

Do While infinite = 0	

	iUserInput = Inputbox(strNetConnectionList , "Choose network adapter :")
	
	If isNumeric(iUserInput) then
		iUserInput = CInt(iUserInput)
		
		If iUserInput = 0 then Exit Do

		If ((iUserInput > 0) and (iUserInput < iAdaptersCount + 1)) then
			iUserInput = iUserInput - 1
			strAdapterName =  arrAdaptersList(iUserInput)
			Wscript.Echo iUserInput + 1 & chr(32) & strAdapterName
			Exit Do
		End If
		
	Else		
		r = msgBox("Input is not a number.", vbOKOnly , "Input error.")
	End If		
Loop 


' create regex query
' -------------------'
 Set re = New RegExp
 With re
     .Pattern    = "^(\d{1,2}|[01]\d{2}|2[0-4]\d|25[0-5])\.(\d{1,2}|[01]\d{2}|2[0-4]\d|25[0-5])\.(\d{1,2}|[01]\d{2}|2[0-4]\d|25[0-5])\.(\d{1,2}|[01]\d{2}|2[0-4]\d|25[0-5])$"
     .IgnoreCase = False
     .Global     = False
 End With

' Get new IP adress
' -----------------'
' 192.168.77.130

Do While infinite = 0
	strNewIP = InputBox("Type new IP adress:" & vbNewLine & "format:" & vbNewLine & "xxx.xxx.xxx.xxx" , "New IP for " & strAdapterName)
	 
	If re.Test( strNewIP ) Then
		Wscript.echo (strNewIP & " is a valid IP adress.")
		Exit Do
	Else
		Wscript.echo (strNewIP & " is a NOT valid IP adress. Please try again.")
	End If
Loop

' Get new submask
' -----------------'
Do While infinite = 0
	strNewMask = InputBox("Type new sub mask:" & vbNewLine & "format:" & vbNewLine & "xxx.xxx.xxx.xxx" , "New sub mask for " & strAdapterName)
	 
	If re.Test( strNewIP ) Then
		Wscript.echo (strNewMask & " is a valid submask.")
		Exit Do
	Else
		Wscript.echo (strNewMask & " is a NOT valid submask. Please try again.")
	End If
Loop

' Get new gateway
' -----------------'
Do While infinite = 0
	strNewGateway = InputBox("Type new gateway:" & vbNewLine & "format:" & vbNewLine & "xxx.xxx.xxx.xxx" , "New gateway for " & strAdapterName)
	 
	If re.Test( strNewIP ) Then
		Wscript.echo (strNewGateway & " is a valid gateway.")
		Exit Do
	Else
		Wscript.echo (strNewGateway & " is a NOT valid gateway. Please try again.")
	End If
Loop

' Pass new values to command line & and execute changes
' -----------------------------------------------------'
msg = "Please confirm following changes:" & vbNewLine & _
	  "	IP  : " & strNewIP & vbNewLine & _
	  "	Mask: " & strNewMask & vbNewLine & _
	  "	Gate: " & strNewGateway & vbNewLine

result = msgBox(msg, vbYesNo+vbInformation, "Confirm changes:")

If result = vbOk Then
	cmdNetshCommand = "netsh interface ip set address name=" & chr(34) & strAdapterName & chr(34) & " static " & _
					  strNewIP & chr(32) & strNewMask & chr(32) & strNewGateway & chr(32) & "1"
					  
	Set objShell = WScript.CreateObject("WScript.Shell")
	
	isChangeSuccess = objShell.run "cmd /c" & cmdNetshCommand, 0, True
	If isChangeSuccess = 0 Then
		Wscript.Echo "Success !"
	Else
		Wscript.Echo "Something went wrong!"
	End If

Else
	Wscript.Echo "Changes terminated."
	
End If

'netsh interface ip set address name="Virtual Loopback" static 123.123.123.123 255.255.255.0 123.123.123.1 1

