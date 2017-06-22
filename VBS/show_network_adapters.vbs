On Error Resume Next
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNicConfigs = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
WScript.Echo VbCrLf & "MAC & IP Addresses & Subnet Masks"
 
For Each objNicConfig In colNicConfigs
 
  Set objNic = objWMIService.Get _
   ("Win32_NetworkAdapter.DeviceID=" & objNicConfig.Index)
 
  WScript.Echo VbCrLf & "  " & objNic.AdapterType & " " & _
   objNic.NetConnectionID
  If Err Then
    WScript.Echo VbCrLf & "  Network Adapter " & objNicConfig.Index
  End If
 
  WScript.Echo "    " & objNicConfig.Description & VbCrLf & "    MAC Address:" & VbCrLf & "        " & objNic.MACAddress & _
  "    IP Address(es):"
  For Each strIPAddress In objNicConfig.IPAddress
    WScript.Echo "        " & strIPAddress
  Next
  WScript.Echo "    Subnet Mask(s):"
  For Each strIPSubnet In objNicConfig.IPSubnet
    WScript.Echo "        " & strIPSubnet
  Next
 
Next
