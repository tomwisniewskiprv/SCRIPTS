' vbs
' execute_hidden.vbs
' github.com/tomwisniewskiprv/scripts/vbs
'
' execute hidden command
' -------------------------------------------------------'

Set objShella = WScript.CreateObject("WScript.Shell")
Dim writen 
writen = False

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
              
			  letter = wmiLogicalDisk.DeviceID
			  
			  if letter <> "C:" and writen <> True Then				
				objShella.run "cmd /c ping -n 5 127.0.0.1 >"& letter & "\output.txt", 0, True
				writen = True
			  End If

        Next      
    Next
Next