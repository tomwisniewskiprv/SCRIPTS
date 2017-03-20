' vbs
' proc_terminate_process_with_high_demand.vbs
' github.com/tomwisniewskiprv/scripts/vbs
'
' Script terminates processes with high demand for cpu.
' -------------------------------------------------------'

CONST MAX_PROCS = 50

strComputer = "."    
proc_number = 0
Dim proc_to_kill(50)


' Power consuption level , default 50%
pow_consumption = InputBox("Enter value:", "Power consumption, default 50%", 50)	

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")

' Get list of all processes
Set colPerfProcessList = objWMIService.ExecQuery(_
    "SELECT * FROM Win32_PerfFormattedData_PerfProc_Process")

For Each process in colPerfProcessList
    If process.PercentProcessorTime > pow_consumption and _
        process.Name <> "Idle" and process.Name <> " System Idle Process" and _
        process.Name <> "_Total" and process.Name <> "System" then

        If proc_number < MAX_PROCS then
            proc_to_kill(proc_number) = process.IDProcess
            proc_number = proc_number + 1
        End if

    End if
Next

' Get process one by one and terminate it.

Dim proc_terminated(50)
count_terminated = 0

If proc_number > 0 then	
	For index = 0 to proc_number
		Set colProcessList = objWMIService.ExecQuery( _
			"SELECT * FROM Win32_Process WHERE ProcessId ='"& proc_to_kill(index) &"'" )
		
		For Each proc in colProcessList		
			proc_terminated(count_terminated) = proc_to_kill(index) & " " & proc.Name & VBNewLine
			count_terminated = count_terminated + 1			
			
			proc.terminate()
		Next
	Next
	
	' Display result
	result = "Terminated: " & count_terminated & VBNewLine 

	For index = 0 to count_terminated
		result = result & proc_terminated(index)
	next

	Wscript.Echo result
	
end if


Wscript.Quit

' Library :
' https://msdn.microsoft.com/en-us/library/aa394323(v=vs.85).aspx
' https://msdn.microsoft.com/en-us/library/aa394599(v=vs.85).aspx
' https://msdn.microsoft.com/en-us/library/aa394372(v=vs.85).aspx
' https://www.tutorialspoint.com/vbscript/index.htm