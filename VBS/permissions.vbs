' rem set only readonly right to file Dane.txt for user
' cacls.exe c:\dane.txt /g "user":r

' rem show permissions of directory:
' cacls.exe c:\windows

' rem ustawia pe�ne upradnienia dla pliku:
' cacls.exe c:\dane.txt /g "user":f

' rem zezwala tylko na zapis do pliku:
' cacls.exe c:\dane.txt /g "user":windows

' rem zabranianie usuni�cia pliku
' cacls.exe c:\dane.txt /deny "user":D

' rem zabranianie wykonanie pliku
' cacls.exe c:\dane.txt /deny "user":rx

' rem przywr�cenie wszyskich praw zapisu i odczytu
' cacls.exe c:\dane.txt /grant "user":f

' cacls.exe c:\dane.txt /grant "user":n

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")

strWQLQuery1 = "Select * From win32_useraccount"

Set colItems = objWMIService.ExecQuery(strWQLQuery1)

for each objitem in colItems
	wscript.ech "user:" & objitem.Name
Next