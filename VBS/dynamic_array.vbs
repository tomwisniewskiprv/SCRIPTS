' vbs
' dynamic_array.vbs
' github.com/tomwisniewskiprv/scripts/vbs
'
' Very poor dynamic array. Use only for small data because it grows super fast !
' -----------------------------------------------------------------------------'

Dim strDynamicArray
strDynamicArray = array()

For i = 0 to 3
	ReDim Preserve strDynamicArray(UBound(strDynamicArray) + 1)
	strDynamicArray(UBound(strDynamicArray)) = i & chr(32) & "Apple"
Next

For Each strElement in strDynamicArray
	Wscript.Echo strElement
Next