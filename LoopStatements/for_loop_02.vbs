'=================================================================
' Loop: For Loop with array
' Increament Index 0 to 9
'=================================================================

Option Explicit
' Variables
Dim total, arrNums, I

'Array

arrNums = array(1, 2, 3, 4, 5, 6, 6, 8, 9, 10)


'Using array concepts ina FOR loop
For i = 0 to 10
'For i = LBound(arrNums) to UBound(arrNums)
    total = total + i
    if total > 10 Then
        WScript.Echo " Hrray! The Total is greater than 10!" & vbCrLf & " current total is " & total & "! !"
    End If
        WScript.Echo "-----"

    Next
         WScript.Echo " The sum of all the elements in the array is " & total