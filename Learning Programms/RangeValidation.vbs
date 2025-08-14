'Read a number and verify that number range whether in between 1 to 100 or 101 to 100
Option Explicit

Dim InpNumber, i

InpNumber = InputBox("Enter any one number","Plzz Enter","Value Please")

InpNumber = CDbl(InpNumber)

If InpNumber<=100 Then
    For i =  1 to 100
        if InpNumber = i Then
            MsgBox " Entered value is in b/w range 1 to 100 and the enter value is " & InpNumber
        End If
    Next
Else
    For i = 101 to 1000
        if InpNumber = i Then
            MsgBox " Entered value is in b/w range 101 to 1000 and the enter value is " & InpNumber
        End If
    Next
End If
