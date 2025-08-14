'Read a number and verify that number range whether in between 1 to 100 or 101 to 100
Option Explicit

Dim InpNumber, i, j

InpNumber = InputBox("Enter any one number","Plzz Enter","Value Please")

If InpNumber<=100 Then
    MsgBox(" Enter Value is in b/w 1 to 100 & the value is "  & InpNumber )
Else
    MsgBox(" Enter Value is in b/w 101 to 1000 & the value is "  & InpNumber )
End If




    