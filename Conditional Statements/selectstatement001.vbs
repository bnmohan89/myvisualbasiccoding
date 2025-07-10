'=====================================================================================
'Select Case with Built in function
'=====================================================================================
'Error Handling
Option Explicit


Dim mNumber

mNumber = InputBox("Enter the Value", 1 , "Enter the number here")

Select Case mNumber
    Case 100
    MsgBox "We got 100"
    Case 100
    MsgBox "We got 200"
    Case 200
    MsgBox "We got 300"
    Case 300
    MsgBox "We got 300"
    Case 400
    MsgBox "We got 400"
    Case Else
        MsgBox "We got different number"
End Select