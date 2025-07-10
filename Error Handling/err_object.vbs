'===========================================================================
' Err Object
'       Contains information about run-time errors.
'       Accept the Raise and Clear method for generating
'       and clearing run-time errors.
'       >> Raise & Clear Methods
'===========================================================================
Option Explicit

'Declare variables
Dim num1, num2, div

num1 = 10
num2 = 0

'processing

On Error Resume Next 
div = num1 / num2
MsgBox div

MsgBox "Error: " & Err.Number & " " & Err.Description

Err.Raise 101
MsgBox "Error: " & Err.Number & " " & Err.Description

Err.Clear 'Reset
MsgBox "Error: " & Err.Number & " " & Err.Description & "EOM"
WScript.Echo "Error: " & Err.Number & " " & Err.Description & "EOM"