Option Explicit

Dim InputValue, i, output

InputValue = InputBox ("Enter the Input value from 0 to 100" , "Plzz Enter Number", " Please Enter Numeric Value ")

if IsNumeric(InputValue) Then
   InputValue =  CInt(InputValue)

If InputValue < 1 or InputValue > 100 Then
    WScript.Echo "The number is outside the valid range (1 to 100). Exiting."
    'WScript.Quit
Else
    'output = ""
    For i = 1 to InputValue
        If i Mod 10 = 0 Then
            output = output & i & vbCrLf
            'WScript.Echo "Numbers divisible by 10 from 1 to " & InputValue & ":" & vbCrLf & vbCrLf & output
            'MsgBox "Numbers divisible by 10 from 1 to " & InputValue & ":" & vbCrLf & vbCrLf & output
        End If
    Next
        if output = "" Then
            WScript.echo " Nothing there to display "
        Else
            WScript.echo "Numbers divisible by 10 from 1 to " & InputValue & ":" & vbCrLf & vbCrLf & output
        End If
End If
End If