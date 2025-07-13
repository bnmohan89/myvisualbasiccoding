'==================================================================================
'Assignemt 3
'Ask the user to enter a number between 1 and 100.
'If the value enter is outside the range, display a message and exit.
'If the value entered is within the range, display all the numbers that are divisible by 10, between 1 and the number entered.
'==================================================================================

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
        End If
    Next

if output = "" Then
        MsgBox "There are no numbers divisible by 10 in this range."
    Else
            MsgBox "Numbers divisible by 10 from 1 to " & InputValue & ":" & vbCrLf & vbCrLf & output
            WScript.echo "Numbers divisible by 10 from 1 to " & InputValue & ":" & vbCrLf & vbCrLf & output
End If
End If
Else
    MsgBox "Invalid input. Please enter a numeric value. Exiting."
    WScript.Quit
End If