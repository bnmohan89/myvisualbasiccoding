'==================================================================================
'Assignemt 3
'Ask the user to enter a number between 1 and 100.
'If the value enter is outside the range, display a message and exit.
'If the value entered is within the range, display all the numbers that are divisible by 10, between 1 and the number entered.
'==================================================================================
Option Explicit
 
Dim Inputnumber, I, output

Inputnumber = InputBox ("Enter the Input value from 0 to 100" , "Plzz Enter Number", " Please Enter Numeric Value ")

if IsNumeric(Inputnumber) Then
    Inputnumber = CInt(Inputnumber)

'if Inputnumber < 1 or Inputnumber > 100 Then
if Inputnumber >= 100 Then
    WScript.echo "The number is outside the valid range (1 to 100). Exiting."
    wscript.Quit
Else
    wscript.echo " Value entered is in b/w 1 to 100 and the entered value is : " & Inputnumber
    For i = 1 to Inputnumber   
        if i Mod 10 = 0 Then
            WScript.echo  "Numbers divisible by 10 from 1 to " & inputNumber & ":" & vbCrLf & vbCrLf &  i
        Else
            WScript.Echo " Quit "
        End if
        Next
    End if

End If


