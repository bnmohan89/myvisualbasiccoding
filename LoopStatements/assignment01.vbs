'==================================================================================
'Assignemt 1
'Ask the user to enter a number between 1 and 25.
'If the value enter is outside the range, display a message and exit.
'If the value entered is within the range, display numbers from 1 through the number entered.
'==================================================================================

Option Explicit

Dim mInputvalue, i

mInputvalue = InputBox("Enter any value", "Please Enter you Inputs ", " Inputs Plz ")

if IsNumeric(mInputvalue) then
    mInputvalue = CInt(mInputvalue)

MsgBox " Input value is Numeric " & mInputvalue

if mInputvalue >= 25 Then
    MsgBox " Input value entered is not in the range " & mInputvalue
else
    For i = 0 to mInputvalue
        MsgBox i
        'WScript.echo ("Hurry !! Input value entered is below 25 " , & i)
    Next
end if
else
        MsgBox " Input value is not in range " & mInputvalue
End if

    