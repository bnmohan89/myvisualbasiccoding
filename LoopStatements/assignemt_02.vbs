'==================================================================================
'Assignemt 1
'Ask the user to enter a number between 1 and 25.
'If the value enter is outside the range, display a message and exit.
'If the value entered is within the range, display numbers from 1 through the number entered.
'==================================================================================

'Error Handling
Option Explicit

'Variable Decleration
Dim vInputvalue

vInputvalue = InputBox ( " Enter the Value from 0 to 25", "InputVlaue MsgBox", " plz Enter the Value")

Do While vInputvalue < 25
    MsgBox " Input Value is .. " & vInputvalue
    vInputvalue = vInputvalue + 1
    MsgBox " Input value after increatment is.." & vInputvalue
Loop
    MsgBox  " ---***Outside of Loop***---- "
