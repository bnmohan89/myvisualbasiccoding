'=======================================================
' Variables, Constants, Dim, &, &_
' This Dim keyword is used to declare a variable.
' This Dine keyword is short for Dimension
' &    --> String concatenation (Ambresent)
' &_   --> Line Continuation
' Constant - value will not change and its not allow to change 
'=======================================================

' Variables
Dim amtTran1, amtTran2
Dim amtTotal
Dim StrName
Const bonus = 5

amtTran1 = 10
amtTran2 = 20
StrName = 100


' Processing
amtTotal = amtTran1 + amtTran2 + bonus

MsgBox amtTotal 'Display the value
wscript.Echo amtTotal ' Disply the result in DEBUG CONSOLE

'Concatenation
MsgBox " This is sum of " & amtTran1 & " and " & amtTran2 & " is "  & amtTotal

WScript.Sleep 10000

'Line Continuation

MsgBox " This is sum of " & amtTran1 & " and " & amtTran2 &_
            " is " & amtTotal