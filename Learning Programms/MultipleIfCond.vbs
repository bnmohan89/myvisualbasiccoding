'Verify weather the entered 10 digit value is numeric value or not?

Dim InpNumber, ValueRes

InpNumber = InputBox("Enter thr Number From 1 to 10")

ValueRes = Mid(InpNumber,1,10)
MsgBox ValueRes

if IsNumeric(InpNumber) AND Len(InpNumber) = 10 Then
    wscript.Echo " Entered value is valid "
Else
    wscript.Echo " Entered value is not valid "
end if