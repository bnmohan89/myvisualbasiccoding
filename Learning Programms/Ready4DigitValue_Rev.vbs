Dim InpBoxNum, RevStrVal


InpBoxNum = InputBox("Enterr the Value", Hi, Hello)

if Len(InpBoxNum) = 4 Then
    RevStrVal = StrReverse(InpBoxNum)
    MsgBox ("Input Value is   :  " &  InpBoxNum  & vbCrLf & "Reversed Value is   :   " & RevStrVal)
Else
    MsgBox ("Input Number is not equal to 4 digits")
End If