Dim InpNumber

InpNumber = InputBox("Enter any one number","Plzz Enter","Value Please")

If IsNumeric(InpNumber) Then
    InpNumber = CLng(InpNumber) ' Convert input to a numeric value
    If InpNumber >= 1 AND InpNumber <=100 Then
         MsgBox " Entered value is in b/w range 1 to 100 and the enter value is " & InpNumber
    ElseIf InpNumber >= 101 AND InpNumber <=1000 Then
          MsgBox " Entered value is in b/w range 101 to 1000 and the enter value is " & InpNumber
    Else
          MsgBox "Entered value is outside the expected range (1 to 1000)."
    End If
Else
    WScript.echo "Input Value : Please enter the numeric value."
End If