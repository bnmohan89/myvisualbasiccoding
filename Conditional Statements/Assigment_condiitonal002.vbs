'====================================================
' Assignment002
'====================================================

Dim mName1, mName2
Const TitleName = "Input Values"

mName1 = InputBox("Enter the John's Age", TitleName , "Your Age Please")
mName2 = InputBox("Enter the Kerry's Age", TitleName, "Your Age Please")

'If mName1 > mName2 Then
'   MsgBox "John is older than Kerry"
'Else
'    MsgBox "Kerry is older than John"
'End IF

If mName1 > mName2 or mName1 < mName2 Then
   MsgBox "John is older than Kerry"
    Else
    MsgBox "Kerry is older than John"
    End IF