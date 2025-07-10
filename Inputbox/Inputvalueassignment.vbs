'Input value assigenment

' Variable & Constant
Dim StrName
Dim numAge
Dim StrCity
Const TITLE_ASS = " TITLE ASSIGMENT ":


' Getting the input from user & Adding little Customization:Default text

StrName = InputBox (" Enter your name : ", TITLE_ASS , " Your Name ")
numAge = InputBox (" Enter your age : ", TITLE_ASS)
StrCity = InputBox (" Enter your city : ", TITLE_ASS)

' Displaying the Result

'MsgBox (" Hello "  & StrName & " is " & numAge & " Years old and lives in " & StrCity )
MsgBox " Hello "  & StrName & " is " & numAge & " Years old and lives in " & StrCity , 1 , TITLE_ASS