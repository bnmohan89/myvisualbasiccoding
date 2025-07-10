'====================================================================================
' Select Case Assignment
'====================================================================================

'Error Handling
Option Explicit

'Constant Varaiable
Const InputBox_Title = "ABC Inc"

'Variable Decleration
Dim mName

mName = InputBox("Enter Your LastName", InputBox_Title, "You Last Name Please")

Select Case mName
    case "Thompson"
        MsgBox  " Hello "  & mName &  "you are part of ABC Inc with position President"
    case "Rooter"
         MsgBox  " Hello "  & mName &  " you are part of ABC Inc with position Sr.Vice President "
    case "Cooper"
         MsgBox  " Hello "  & mName &  " you are part of ABC Inc with position Vice President "
    case "Parker"
         MsgBox  " Hello "  & mName &  " you are part of ABC Inc with position Manger President "
    case else
        MsgBox " Hello you are not part of ABC Inc company and not with any management team "
end select

