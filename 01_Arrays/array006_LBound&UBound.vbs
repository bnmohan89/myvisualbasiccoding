'==========================================================================================================
'Lower Index of Array & Upper Index of Array
'LBound Function - Return Lower Index of Array
'UBound - Return Upper Index of Array
'==========================================================================================================
'Error Handling
Option Explicit

'Declare Variable
Const TitleName = "LBound&UBound"
Dim LowerIndex, UpperIndex
Dim Marrynum(4)


Marrynum(0) = 100
Marrynum(1) = 100
Marrynum(2) = 100
Marrynum(3) = 100
Marrynum(4) = 100

LowerIndex = LBound(Marrynum)
UpperIndex = UBound(Marrynum)

'MSGBOX

MsgBox "Lower Index of Array is  " & LowerIndex, 64, TitleName
MsgBox "Upper Index of Array is  " & UpperIndex , 64, "L&UBound"