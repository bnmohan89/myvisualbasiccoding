'==========================================================================================================
'Lbound&Ubound in 2 dimenstional array
'==========================================================================================================
'Error Handling
Option Explicit

'Declare Variables
Dim Mphoneboox()

Redim Mphoneboox(2,20)

MsgBox LBound(Mphoneboox)
MsgBox UBound(Mphoneboox)

'Specifing the dimenion which we want to identify
'Check the 1st Dimension Column
MsgBox LBound(Mphoneboox, 1)
MsgBox UBound(Mphoneboox, 1)

'Chech the 2nd dimensnion Row
MsgBox LBound(Mphoneboox, 2)
MsgBox UBound(Mphoneboox, 2)
