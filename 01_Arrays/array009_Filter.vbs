'==========================================================================================================
'Array - Filter
'To Filter the data , extract all items which is having certain strint
'0 means binary comparision
'1 textual comperation
'==========================================================================================================
'Error Handling
Option Explicit

'Variables Decleration
Const SITE_TILE =  "www.Mohanscodingworld.com"
Dim arrManimals
Dim arrMnumbers
Dim matched, not_matched, items

'Filter (Filter(inutsstrings,values{,include[,compare]]))
'Compare:
'   0 = vbBinaryCompare (Binary Comparison)
'   1 = vbTextCompare (TextualComparision)

'Passing a value to Array
arrManimals = Array("Lion","Elephant","Snake","Chicken","Tiger")
arrMnumbers = Array("100", "111", "222", "200", "255")

'mathed condition

matched = Filter(arrManimals, "a" , True, 1)

MsgBox matched(0)
MsgBox matched(1)

'not_matched condition
 
not_matched = Filter(arrManimals, "a" , False, 1)

MsgBox not_matched(0)
MsgBox not_matched(1)
MsgBox not_matched(2)


'Validating the VbBinary Check Comapre

matched = Filter(arrMnumbers, "0", True, 1)
MsgBox matched(0)
MsgBox matched(1)
'MsgBox matched(2)

not_matched = Filter(arrMnumbers, "0", False,1)

WScript.Echo not_matched(0)
WScript.Echo not_matched(1)
WScript.Echo not_matched(2)
'WScript.Echo not_matched(3)
