'===========================================================================
' Assigment on Array
'===========================================================================
'Error Handling
Option Explicit

'Declare a Varailable
Dim arrNames
Dim arrNamePrnt

'Passing a value to array index
arrNames = Array("Staton Lugo","Corrinne Tunison","Wilber Pyne","Ava Enos","Marina Forst","Yetta Wile","Kum Gorrell","Omar Gonser","Blanch Kluth","Joanne Pfaff")

arrNamePrnt = arrNames

'MsgBox arrNamePrnt(0)
'WScript.Echo arrNamePrnt(0)

MsgBox arrNamePrnt(0) & " met " & arrNamePrnt(5) &  " in washington DC. "
WScript.Echo (arrNamePrnt(4) & " celebrated her birthday with " & arrNamePrnt(2))