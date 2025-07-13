'===========================================================================================
'Declare an array with 10 elements.
'Loop through the array to display a message for each array element with the element text.
'===========================================================================================
 
Option Explicit

Dim Arraynames, i

Arraynames = Array("Mohan", "Rahul", "Hemu", "Ajay", "Narasimha")

'MsgBox LBound(Arraynames)
'MsgBox UBound(Arraynames)

For i = LBound(Arraynames) to UBound(Arraynames)
    WScript.Echo " Array Element " &    i    & " : " &  Arraynames(i)
Next