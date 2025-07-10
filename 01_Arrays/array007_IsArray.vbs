'==========================================================================================================
'Array - IsArray
'==========================================================================================================

'Error Handling
Option Explicit

'Declaring a Variable
Dim MName
Dim MArrayNames(2)

'Passing a string to varaiable
MName = "King"

'Passing a values to arrays
MArrayNames(0) = "Vinuthna"
MArrayNames(1) = "Joshnika"
MArrayNames(2) = 10000000


'Checing array or not

MsgBox "Checking MArrayNames is array or not   : " & IsArray(MArrayNames), 1 , "IsArrayCheck Validation"
MsgBox "Checking MName is array or not     : " & IsArray(Mname), 1 , "IsArrayCheck Validation", 5000, 100000