'==========================================================================================================
'Array - Erase
'==========================================================================================================

'Error Handlig
Option Explicit

'Variablis
Dim arrMfriends
arrMfriends = Array("Rahul","Yuva""Hemu","Karthik")

'Erase ArrayValidation

MsgBox "Checking Poing 1 ===>"
WScript.Echo arrMfriends(1)

MsgBox "Checking Point 2 ===>"
Erase arrMfriends
WScript.Echo arrMfriends(1)