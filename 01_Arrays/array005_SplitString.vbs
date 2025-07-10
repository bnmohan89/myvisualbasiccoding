'==========================================================================================================
'Array - Split Function
' Spilt a string and convert that into Array & also join a array and exract string out of it
'
'==========================================================================================================
'Error Handling
Option Explicit

'Constant Variable
Const WebTitle = "www.MohanWorld.Com"

'Declare Varibale
Dim Msplit1
Dim Msplit2

'Declare a string to varibale
Msplit2 = "www.Mohanworld1.Compare"


'Split Function
Msplit1 = Split(WebTitle, ".")
Msplit2 = Split(Msplit2, ".")

WScript.Echo Msplit1(0)
WScript.Echo Msplit1(1)
WScript.Echo Msplit1(2)
'WScript.Echo Msplit1(3)

WScript.Echo Msplit2(0)
WScript.Echo Msplit2(1)
WScript.Echo Msplit2(2)