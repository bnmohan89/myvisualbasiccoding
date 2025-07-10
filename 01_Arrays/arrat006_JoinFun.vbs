'==========================================================================================================
'Array - Join Function
' Join a string 
'
'==========================================================================================================

'Error Handling
Option Explicit

'Delaring the Varaiable
Dim Websitrurl(2)
Dim WebSiteurlPrint
DIm OutValue


'Index value is 2 and maxminum leght is 3 and passing as value to arrat index
Websitrurl(0) = "www"
Websitrurl(1) = "Mohansworld"
Websitrurl(2) = "in"

WebSiteurlPrint = Join(Websitrurl, ".")
WScript.Echo (WebSiteurlPrint)

OutValue = InputBox(WebSiteurlPrint)
MsgBox (WebSiteurlPrint)

