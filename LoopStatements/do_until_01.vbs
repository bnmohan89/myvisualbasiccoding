'=================================================================
' Loop: Do..Until
'=================================================================
'Exception Handling

Option Explicit

'Varaiabl Declaration
Dim count

count = 1

Do Until count >= 10 ' Exit when True
    WScript.Echo " The value of count is : " & count
    count = count + 1
    WScript.Echo " After increament the value of count is : " & count
Loop
    WScript.Echo " Loop is done & outside of the loop "