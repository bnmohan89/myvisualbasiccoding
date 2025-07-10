'=================================================================
' Loop: Exiting
'=================================================================

Option Explicit

'Variab
Dim a

a = 1

Do While a < 25 ' Exit when False
    MsgBox " a = " & a
    WScript.Echo a & " Welcome to mohancodingworld.com "
        if a = 5 Then
            WScript.Echo " The value of a is equal to : " & a
            WScript.Echo " Exiting the Loop....."
            Exit Do
        end if
         a = a + 1
            MsgBox " The value of a is not equal to 5 : True ", 1 , "PASS"

    Loop

        WScript.Echo " Bye Bye "