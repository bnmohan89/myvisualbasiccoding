'=================================================================
' Loop: Do
'=================================================================
'Exception Handling

Option Explicit

'Varaiabl Declaration
Dim count

count  = 1


Do
    MsgBox " I'm Inside the Loop : " & count 
    count = count + 1
    If count = 10 Then
       ' WScript.Quit ' it will end the scipt in line 18
        Exit Do
    End If
Loop

    WScript.Echo " Loop is done & I'm outside of the loop "