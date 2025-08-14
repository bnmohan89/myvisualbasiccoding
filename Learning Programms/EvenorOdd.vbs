'Programme to find out the given Input number is Even or Odd

Dim inpNumber

inpNumber = InputBox("Enter the Number","Input Box","Plz Enter Your Number",10000,10000)


if  IsNumeric(inpNumber) Then
    if inpNumber mod 2=0 Then 
        WScript.Echo "This is Even Number"
    Else
        WScript.Echo "This is Odd Number"
    End If
Else
    
    WScript.Echo "Input Number is not Numeric"
End If
