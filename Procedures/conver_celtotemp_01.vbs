'============================================================================================
' Sub Procedure --- Does not return a vlaue
' FUnction Procedure -- Can return a value
 '============================================================================================

 Option Explicit

 'variable decleration
 Dim temp,Sum,num1,num2, total

num1 = 10
num2 = 10

 'inCelsius, tempInput

 MsgBox "All about to call ConvertTemp Sub Procedure"
    Call ConvertTemp


'A Sub procedure -- does not returns a calue
    sub  ConvertTemp
        temp = InputBox("Please enter the  temperature in degrees F." , 1)
        MsgBox "CovertTemp: About to call Celsius Function", 64
        MsgBox "The temperature is " & Round(Celsius(temp),3) & " degrees Celsius.", 64
        MsgBox "The sum of num1+ num2 is " & Add(num1, num2) , 64
    End sub


'A Function procedure  -- can return a value
    Function Celsius(fdegrees)
        MsgBox "Celsius: I am inside Celsius ", 64
        Celsius = (fdegrees - 32) * 5 / 9
    End Function


'Another Sunction Procedure -- can return a value
 Function Add(num1, num2)
   'num1 = InputBox("Enter the Value of number 1 ", 1)
    'num2 = InputBox("Enter the Value of number 2 ", 1)
    Sum = num1 + num2
    Add = Sum
   
   
 End Function
