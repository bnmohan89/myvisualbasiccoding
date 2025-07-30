'============================================================================================
' Sub Procedure --- Does not return a vlaue
' FUnction Procedure -- Can return a value
 '============================================================================================


 Option Explicit

 'variables decleration
 Dim num1, num2, total, sum

 num1 = 10
 num2 = 10

 total = add(num1, num2)

    WScript.Echo " Hello the sum of the two number is : " & total

    'A function procedure -- can return a value
    Function Add(num1,num2)
        sum  = num1 + num2
         Add = sum
    End Function