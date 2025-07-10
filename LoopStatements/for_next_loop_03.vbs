'===============================================================
'For..Next..Loop with conditional If..Then..End IF statement
'Decreament the Index
'===============================================================

Option Explicit

'Variable Decleration
Dim arrNum, total, i


arrNum = Array(1,2,3,4,5,6,7,8,9,10)


'Using array concept in For loop
For i = Ubound(arrNum) to Lbound(arrNum) step -1
    total = total + arrNum(i)
        WScript.Echo "Printing the Total Value" & total
            if total > 10 Then
                WScript.Echo " Hurry The total is greate than 10!" & vbCrLf & "Current total value is " & total & " !  !"
            End If
                WScript.Echo " -----**--**--**----- "
Next
             WScript.Echo " The sum of all the elements in the array is :  " & total