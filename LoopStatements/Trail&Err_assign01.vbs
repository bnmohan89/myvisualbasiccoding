Dim inputNumber, i, output

inputNumber = InputBox("Enter a number between 1 and 100:")

' Check if input is numeric
If IsNumeric(inputNumber) Then
    inputNumber = CInt(inputNumber)
    
    ' Check if input is within range
    If inputNumber < 1 Or inputNumber > 100 Then
        MsgBox "The number is outside the valid range (1 to 100). Exiting."
        WScript.Quit
    Else
        ' Generate numbers divisible by 10
        output = ""
        For i = 1 To inputNumber
            If i Mod 10 = 0 Then
                output = output & i & vbCrLf
            End If
        Next
        
        If output = "" Then
            MsgBox "There are no numbers divisible by 10 in this range."
        Else
            MsgBox "Numbers divisible by 10 from 1 to " & inputNumber & ":" & vbCrLf & vbCrLf & output
            WScript.Echo "Numbers divisible by 10 from 1 to " & inputNumber & ":" & vbCrLf & vbCrLf & output
        End If
    End If
Else
    MsgBox "Invalid input. Please enter a numeric value. Exiting."
    WScript.Quit
End If