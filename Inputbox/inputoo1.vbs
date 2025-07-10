'===============================================================================================
' Usinh Input  -> Getting input from the user
' CInt --> Checking input value is integer or not
' CInt --> CInt & Customizing the input box
'===============================================================================================

'Variable & Constant
Const SITE_TITLE = " >> Inputvalue.io "
Dim input1, input2, input3 
Dim sum 
Dim name


' Getting the input from user
 name = InputBox("Enter Your Name: ", SITE_TITLE)
 input1 = InputBox(" Enter the first value: " , SITE_TITLE)
 input2 = InputBox(" ENter the second value: ", SITE_TITLE, 5000, 5000)

 ' Adding little Customization:Default text
 input3 = InputBox(" Enter the Third Value: " , SITE_TITLE , "Enter Input HEre"  )
 
'Processing
sum = (input1) + input2 + input3
'sum = input1 + CIntinput2 + input3

' Displaying the Result
MsgBox sum

MsgBox " Hello " & name & " and sum of user inputs is " & sum & " ! ! ! " , 1 , SITE_TITLE
MsgBox " The Sum of two numbers is " & sum, 64, SITE_TITLE