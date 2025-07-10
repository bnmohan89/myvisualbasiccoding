'==========================================================================================================
'Array - 2dimensional arrays
'Column Index, Row Index(Ci,Ri)
'vbCrLf - Print the value in next line
'Preserve - Techically preserve the data
'==========================================================================================================
' Error Handling
Option Explicit

' Declaring a Variable and it is 2 dimensional array
Dim PhoneBook(2,3) ' 3 Columns 4 Rows & Here the size is defined

' Passing a value to array
PhoneBook(0,0) = "Mohan"
PhoneBook(1,0) = "Chennai"
PhoneBook(2,0) = "+91 9791346001"

PhoneBook(0,1) = "Rahul"
PhoneBook(1,1) = "Banglore"
PhoneBook(2,1) = "+91 7785224510"

PhoneBook(0,2) = "Yuva"
PhoneBook(1,2) = "Hydrbad"
PhoneBook(2,2) = "+91 7800520100"

PhoneBook(0,3) = "Srikant"
PhoneBook(1,3) = "Hydrbad"
PhoneBook(2,3) = "+91 7800520500"

MsgBox PhoneBook(0,0) & " lives in " & PhoneBook(1,0) & " . " & vbCrLf & " Contact No : " & PhoneBook(2,0)
MsgBox PhoneBook(0,1) & " lives in " & PhoneBook(1,1) & " . " & vbCrLf & " Contact No : " & PhoneBook(2,1)
MsgBox PhoneBook(0,2) & " lives in " & PhoneBook(1,2) & " . " & vbCrLf & " Contact No : " & PhoneBook(2,2)
MsgBox PhoneBook(0,3) & " lives in " & PhoneBook(1,3) & " . " & vbCrLf & " Contact No : " & PhoneBook(2,3)