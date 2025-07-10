'===============================================================
'Array - 2dim
'===============================================================
'Error Handling

Option Explicit

'Variable
Dim PhoneBook(2,4)  '3 Columns, 5 Rows
Dim colLowerIndex, ColHigherIndex, rowLowerIndex, rowHigherIndex
Dim ri, ci 

'Assigning values to the elements in Array
' PhoneBook(column, Row)

PhoneBook(0,0) = "Mohan"
PhoneBook(1,0) = "Chennai"
PhoneBook(2,0) = "111-111-001"

PhoneBook(0,1) = "Narasiha"
PhoneBook(1,1) = "Banglore"
PhoneBook(2,1) = "111-111-002"

PhoneBook(0,2) = "Rahul"
PhoneBook(1,2) = "Hyd"
PhoneBook(2,2) = "111-111-003"

PhoneBook(0,3) = "Yuva"
PhoneBook(1,3) = "Banglore"
PhoneBook(2,3) = "111-111-004"

PhoneBook(0,4) = "Hemu"
PhoneBook(1,4) = "Pune"
PhoneBook(2,4) = "111-111-005"

'Getting the counts for Processin
colLowerIndex = LBound(PhoneBook, 1) ' we can use 0 index value also
ColHigherIndex = UBound(PhoneBook, 1)

rowLowerIndex = LBound(PhoneBook, 2) ' we can use 0 index value also
rowHigherIndex = UBound(PhoneBook, 2)

'Using Nested For Loop

for ri = rowLowerIndex to rowHigherIndex
    for ci = colLowerIndex to ColHigherIndex
        WScript.Echo PhoneBook(ci,ri)
    Next
        WScript.Echo "--------------------"

Next

