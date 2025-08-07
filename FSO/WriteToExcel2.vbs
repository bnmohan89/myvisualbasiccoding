'------------------------------------------------------------------
' We will Explore:
'-[Excel.Application] --> [Workbook] --> [Sheets] --> [Rows, Columns,Cells]
'-Reading Excel Document (Specific Sheet & Cell)
' Using **sheet object**
' -Sheets return a collection of sheets ina workbook
' -To access a specific shet, use (1) sheet name or (b) index number
'           1 for 1st sheet, 2 for 2nd sheet, etc
' Sheets property retuens a collection of sheets
' Sheets collection includes all the sheets (chart sheets and worksheets)
'------------------------------------------------------------------
'' We will Explore:
'-[Excel.Application] --> [Workbook] --> [Sheets] --> [Rows, Columns,Cells]
'- Editing the Excel Document
'- We will use the sheet object here.
' - Example : Range("A1").value = "Rahul"
'-We need to get the desired cell, and then set the value.
'---------------------------------------------------------------------

Option Explicit

Dim ObjExcel
Dim ObjWorkbook
Dim ObjWorkSheet
Dim ObjSheet ' <----------

Dim StrDirectory
Dim ExcelDocName
Dim SheetName
Dim name, age, email

Dim UpdateName, x

StrDirectory = "C:\Users\MOHAN\UFT Learning"
ExcelDocName = "ReadingData2"
SheetName = "Data1"


' Creating a Object Excel
set ObjExcel =  CreateObject("Excel.Application")

ObjExcel.Visible = "True"


'Opening the worksheet
'Reading data from cells (using sheets property of WorkBook Object)

set ObjWorkbook = ObjExcel.workbooks.open(StrDirectory & "\" & ExcelDocName)

'we are binding to specific sheet based on the same
'set ObjSheet = ObjWorkbook.sheets(SheetName)

'Set the value to particular cell
 'ObjExcel.sheets(SheetName).Range("A3").value = "Mohan"
 'ObjExcel.sheets(SheetName).Range("B3").value = "35"
 'ObjExcel.sheets(SheetName).Range("C3").value = "mohan.n.bodavula@gmail.com"

 'ObjExcel.sheets(SheetName).Range("A4").value = "Rahul"
 'ObjExcel.sheets(SheetName).Range("B4").value = "40"
 'ObjExcel.sheets(SheetName).Range("C4").value = "RahulJain@gmail.com"

'ObjExcel.sheets(SheetName).Range("A5").value = "Vivek"
'ObjExcel.sheets(SheetName).Range("B5").value = "38"
'ObjExcel.sheets(SheetName).Range("C5").value = "Vivek@gmail.com"

'Update email address of Vivek
for x = 1 to 5 Step 1
    UpdateName = ObjExcel.sheets(SheetName).cells(x,1).value '--> Row,column
        If UpdateName = "Vivek" Then
            ObjExcel.sheets(SheetName).cells(X,3).Value = "VivekanandGanesan@gmail.com"
        End If
    Next

    MsgBox " Editing Complete      ", 0, "Status of Editing"


  'Colse & Quite
    ObjWorkbook.save
    ObjWorkbook.Close
    ObjExcel.Quit

    'Release Memory
    Set ObjExcel = Nothing
    Set ObjWorkbook = Nothing