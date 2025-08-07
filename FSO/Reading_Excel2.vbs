'------------------------------------------------------------------
' We will Explore:
'-[Excel.Application] --> [Workbook] --> [Sheets] --> [Rows, Columns]
'-Reading Excel Document (Specific Sheet & Cell)
' Using **sheet object**
' -Sheets return a collection of sheets ina workbook
' -To access a specific shet, use (1) sheet name or (b) index number
'           1 for 1st sheet, 2 for 2nd sheet, etc
' Sheets property retuens a collection of sheets
' Sheets collection includes all the sheets (chart sheets and worksheets)
'------------------------------------------------------------------
Option Explicit

Dim ObjExcel
Dim ObjWorkbook
Dim ObjWorkSheet
Dim ObjSheet ' <----------
Dim StrDirectory
Dim ExcelDocName
Dim SheetName
Dim name, age, email

StrDirectory = "C:\Users\MOHAN\UFT Learning"
ExcelDocName = "ReadingData1"
SheetName = "Data1"


' Creating a Object Excel
set ObjExcel =  CreateObject("Excel.Application")

ObjExcel.Visible = "True"


' Opening the worksheet
'Reading data from cells (using sheets property of WorkBook Object)

set ObjWorkbook = ObjExcel.workbooks.open(StrDirectory & "\" & ExcelDocName)

'we are binding to specific sheet based on the same
set ObjSheet = ObjWorkbook.sheets(SheetName)

'Reading the data from excel
name = ObjSheet.Range("A2").Value
    MsgBox "Name: "& name, 0, "Reading Excel Document..."

 age  = ObjSheet.Range("B2").Value
 MsgBox "age: " & age, 0, "Reading Excel Document..."

 email = ObjSheet.Range("C2").value
 MSGBox " email id is " & email , 0 , "Reading value from excel..."


  'Colse & Quite
    ObjWorkbook.Close
    ObjExcel.Quit

    'Release Memory
    Set ObjExcel = Nothing
    Set ObjWorkbook = Nothing
 






