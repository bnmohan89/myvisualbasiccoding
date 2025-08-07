'------------------------------------------------------------------
' We will Explore:
'-[Excel.Application] --> [Workbook] --> [Sheets] --> [Rows, Columns]
'-Reading Excel Document (Specific Sheet & Cell)
'------------------------------------------------------------------

Option Explicit

Dim ObjExcel
Dim ObjWorkbook
Dim ObjWorkSheet
Dim ObjSheet
Dim StrDirectory
Dim ExcelDocName
Dim SheetName
Dim name, age, email

StrDirectory = "C:\Users\MOHAN\UFT Learning"
ExcelDocName = "ReadingData1"
SheetName = "Data1"

'Create Excel Object
set ObjExcel = CreateObject("Excel.Application")


'Excel Document Visible Property
ObjExcel.visible = True

'Getting an workbook object (using specific file)
Set ObjWorkbook = ObjExcel.workbooks.open(StrDirectory & "\" & ExcelDocName)

'Reading data from cells (using sheets property of WorkBook Object)
    name = ObjWorkbook.sheets(SheetName).Range("A2").Value
    MsgBox "Name: "& name, 0, "Reading Excel Document..."

    age = ObjWorkbook.sheets(SheetName).Range("B2").Value
    MsgBox "age: "& age, 0, "Reading Excel Document..."

    email = ObjWorkbook.sheets(SheetName).Range("C2").Value
    MsgBox "email: "& email, 0, "Reading Excel Document..."

    'Colse & Quite
    ObjWorkbook.Close
    ObjExcel.Quit

    'Release Memory
    Set ObjExcel = Nothing
    Set ObjWorkbook = Nothing



