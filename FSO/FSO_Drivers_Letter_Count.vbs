'-----------------------------------------------------------------------------------------------------
'Object.Drives Property - Object is always FileSystemObject & Returns a Collection with drive objects
'Object.DriveLetter Property - Object is always a Drive Object & Returns the drive letter
'Object.Count Property - Object is always a collection & Returns total items in a collection
'-----------------------------------------------------------------------------------------------------

Option Explicit

Dim oFSO, CDrives, oDrive
Dim ListofDrives

set oFSO = CreateObject("Scripting.FileSystemObject")

Set CDrives = oFSO.Drives
 
For Each oDrive in CDrives

    MsgBox " Drive Letter : " &oDrive.DriveLetter, 0 , "Drive on Your Computer:                "
    ListofDrives = ListofDrives & "   " & oDrive.DriveLetter

Next

    MsgBox " Number of Drivers : " & CDrives.Count, 0, "Drives on Your Computer:"
    MsgBox " drive Letter  : " & ListofDrives , 0, "Drives on Your Computer:"