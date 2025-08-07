'-----------------------------------------------------------------
'We will explore : CreateObject("Scripting.FileSystemObject")
'-----------------------------------------------------------------

Option Explicit

Dim oFSO

set oFSO = CreateObject("Scripting.FileSystemObject")

'createfolder - method
oFSO.CreateFolder("C:\Users\MOHAN\UFT Learning\Drive1")