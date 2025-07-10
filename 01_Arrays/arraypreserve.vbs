'Increasing the Lenght of array
'ReDim Preserve -  Save the data in a exisitng array
'                   When you change the sie of the last dimesnion

'=======================================================================================

'Error Handling
Option Explicit

'Declare Constant
Const TITLENAME = "FRIENDS"

'Declare an array
Dim arrFriends
arrFriends = Array("Rahul","Yuva","Hemu")

'MsgBox arrFriends(2),1,TITLENAME
ReDim Preserve arrFriends(3) 
arrFriends(3)= "Hemu"
MsgBox arrFriends(3),1,TITLENAME

ReDim Preserve arrFriends(4)
arrFriends(4) = "Karthik"

MsgBox arrFriends(4),68,TITLENAME
MsgBox arrFriends(0),68,TITLENAME