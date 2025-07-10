'=====================================================================================
'Select Case with Built in function
'=====================================================================================
'Error Handling
Option Explicit

Dim X
x = Weekday(date) ' Built in Function



Select Case X
    case vbSunday
        MsgBox "Today is Sun"
    case vbMonday
        MsgBox "Today is Mon"
    case vbTuesday
        MsgBox "Today is Tues"
        WScript.Echo CStr(vbTuesday)
    case vbWednesday
        MsgBox "Today is Wed"
    case vbThursday
        MsgBox "Today is Thur"
    case vbFriday
        MsgBox "Today is Fri"
    case vbSaturday
        MsgBox "Today is Sat"
    End select