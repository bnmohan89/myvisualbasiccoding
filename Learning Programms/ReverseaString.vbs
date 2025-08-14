'---------------------------------------
'Reverse a String with strReverse Function

   ' Dim str, var
   ' str = "OM NAMA SHIVAYA"
   ' var = StrReverse(Str)
   ' WScript.Echo var


'Reverse a String without strReverse Function

    Dim str, Length, i, val, Reverse
    str = "OM NAMA SHIVAYA"
    Length = Len(Str)
    'Wscript.Echo  


    For i = Length to 1 Step -1
        WScript.Echo i
        val = Mid(str,i,1)
        WScript.Echo val
        rev = rev&val
        WScript.Echo rev
    Next

    'MsgBox rev
    