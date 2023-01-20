call errorHandling
sub errorHandling() 

    a = 0
    b = 1
    On Error resume next
    c = b/a

    WScript.Echo Err.Number

    If Err.Number <> 0 Then
        'error handling:
        WScript.Echo Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
        Err.Clear
    End If
end sub
