Option Explicit
 
Sub RemNamedRanges()
     
    Dim nm              As name
     
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        nm.Delete
    Next
    On Error GoTo 0
     
End Sub