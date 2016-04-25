Sub RemoveTableStyles()
    Dim oSh As Worksheet
    Dim Table As String
    Table = ActiveSheet.ListObjects(1)
    Set oSh = ActiveSheet
    'remove table or list style
    oSh.ListObjects(Table).Unlist
End Sub