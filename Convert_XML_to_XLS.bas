Sub Convert_XML_to_XLS()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim FolderPath As String, path As String, count As Integer
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim cutting As String
Dim i As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim pos3 As Integer
'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder("C:\Sem-o_Archive\XML")
i = 1
'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
    Workbooks.Add
    ActiveWorkbook.XmlImport URL:="C:\Sem-o_Archive\XML\" & objFile.name, ImportMap:=Nothing, _
                         Overwrite:=True, Destination:=range("$A$1")
    pos1 = InStr(objFile.name, "Metered")
    pos2 = InStr(objFile.name, "Dispatch")
    pos3 = InStr(objFile.name, "Actual")
    If pos1 > 0 Then
        cutting = Left(objFile.name, 36)
    ElseIf pos2 > 0 Then
        cutting = Left(objFile.name, 27)
    ElseIf pos3 > 0 Then
        cutting = Left(objFile.name, 29)
    End If
    ActiveWorkbook.SaveAs "C:\Sem-o_Archive\Source Files\" & cutting & "xls"
    ActiveWorkbook.Close True
    'print file path
    'Cells(i + 1, 2) = objFile.path
    i = i + 1
Next objFile

''ActiveWorkbook.XmlImport URL:="C:\Sem-o_Archive\Source Files\Actual Availability_05_01_15.xml", ImportMap:=Nothing, _
                         Overwrite:=True, Destination:=range("$A$1")

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub