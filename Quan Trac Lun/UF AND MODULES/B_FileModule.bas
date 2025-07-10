Attribute VB_Name = "B_FileModule"
'---> FILE SUB AND FUNCTION <---
'<-- Lay Duong Dan File -->
Public Function GetFilePath(MultiSelect As Boolean, Optional FileType As String = "Excel File", Optional FileExtension As String = "*.xls; *.xlsx; *.xlsm") As String
Dim FileBox As Object
Set FileBox = Application.FileDialog(msoFileDialogFilePicker)
  With FileBox
    .Title = "Select File"
    .AllowMultiSelect = MultiSelect
    .InitialFileName = "C:\"
    .Filters.Clear
    .Filters.Add FileType, FileExtension
    .Filters.Add "All", "*.*"
  End With
  
  If FileBox.Show = True Then
    GetFilePath = FileBox.SelectedItems(1)
  Else
    GetFilePath = ""
  End If
End Function
'<-- Copy File -->
Public Sub CreateCopyFile(FileGoc$, FileDich$)
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
  If FSO.FileExists(FileGoc) Then
    FSO.CopyFile FileGoc, FileDich, True
  End If
Set FSO = Nothing
End Sub
'<-- Kiem Tra Duong Dan File -->
Public Function CheckFileExists(DhPath$) As Boolean
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
  If FSO.FileExists(DhPath) Then
    CheckFileExists = True
  Else
    CheckFileExists = False
  End If
Set FSO = Nothing
End Function
'<-- Lay Danh Sach File XLSX -->
Function getFileXLSX(FPath As String) As Variant
Dim FSO As Object, Folder As Object, File As Object
Dim result(), NumF As Long
Set FSO = CreateObject("Scripting.FileSystemObject")
  If Not FSO.FolderExists(FPath) Then
    getFileXLSX = Array()
    Exit Function
  End If
Set Folder = FSO.GetFolder(FPath)
  NumF = 0
  For Each File In Folder.Files
    If LCase(FSO.GetExtensionName(File.Name)) = "xlsx" Then
      ReDim Preserve result(NumF)
      result(NumF) = FSO.GetBaseName(File.Name)
      NumF = NumF + 1
    End If
  Next File
  If NumF = 0 Then
    getFileXLSX = Array()
  Else
    getFileXLSX = result
  End If
Set File = Nothing
Set Folder = Nothing
Set FSO = Nothing
End Function

