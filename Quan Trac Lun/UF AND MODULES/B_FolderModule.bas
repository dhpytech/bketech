Attribute VB_Name = "B_FolderModule"
Option Explicit
'---> FOLDER SUB AND FUNCTION <---
'<-- Lay Duong Dan Folder -->
Function GetFolderPath() As String
Dim FolderBox As Object
Set FolderBox = Application.FileDialog(msoFileDialogFolderPicker)
  With FolderBox
    .Title = "Ch" & ChrW(7885) & "n Folder L" & ChrW(432) & "u File"
    .AllowMultiSelect = False
    .InitialFileName = dhRoot
  End With
  
  If FolderBox.Show = True Then
    GetFolderPath = FolderBox.SelectedItems(1)
  Else
    GetFolderPath = ""
  End If
End Function
'<-- Kiem Tra Duong Dan Folder -->
Function CheckFolderPath(DhPath$) As Boolean
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
  If FSO.FolderExists(DhPath) Then
    CheckFolderPath = True
  Else
    CheckFolderPath = False
  End If
Set FSO = Nothing
End Function
'<-- Tao Folder Moi -->
Sub CreateFolder(DhPath$)
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
  If Not CheckFolderPath(DhPath) Then
    FSO.CreateFolder (DhPath)
  End If
Set FSO = Nothing
End Sub
'<-- Tao Danh Sach SubFolder -->
Function GetListSubFolder(FPath As String)
Dim FSO As Object, Folder As Object, SubFolder As Object
Dim result(), NumF As Long

Set FSO = CreateObject("Scripting.FileSystemObject")
  If FSO.FolderExists(FPath) Then
    Set Folder = FSO.GetFolder(FPath)
    If Folder.SubFolders.Count = 0 Then
      GetListSubFolder = Array()
      Exit Function
    End If
    ReDim result(0 To Folder.SubFolders.Count - 1)
    NumF = 0
    For Each SubFolder In Folder.SubFolders
      result(NumF) = SubFolder.Name
      NumF = NumF + 1
    Next
    GetListSubFolder = result
  Else
    GetListSubFolder = Array()
  End If
End Function


