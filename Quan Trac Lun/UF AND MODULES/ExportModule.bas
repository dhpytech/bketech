Attribute VB_Name = "ExportModule"
Option Explicit
Public Sub range2Image(dhRange As Range, dhFileName As String, Optional ByVal DhPath As String = "")
Dim lWidth!, lHeight!, imageName$
Dim oChrtO As Object
If dhCheckFolderPath(DhPath) = False Then
  DhPath = dhRoot & "\Caches"
Else
End If
  imageName = DhPath & "\" & dhFileName & ".jpg"
  dhRange.CopyPicture xlScreen, xlPicture
  lWidth = dhRange.Width
  lHeight = dhRange.Height
  Set oChrtO = wsCache.ChartObjects.Add(Left:=0, Top:=0, Width:=lWidth, Height:=lHeight)
  oChrtO.Activate
  With oChrtO.Chart
   .Paste (13)
   .Export Filename:=imageName, Filtername:="JPG"
  End With
  oChrtO.Delete
End Sub

