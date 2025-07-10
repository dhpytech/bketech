Attribute VB_Name = "SheetsControl"
Option Explicit
Sub HideAllSheet(dhBoolean As Boolean, numSheet As Integer)
Dim NumLp As Integer
For NumLp = 1 To numSheet
  With ThisWorkbook.Sheets(NumLp)
    If dhBoolean = True Then
      If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
    Else
      If .Visible = xlSheetVisible Then .Visible = xlVeryHidden
    End If
  End With
Next NumLp
End Sub

