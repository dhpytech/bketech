Attribute VB_Name = "SystemControl"
Declare PtrSafe Function GetSystemMetrics32 Lib "user32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Public Function widthRatio(Optional WScreen As Long = 1163, Optional TwoScreen As Boolean = False) As Integer
Dim excelWidth As Integer
If Not TwoScreen Then
  excelWidth = ThisWorkbook.Application.Width
Else
  excelWidth = WScreen
End If
  widthRatio = Int(excelWidth / 1163 * 100)
End Function
Public Function heightRatio(Optional HScreen As Long = 623, Optional TwoScreen As Boolean = False) As Integer
Dim excelHeight As Integer
If Not TwoScreen Then
  excelHeight = ThisWorkbook.Application.Height
Else
  excelHeight = HScreen
End If
  heightRatio = Int(excelHeight / 623 * 100)
End Function
Public Function ShowScreenXDimensions()
   Dim X As Long
   Dim Y As Long
   ShowScreenXDimensions = GetSystemMetrics32(SM_CXSCREEN)
End Function
Public Function ShowScreenYDimensions()
   Dim X As Long
   Dim Y As Long
   ShowScreenYDimensions = GetSystemMetrics32(SM_CYSCREEN)
End Function
Public Function screenRate()
  dhRatio = ShowScreenXDimensions() / ShowScreenYDimensions() * 9
  screenRate = CStr(dhRatio) & ":9"
End Function
Public Function checkPathExists(Path) As Boolean
Dim X As String
  On Error Resume Next
  X = GetAttr(Path) And 0
  If Err = 0 Then checkPathExists = True Else checkPathExists = False
End Function


Public Function dhRoot() As String
  dhRoot = ThisWorkbook.Path
End Function

Sub ClearMemory(dhUF As UserForm, Values As Variant)
Dim numLoop&
  For numLoop = LBound(Values, 1) To UBound(Values, 1)
    dhUF.Controls(Values(numLoop)).Picture = Nothing
  Next numLoop
End Sub
