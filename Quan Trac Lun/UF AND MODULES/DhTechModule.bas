Attribute VB_Name = "DhTechModule"
Option Explicit
'<-- Get End Row -->
Public Function getERow(ws As Worksheet, Col As String) As Long
  getERow = ws.Range(Col & Rows.Count).End(xlUp).Row
End Function
'<-- Get Data Row -->
Public Function getDRow(ws As Worksheet, Col As String) As Long
  getDRow = ws.Range(Col & Rows.Count).End(xlUp).Row + 1
End Function
'<-- Get End Column -->
Public Function getECol(ws As Worksheet, RowNum As Long) As Long
  getECol = ws.Cells(RowNum, ws.Columns.Count).End(xlToLeft).Column
End Function
'<-- Thong Bao Thieu Thong Tin -->
Function MsgThieuTT() As String
  MsgThieuTT = "Vui L" & ChrW(242) & "ng " & ChrW(272) & "i" & ChrW(7873) & "n " & ChrW(272) & ChrW(7847) & "y " & ChrW(272) & ChrW(7911) & " Th" & ChrW(244) & "ng Tin!"
End Function
'<-- Convert Text To Number -->
Function Text2Num(Value As Variant) As Double
  If Not IsNumeric(Value) Then
    Text2Num = 0
  Else
    Text2Num = Value
  End If
End Function
'<-- Check Exists Array -->
Function CheckArrExists(Arr As Variant, TextValue As String, Optional ArrSize As Integer = 1, Optional NumCol As Long = 1)
Dim NumLp As Long, LArr As Long, UArr As Long, TenMa As String
CheckArrExists = False
If Not IsNull(Arr) Then
  If ArrSize = 1 Then
    LArr = LBound(Arr)
    UArr = UBound(Arr)
  Else
    LArr = LBound(Arr, 1)
    UArr = UBound(Arr, 1)
  End If
  
  For NumLp = LArr To UArr
    If ArrSize = 1 Then
      TenMa = Arr(NumLp)
    Else
      TenMa = Arr(NumLp, NumCol)
    End If
    
    If TextValue = TenMa Then
      CheckArrExists = True
      Exit For
    End If
  Next NumLp
End If
End Function
'<-- Check Empty -->
Function CheckEmpty(Arr As Variant)
Dim NumLp As Long
CheckEmpty = False
  For NumLp = LBound(Arr) To UBound(Arr)
    If Arr(NumLp) = "" Then
      CheckEmpty = True
      Exit For
    End If
  Next NumLp
End Function
'<-- Get Each Column Of 2D Array -->

Function GetSubArr(ByVal Arr2D As Variant, ByVal SCol As Long, ByVal ECol As Long, Optional ByVal ArrBase As Long = 0) As Variant
Dim numRows As Long, numCols As Long
Dim NumLp As Long, NumCol As Long, NewBound As Long
Dim result() As Variant
If IsNull(Arr2D) Then
  GetSubArr = Array()
Else
NewBound = IIf(ArrBase = 0, UBound(Arr2D, 1), UBound(Arr2D, 1) - 1)
ReDim result(0 To NewBound, 0 To ECol - SCol)
  For NumLp = 0 To NewBound
    For NumCol = SCol To ECol
      result(NumLp, NumCol - SCol) = Arr2D(NumLp + ArrBase, NumCol)
    Next NumCol
  Next NumLp

GetSubArr = result
End If
End Function
'<-- Tao Ma Hash 256 -->
Function HashSHA256(ByVal inputText As String) As String
  HashSHA256 = inputText
End Function

Function CEm2None(dhValue As Variant, Optional ByVal dhReverse As Boolean = False)
If dhReverse = False Then
  If dhValue = "" Then
    CEm2None = "None"
  Else
    CEm2None = dhValue
  End If
Else
  If dhValue = "None" Then
    CEm2None = ""
  Else
    CEm2None = dhValue
  End If
End If
End Function

Function VNum(dhNum$) As String
Dim PhanNguyen$, PhanThapPhan$
If InStr(1, dhNum, ".", vbTextCompare) > 0 Then
  PhanNguyen = Replace(Split(dhNum, ".", , vbTextCompare)(0), ",", ".", , , vbTextCompare)
  PhanThapPhan = Split(dhNum, ".", , vbTextCompare)(1)
  VNum = Join(Array(PhanNguyen, PhanThapPhan), ",")
Else
  VNum = Replace(dhNum, ",", ".", , , vbTextCompare)
End If
End Function

Function NoneDict(KeyArr As Variant) As Object
Dim DhDict As Object, NumLp As Long
Set DhDict = CreateObject("Scripting.Dictionary")
  For NumLp = LBound(KeyArr) To UBound(KeyArr)
    DhDict.Add Key:=KeyArr(NumLp), Item:="None"
  Next NumLp
Set NoneDict = DhDict
Set DhDict = Nothing
End Function

Function Dict2Str(DhDict As Object) As String
Dim DhKey As Variant, DhItem As Variant
Dim DhStr$, dhText$
DhStr = ""
  For Each DhKey In DhDict.keys
    dhText = DhKey & "|:" & DhDict(DhKey)
    If DhStr = "" Then
      DhStr = dhText
    Else
      DhStr = DhStr & "<|>" & dhText
    End If
  Next DhKey
Dict2Str = DhStr
End Function

Function Str2Dict(DhStr As String) As Object
Dim DhDict As Object
Set DhDict = CreateObject("Scripting.Dictionary")
Dim ArrData As Variant, DhKey As String, DhItem As String, NumLp As Long

ArrData = Split(DhStr, "<|>", , vbTextCompare)
  For NumLp = LBound(ArrData, 1) To UBound(ArrData, 1)
    DhKey = Split(ArrData(NumLp), "|:", , vbTextCompare)(0)
    DhItem = Split(ArrData(NumLp), "|:", , vbTextCompare)(1)
    DhDict.Add Key:=DhKey, Item:=DhItem
  Next NumLp
Set Str2Dict = DhDict
Set DhDict = Nothing
End Function

Function NumRowTB(ArrData As Variant, Text As String) As Long
Dim NumLp As Long
NumRowTB = -1
For NumLp = LBound(ArrData) To UBound(ArrData)
  If ArrData(NumLp, 1) = Text Then
    NumRowTB = NumLp
    Exit For
  End If
Next NumLp
End Function
