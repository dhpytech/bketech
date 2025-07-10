Attribute VB_Name = "A_DataFunction"
'<-- Lay Duong Dan Folder App -->
Function DataPath(Optional SubPath As String = "") As String
If SubPath = "" Then
  DataPath = ThisWorkbook.Path & "\Data"
Else
  DataPath = ThisWorkbook.Path & "\Data\" & SubPath
End If
End Function

'<-- Dat Ten Cho Sheet Du Lieu-->
Function wsDSGoiThau() As Worksheet
  Set wsDSGoiThau = Sheet1
End Function

Function wsDSMatCat() As Worksheet
  Set wsDSMatCat = Sheet2
End Function

'<-- Mang Danh Sach Goi Thau -->
Function DSGoiThau(Header As Boolean, Optional NumCol As Long = 5) As Variant
Dim SRow As Long
  SRow = IIf(Header, 1, 2)
  DSGoiThau = wsDSGoiThau.Range(wsDSGoiThau.Cells(SRow, 1), wsDSGoiThau.Cells(getERow(wsDSGoiThau, "A"), NumCol)).Value2
End Function

'<-- Mang Danh Sach Goi Thau Theo Trang Thai -->
Function DSGTByOpt(Header As Boolean, Optional Opt As String = "All", Optional NumCol As Integer = 5) As Variant
Dim ArrData As Variant, SRow As Long, Status As String, MaGoiThau As String
Dim DictData As Object
Set DictData = CreateObject("Scripting.Dictionary")
SRow = IIf(Header, 1, 2)
ArrData = wsDSGoiThau.Range(wsDSGoiThau.Cells(SRow, 1), wsDSGoiThau.Cells(getERow(wsDSGoiThau, "A"), NumCol)).Value2


If Opt = "All" Then
  DSGTByOpt = DSGoiThau(False, 1)
Else
  For NumLp = LBound(ArrData, 1) To UBound(ArrData, 1)
    Status = ArrData(NumLp, 3)
    MaGoiThau = ArrData(NumLp, 1)
    
    If Status = Opt And Not DictData.exists(MaGoiThau) Then
      DictData.Add Key:=MaGoiThau, Item:=""
    End If
  Next NumLp
  DSGTByOpt = DictData.keys
End If
Set DictData = Nothing
End Function

'<-- Mang Danh Sach MCN -->
Function DSMatCat(Header As Boolean, Optional NumCol As Long = 5) As Variant
Dim SRow As Long
  SRow = IIf(Header, 1, 2)
  DSMatCat = wsDSMatCat.Range(wsDSMatCat.Cells(SRow, 1), wsDSMatCat.Cells(getERow(wsDSMatCat, "A"), NumCol)).Value2
End Function

'<-- Mang Danh Sach MCN Theo Goi Thau -->
Function DSMatCatByOpt(Header As Boolean, Optional Opt As String = "All", Optional NumCol As Integer = 5) As Variant
Dim ArrData As Variant, SRow As Long, Status As String, MaMatCat As String
Dim DictData As Object
Set DictData = CreateObject("Scripting.Dictionary")
SRow = IIf(Header, 1, 2)
ArrData = wsDSMatCat.Range(wsDSMatCat.Cells(SRow, 1), wsDSMatCat.Cells(getERow(wsDSMatCat, "A"), NumCol)).Value2

If Opt = "All" Then
  DSMatCatByOpt = DSMatCat(False, 1)
Else
  For NumLp = LBound(ArrData, 1) To UBound(ArrData, 1)
    TenGT = ArrData(NumLp, 3)
    MaMatCat = ArrData(NumLp, 1)
    TenMatCat = ArrData(NumLp, 2)
    
    If TenGT = Opt And Not DictData.exists(TenMatCat) Then
      DictData.Add Key:=TenMatCat, Item:=""
    End If
  Next NumLp
  DSMatCatByOpt = DictData.keys
End If
Set DictData = Nothing
End Function


