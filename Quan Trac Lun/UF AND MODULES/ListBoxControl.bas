Attribute VB_Name = "ListBoxControl"
Sub UpdateLB(dhLB As Object, ArrData As Variant)
  If IsEmpty(ArrData) Then
    dhLB.Clear
    Exit Sub
  End If
  dhLB.List = ArrData
  If LBound(ArrData, 2) = 1 Then
    dhLB.ColumnCount = UBound(ArrData, 2)
  Else
    dhLB.ColumnCount = UBound(ArrData, 2) + 1
  End If
End Sub
Function LBIndex(LB As Object, TimKiem As String, Optional NumCol As Long = 0)
Dim NumLp As Long, DongThongBao As String
Dim ArrData As Variant
ArrData = LB.List
LBIndex = -1
If LB.ListCount > 0 Then
  For NumLp = LBound(ArrData, 1) To UBound(ArrData, 1)
    If ArrData(NumLp, NumCol) = TimKiem Then
      LBIndex = NumLp
      Exit For
    End If
  Next NumLp
End If
End Function

Sub UpdateSearchLB(dhListBox As Object, dataArr As Variant, SearchText As String, Optional ByVal num As Integer = 0)
Dim ArrData, NumRow&, NumCol%
Dim arrSearch()
If IsEmpty(dataArr) Then
  dhListBox.Clear
  Exit Sub
End If
  
dhListBox.ColumnCount = UBound(dataArr, 2)
ArrData = dataArr
  ReDim arrSearch(num To UBound(ArrData, 1), num To UBound(ArrData, 2))
  numArr = num
  For NumRow = LBound(ArrData, 1) To UBound(ArrData, 1)
    For NumCol = LBound(ArrData, 2) To UBound(ArrData, 2)
      If LCase(ArrData(NumRow, NumCol)) Like "*" & LCase(SearchText) & "*" Then
        For numSearch = LBound(ArrData, 2) To UBound(ArrData, 2)
          arrSearch(numArr, numSearch) = ArrData(NumRow, numSearch)
        Next numSearch
        numArr = numArr + 1
        Exit For
      Else
      End If
    Next NumCol
  Next NumRow
  dhListBox.List = arrSearch
End Sub
Function totalListBox(dhListBox As Object, sumCol As Integer)
Dim ArrData As Variant, total#
  ArrData = dhListBox.List
  total = 0
  For num = LBound(ArrData, 1) To UBound(ArrData, 1)
    total = total + CSng(CnullToZero(ArrData(num, sumCol)))
  Next num
  totalListBox = total
  Erase ArrData
End Function

Sub UpdateLBByCol(dhLB As Object, ArrData As Variant, SearchText As String, ColIndex As Integer, Optional ByVal numType As Integer = 1, Optional Title As Boolean = False)
Dim NumRow&, numS As Long, numKey As Long
Dim DictSearch As Object
Dim ArrItems As Variant, Item As Variant

Dim arrResult(), ArrTemp(), arrCaches As Variant
If IsEmpty(ArrData) Then
  dhListBox.Clear
  Exit Sub
Else
  Set DictSearch = CreateObject("Scripting.Dictionary")
End If

numKey = 1

For NumRow = LBound(ArrData, 1) To UBound(ArrData, 1)
If Title = False Then
  SS = (LCase(ArrData(NumRow, ColIndex)) Like "*" & LCase(SearchText) & "*")
Else
  SS = (LCase(ArrData(NumRow, ColIndex)) Like "*" & LCase(SearchText) & "*" Or NumRow = LBound(ArrData, 1))
End If
  If SS Then
    ReDim ArrTemp(numType To UBound(ArrData, 2))
    
    For numS = LBound(ArrData, 2) To UBound(ArrData, 2)
      ArrTemp(numS) = ArrData(NumRow, numS)
    Next numS
    DictSearch.Add Key:=numKey, Item:=Join(ArrTemp, "<||>")
    numKey = numKey + 1
  End If
Next NumRow

If DictSearch.Count > 0 Then
  ArrItems = DictSearch.items
  ReDim arrResult(0 To UBound(ArrItems), 0 To UBound(ArrData, 2))
  NumRow = 0
  For Each Item In ArrItems
    arrCaches = Split(Item, "<||>", , vbTextCompare)
    For NumCol = LBound(arrCaches) To UBound(arrCaches)
      arrResult(NumRow, NumCol) = arrCaches(NumCol)
    Next NumCol
    NumRow = NumRow + 1
  Next Item
  dhLB.List = arrResult
End If
Set DictSearch = Nothing
End Sub

Sub UpdateLBByDateCol(dhLB As Object, ArrData As Variant, SearchText As String, ColIndex As Integer, Compare As Boolean, Optional ByVal numType As Integer = 1, Optional Title As Boolean = False)
Dim NumRow&, numS As Long, numKey As Long
Dim DictSearch As Object
Dim ArrItems As Variant, Item As Variant

Dim arrResult(), ArrTemp(), arrCaches As Variant
If IsEmpty(ArrData) Then
  dhListBox.Clear
  Exit Sub
Else
  Set DictSearch = CreateObject("Scripting.Dictionary")
End If

numKey = 1
For NumRow = LBound(ArrData, 1) To UBound(ArrData, 1)
  If Compare Then
    If Title Then
      SS = (dhDate(CStr(ArrData(NumRow, ColIndex))) <= Date Or NumRow = LBound(ArrData, 1))
    Else
      SS = (dhDate(CStr(ArrData(NumRow, ColIndex))) <= Date)
    End If
  Else
    If Title Then
      SS = (dhDate(CStr(ArrData(NumRow, ColIndex))) > Date Or NumRow = LBound(ArrData, 1))
    Else
      SS = (dhDate(CStr(ArrData(NumRow, ColIndex))) > Date)
    End If
  End If
  If SS Then
    ReDim ArrTemp(numType To UBound(ArrData, 2))
    For numS = LBound(ArrData, 2) To UBound(ArrData, 2)
      ArrTemp(numS) = ArrData(NumRow, numS)
    Next numS
    DictSearch.Add Key:=numKey, Item:=Join(ArrTemp, "<||>")
    numKey = numKey + 1
  End If
Next NumRow

If DictSearch.Count > 0 Then
  ArrItems = DictSearch.items
  ReDim arrResult(0 To UBound(ArrItems), 0 To UBound(ArrData, 2))
  NumRow = 0
  For Each Item In ArrItems
    arrCaches = Split(Item, "<||>", , vbTextCompare)
    For NumCol = LBound(arrCaches) To UBound(arrCaches)
      arrResult(NumRow, NumCol) = arrCaches(NumCol)
    Next NumCol
    NumRow = NumRow + 1
  Next Item
  dhLB.List = arrResult
End If
Set DictSearch = Nothing
End Sub
