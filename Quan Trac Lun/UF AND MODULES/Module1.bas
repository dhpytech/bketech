Attribute VB_Name = "Module1"
Function UIController(ArrName As Boolean)
If ArrName Then
  UIController = Array("BTNGoiThau", "BTNQuanTrac", "BTNMatCat", "BTNUser", "BTNLuu", "BTNCaiDat", "BTNThongKe", "BTNThoat", "Logo1", "Logo2")
Else
  UIController = Array("goithau.jpg", "quantrac.jpg", "matcat.jpg", "user.jpg", "luu.jpg", "caidat.jpg", "thongke.jpg", "thoat.jpg", "logo1.jpg", "logo2.jpg")
End If
End Function

Function UIBasic(ArrName As Boolean)
If ArrName Then
  UIBasic = Array("BTNHome")
Else
  UIBasic = Array("home.jpg")
End If
End Function

Sub HideAllUI()
  If ThisWorkbook.ReadOnly Then
    A2ControllerUF.LBLWarning.Visible = True
  Else
    A2ControllerUF.LBLWarning.Visible = False
  End If
  Call HideUI(A2ControllerUF, "", "BTNLuu", "BTNDSKho")
End Sub


Function QuickSort2D(ByVal Arr As Variant, ByVal k As Long) As Variant
  Dim sortedArr As Variant
  Dim first As Long, last As Long

  sortedArr = Arr
  first = LBound(sortedArr, 1)
  last = UBound(sortedArr, 1)
  Call QuickSort2D_Internal(sortedArr, first, last, k)

  QuickSort2D = sortedArr
End Function

Private Sub QuickSort2D_Internal(ByRef Arr As Variant, ByVal first As Long, ByVal last As Long, ByVal k As Long)
Dim i As Long, j As Long
Dim pivot As Variant
Dim temp As Variant
Dim colCount As Long
Dim t As Long

i = first
j = last
pivot = dhDate(CStr(Arr((first + last) \ 2, k)))
colCount = UBound(Arr, 2)

Do While i <= j
  Do While dhDate(CStr(Arr(i, k))) < pivot
    i = i + 1
  Loop
  Do While dhDate(CStr(Arr(j, k))) > pivot
    j = j - 1
  Loop
  If i <= j Then
    For t = LBound(Arr, 2) To colCount
      temp = Arr(i, t)
      Arr(i, t) = Arr(j, t)
      Arr(j, t) = temp
    Next t
    i = i + 1
    j = j - 1
  End If
Loop
  If first < j Then Call QuickSort2D_Internal(Arr, first, j, k)
  If i < last Then Call QuickSort2D_Internal(Arr, i, last, k)
End Sub


Function QuickSort1D(ByVal Arr As Variant) As Variant
    Dim sortedArr As Variant
    Dim first As Long, last As Long
    
    sortedArr = Arr
    first = LBound(sortedArr)
    last = UBound(sortedArr)
    
    Call QuickSort1D_Internal(sortedArr, first, last)
    
    QuickSort1D = sortedArr
End Function

Private Sub QuickSort1D_Internal(ByRef Arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim temp As Variant
    
    i = first
    j = last
    pivot = dhDate(CStr(Arr((first + last) \ 2)))
    
    Do While i <= j
        Do While dhDate(CStr(Arr(i))) < pivot
            i = i + 1
        Loop
        Do While dhDate(CStr(Arr(j))) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = Arr(i)
            Arr(i) = Arr(j)
            Arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then QuickSort1D_Internal Arr, first, j
    If i < last Then QuickSort1D_Internal Arr, i, last
End Sub

