VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} B0ThongKeUF 
   ClientHeight    =   9948.001
   ClientLeft      =   -270
   ClientTop       =   -840
   ClientWidth     =   19095
   OleObjectBlob   =   "B0ThongKeUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "B0ThongKeUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BTNHome_Click()
  Call TatForm
  A2ControllerUF.Show
End Sub

Private Sub BTNLdate_Click()
  Me.TBLDate.Text = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub BTNSDate_Click()
  Me.TBSDate.Text = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

'<-- Unload Form -->
Private Sub TatForm()
  Call ClearMemory(Me, UIBasic(True))
  Unload Me
End Sub

Private Sub BTNThongKe_Click()
Dim KhoXuat As String, KhoNhap As String, Opt As String, TenVT As String, SoLuong As Long
Dim SDate As Date, LDate As Date
Dim DK As Boolean
Dim ArrData As Variant, ArrTK(), ArrKeys As Variant, ArrItems As Variant
Dim DictTK As Object

Set DictTK = CreateObject("Scripting.Dictionary")
SDate = dhDate(Me.TBSDate.Text)
LDate = dhDate(Me.TBLDate.Text)
ArrData = wsLogXN.Range("A2:G" & getERow(wsLogXN, "A")).Value2

Select Case True
  Case Me.OptXuatRaCongTruong.Value
    Opt = "Xu" & ChrW(7845) & "t Ra C" & ChrW(244) & "ng Tr" & ChrW(432) & ChrW(7901) & "ng"
  Case Me.OptBanHang.Value
    Opt = "B" & ChrW(225) & "n H" & ChrW(224) & "ng"
  Case Else
    Opt = ""
End Select

For NumLp = LBound(ArrData, 1) To UBound(ArrData, 1)
  CurDate = dhDate(CStr(ArrData(NumLp, 1)))
  KhoXuat = Split(ArrData(NumLp, 5), "<|>", , vbTextCompare)(0)
  KhoNhap = Split(ArrData(NumLp, 5), "<|>", , vbTextCompare)(1)
  DK = CurDate >= SDate And CurDate <= LDate And KhoXuat = Me.CBTenKho.Text And KhoNhap = Opt
  TenVT = ArrData(NumLp, 2)
  SoLuong = ArrData(NumLp, 3)
  If DK Then
    If Not DictTK.exists(TenVT) Then
      DictTK.Add Key:=TenVT, Item:=SoLuong
    Else
      DictTK(TenVT) = DictTK(TenVT) + SoLuong
    End If
  End If
Next NumLp

Me.LBThongKe.Clear
If DictTK.Count > 0 Then
  ArrKeys = DictTK.keys: ArrItems = DictTK.items
  ReDim ArrTK(0 To DictTK.Count - 1, 2)
  For NumLp = 0 To DictTK.Count - 1
    ArrTK(NumLp, 0) = ArrKeys(NumLp)
    ArrTK(NumLp, 1) = TraDonVi(CStr(ArrTK(NumLp, 0)))
    ArrTK(NumLp, 2) = ArrItems(NumLp)
  Next NumLp
  
  Me.LBThongKe.List = ArrTK
  
End If

Set DictTK = Nothing
End Sub

Private Sub UserForm_Initialize()
  Call UIInit(Me, UIBasic(True), UIBasic(False))
  'Me.CBTenKho.List = DSKhoVT
  
  Me.StartUpPosition = 2
  'Resize UF
  With Me
  .Width = .Width * widthRatio / 100
  .Height = .Height * heightRatio / 100
  .Zoom = Application.WorksheetFunction.Min(widthRatio, heightRatio)
  End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = False Then
  Cancel = True
  End If
End Sub

