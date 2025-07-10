VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} A3DSGoiThauUF 
   ClientHeight    =   10545
   ClientLeft      =   -870
   ClientTop       =   -3045
   ClientWidth     =   19650
   OleObjectBlob   =   "A3DSGoiThauUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "A3DSGoiThauUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BTNCapNhat_Click()
Dim DataDict As Object
ArrName = Array("TDDuAn", "TDGoiThau", "CDT", "TVGS", "NhaThau", "TTKhac")
Set DataDict = NoneDict(ArrName)
ArrDL = Array(Me.TBTieuDeDuAn.Text, Me.TBTieuDeGoiThau.Text, Me.TBChuDauTu.Text, Me.TBTuVanGS.Text, Me.TBNhaThau.Text, Me.TBThongTinKhac.Text)

For NumLp = LBound(ArrName) To UBound(ArrName)
  If ArrDL(NumLp) <> "" Then
    DataDict(ArrName(NumLp)) = ArrDL(NumLp)
  End If
Next NumLp

ArrData = Array(Me.TBMaGoiThau.Text, Me.TBTenGoiThau.Text, Me.CBTrangThai.Text, Me.TBViTri.Text, Dict2Str(DataDict))
With wsDSGoiThau.Range("A" & Me.TBNumRow.Value).Resize(1, UBound(ArrData) + 1)
  .Value2 = ArrData
  .WrapText = False
End With
Set DataDict = Nothing
End Sub

Private Sub BTNChinhSuaTT_Click()
  If BTNChinhSuaTT.Value Then
    Call EnableFR(FRThongTin, True)
  Else
    Call EnableFR(FRThongTin, False)
  End If
End Sub

Private Sub BTNHome_Click()
  Call TatForm
  A2ControllerUF.Show
End Sub

Private Sub BTNTaoMoi_Click()
End Sub

Private Sub BTNThoat_Click()
  Me.OptChinhSua.Value = True
End Sub

Private Sub BTNTimKiem_Click()
Dim NumCol As Integer
If Me.CBTimKiem.Text = "" Then
  Me.LBDSGoiThau.Clear
Else
  Select Case True
  Case Me.OptAll.Value
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), Me.CBTimKiem.Text, 1, 1, False)
  Case Me.OptActivate.Value
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), Me.CBTimKiem.Text, 1, 1, False)
  Case Me.OptStop.Value
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), Me.CBTimKiem.Text, 1, 1, False)
  Case Me.OptFinish.Value
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), Me.CBTimKiem.Text, 1, 1, False)
  Case Else
  End Select
End If
End Sub

Private Sub BTNXoaKho_Click()
Dim NumLp As Long, ERow As Long, NumRow As Variant
Dim ArrData As Variant, ThongBao As String
ERow = getERow(wsDSKhoVT, "A")

ArrData = DSKhoVT("All", 2)
For NumLp = LBound(ArrData, 1) To UBound(ArrData, 1) Step 1
  If ArrData(NumLp, 0) = Me.TBTenKho.Text And ArrData(NumLp, 1) = Me.CBTinhTrang.Text Then
    NumRow = NumLp + 2
  Else
    NumRow = -1
  End If
Next NumLp

ArrCode = ArrKhoCT(5, False)
CheckEX = CheckArrExists(ArrCode, Me.LBDSKho.List(Me.LBDSKho.ListIndex, 0), 2, 5)
If CheckEX Then
  ThongBao = "Kh" & ChrW(244) & "ng Th" & ChrW(7875) & " X" & ChrW(243) & "a Kho " & ChrW(272) & "ang C" & ChrW(242) & "n L" & ChrW(432) & "u Tr" & ChrW(7919) & " V" & ChrW(7853) & "t T" & ChrW(432) & ". Vui L" & ChrW(242) & "ng Ki" & ChrW(7875) & "m Tra L" & ChrW(7841) & "i!"
  Call Msg2User(ThongBao)
  Exit Sub
End If

If NumRow >= 0 Then
  wsDSKhoVT.Rows(NumRow).Delete
  Call UpdateLB(Me.LBDSKho, DSKhoVT("All", 2))
Else
  ThongBao = "Kho Kh" & ChrW(244) & "ng T" & ChrW(7891) & "n T" & ChrW(7841) & "i Tr" & ChrW(234) & "n H" & ChrW(7879) & " Th" & ChrW(7889) & "ng. Vui L" & ChrW(242) & "ng Ki" & ChrW(7875) & "m Tra L" & ChrW(7841) & "i!"
  Msg2User (ThongBao)
End If
End Sub

Private Sub BTNTaoGTMoi_Click()
  B1TaoGTMoiUF.Show
End Sub


Private Sub LBDSGoiThau_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim DictTT As Object
Dim ArrData As Variant
Dim NumList As Long, NumLp As Long
NumList = Me.LBDSGoiThau.ListIndex
If NumList >= 0 Then
  Set DictTT = Str2Dict(Me.LBDSGoiThau.List(NumList, 4))
  
  ArrData = Array(Me.LBDSGoiThau.List(NumList, 0), Me.LBDSGoiThau.List(NumList, 2), Me.LBDSGoiThau.List(NumList, 3), Me.LBDSGoiThau.List(NumList, 1), _
                  DictTT("TDDuAn"), DictTT("TDGoiThau"), DictTT("CDT"), DictTT("TVGS"), DictTT("NhaThau"), DictTT("TTKhac"))
  ArrName = Array("TBMaGoiThau", "CBTrangThai", "TBViTri", "TBTenGoiThau", "TBTieuDeDuAn", "TBTieuDeGoiThau", "TBChuDauTu", "TBTuVanGS", "TBNhaThau", "TBThongTinKhac")
  
  For NumLp = LBound(ArrData) To UBound(ArrData)
    Me.Controls(ArrName(NumLp)).Value = CEm2None(ArrData(NumLp), True)
  Next NumLp
  Set DictTT = Nothing
  
  Me.TBNumRow.Value = NumRowTB(DSGoiThau(True, 1), Me.TBMaGoiThau.Text)
End If
  
End Sub

Private Sub OptActivate_Click()
  Call ChooseOption
End Sub

Private Sub OptAll_Click()
  Call ChooseOption
End Sub

Private Sub OptFinish_Click()
  Call ChooseOption
End Sub

Private Sub OptStop_Click()
  Call ChooseOption
End Sub

Sub EnableFR(FR As Frame, State As Boolean)
Dim Ctrl As Control
For Each Ctrl In FR.Controls
  On Error Resume Next
  Ctrl.Enabled = State
  On Error GoTo 0
Next Ctrl
FR.Enabled = State
End Sub

Private Sub InitUF()
  Call UIInit(Me, UIBasic(True), UIBasic(False))
End Sub

Private Sub UserForm_Initialize()
  Call InitUF
  Call ScaleUF(Me)
  Call EnableFR(FRThongTin, False)
  Call UpdateLB(Me.LBDSGoiThau, DSGoiThau(False, 5))
End Sub
'<-- Unload Form -->
Private Sub TatForm()
  Call ClearMemory(Me, UIBasic(True))
  Unload Me
End Sub
Private Sub ChooseOption()
Me.CBTimKiem.Value = ""
Me.LBDSGoiThau.Clear
Select Case True
  Case Me.OptAll.Value
    Me.CBTimKiem.List = DSGTByOpt(False, "All", 5)
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), "", 3, 1, False)
  Case Me.OptActivate.Value
    Me.CBTimKiem.List = DSGTByOpt(False, "Activate", 5)
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), "Activate", 3, 1, False)
  Case Me.OptStop.Value
    Me.CBTimKiem.List = DSGTByOpt(False, "Stop", 5)
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), "Stop", 3, 1, False)
  Case Me.OptFinish.Value
    Me.CBTimKiem.List = DSGTByOpt(False, "Finish", 5)
    Call UpdateLBByCol(Me.LBDSGoiThau, DSGoiThau(False, 5), "Finish", 3, 1, False)
  Case Else
End Select
End Sub

Private Function ArrThongTin()
  ArrThongTin = Array(Me.TBTenKho.Text, Me.CBTinhTrang.Text)
End Function

Private Sub ClearData()
Dim arrLabel, NumLp As Long
arrLabel = Array("TBTenKho", "CBTinhTrang")
  For NumLp = LBound(arrLabel) To UBound(arrLabel)
    Me.Controls(arrLabel(NumLp)).Value = ""
  Next NumLp
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'If CloseMode = False Then
  'Cancel = True
  'End If
'End Sub

