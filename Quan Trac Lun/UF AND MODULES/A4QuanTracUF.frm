VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} A4QuanTracUF 
   ClientHeight    =   10740
   ClientLeft      =   -360
   ClientTop       =   -1230
   ClientWidth     =   21435
   OleObjectBlob   =   "A4QuanTracUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "A4QuanTracUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTNAddCanNoi_Click()
  If CheckEmpty(Array(Me.TBSCanNoi.Text, Me.TBLCanNoi.Text, Me.TBCDCanNoi.Value)) Then
    Call Msg2User(MsgThieuTT)
    Exit Sub
  End If
  
  If CheckArrExists(Me.LBCanNoi.List, Me.TBSCanNoi.Text, 2, 0) Then
    MsgBox "Ðã Có DL"
    Exit Sub
  End If
  
  If dhDate(Me.TBSCanNoi.Text) > dhDate(Me.TBLCanNoi.Text) Then
    MsgBox "Nhap Sai DL"
    Exit Sub
  End If
  
  If Me.LBCanNoi.ListCount > 0 Then
    For NumLp = 0 To Me.LBCanNoi.ListCount - 1
      dk1 = dhDate(Me.TBSCanNoi.Text) >= dhDate(Me.LBCanNoi.List(NumLp, 0)) And dhDate(Me.TBSCanNoi.Text) <= dhDate(Me.LBCanNoi.List(NumLp, 1))
      dk2 = dhDate(Me.TBLCanNoi.Text) >= dhDate(Me.LBCanNoi.List(NumLp, 0)) And dhDate(Me.TBLCanNoi.Text) <= dhDate(Me.LBCanNoi.List(NumLp, 1))
      If dk1 Or dk2 Then
        MsgBox "Kiem Tra Lai DL"
        Exit Sub
      End If
    Next NumLp
  End If
  
  With Me.LBCanNoi
    ArrTT = Array(Me.TBSCanNoi.Text, Me.TBLCanNoi.Text, Me.TBCDCanNoi.Text, Me.TBCDCanNoi.Text, Me.TBCDCanNoi.Text)
    .AddItem ArrTT(0)
    For NumCol = 1 To UBound(ArrTT)
      .List(.ListCount - 1, NumCol) = ArrTT(NumCol)
    Next NumCol
  End With
  
  Call ClearData(Me, Array("TBSCanNoi", "TBLCanNoi", "TBCDCanNoi"))
End Sub

Private Sub BTNAddDN_Click()
If CheckEmpty(Array(Me.TBNgayDN.Text, Me.TBChieuDay.Value)) Then
  Call Msg2User(MsgThieuTT)
  Exit Sub
End If

If CheckArrExists(Me.LBDapNen.List, Me.TBNgayDN.Text, 2, 0) Then
  MsgBox "Ðã Có DL"
  Exit Sub
End If


With Me.LBDapNen
  .AddItem Me.TBNgayDN.Text
  .List(.ListCount - 1, 1) = Me.TBChieuDay.Value
End With
Call ClearData(Me, Array("TBNgayDN", "TBChieuDay"))
End Sub

Private Sub BTNHome_Click()
  Call TatForm
  A2ControllerUF.Show
End Sub

Private Sub BTNLNoiCan_Click()
  Me.TBLCanNoi = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub BTNMoPhong_Click()
Dim ArrDoLunTrai As Variant, ArrDoLunTim As Variant, ArrDoLunPhai As Variant, ArrData As Variant, ArrCheck As Variant
Dim ArrCDMoc, ArrNgay, ArrSoMiaMoc, ArrrCDMay
Dim DictData As Object: Set DictData = CreateObject("Scripting.Dictionary")
Dim ArrTemp(), ArrDL()
Dim wbMC As Workbook, wsMC As Worksheet
Dim DuongDanMC As String, MaGT As String, TenMC As String

MaGT = Me.CBMaGoiThau.Text
TenMC = Me.CBTenMatCat.Text

ArrCheck = Array(Me.CBMaGoiThau.Text, Me.CBTenMatCat.Text)
If CheckEmpty(ArrCheck) Then
  Call Msg2User(MsgThieuTT)
  Exit Sub
End If

If Me.TBSoNgayPT.Text = "" Or Me.LBCanNoi.ListCount = 0 Or Me.LBPhanPhoiLun.ListCount = 0 Then
  Call Msg2User(MsgThieuTT)
  Exit Sub
End If

If Me.OptXuatMCN.Value Then
  DuongDanMC = DataPath(MaGT) & "\" & TenMC & ".xlsx"
  Set wbMC = Application.Workbooks.Open(DuongDanMC)
  Set wsMC = wbMC.Sheets(1)
Else
  DuongDanMC = ThisWorkbook.Path & "\Templates\Template 03.xlsx"
  Set wbMC = Application.Workbooks.Add(DuongDanMC)
  Set wsMC = wbMC.Sheets(1)
End If

Dim EDate As Long
Dim TongLunTrai As Double, TongLunTim As Double, TongLunPhai As Double
Dim SDate As Date, LDate As Date
ArrData = Me.LBPhanPhoiLun.List

For NumList = LBound(ArrData, 1) To UBound(ArrData, 1)
  SDate = dhDate(CStr(ArrData(NumList, 0)))
  LDate = dhDate(CStr(ArrData(NumList, 1)))
  EDate = ArrData(NumList, 2) - 1
  TongLunTrai = ArrData(NumList, 3)
  TongLunTim = ArrData(NumList, 4)
  TongLunPhai = ArrData(NumList, 5)
  
  ArrNgay = TaoMangNgay(SDate, LDate)
  ArrCDMoc = MangPT(EDate + 1, Me.TBCaoDoMoc.Value)
  ArrSoMiaMoc = TaoMangRandom(EDate + 1, Me.TBSoMiaMin.Value, Me.TBSoMiaMax.Value)
  
  ReDim ArrTemp(0 To UBound(ArrSoMiaMoc))
  For NumLp = 0 To UBound(ArrSoMiaMoc)
    ArrTemp(NumLp) = ArrSoMiaMoc(NumLp) + ArrCDMoc(NumLp)
  Next NumLp
  ArrCDMay = ArrTemp

  ArrDoLunTrai = TaoMangDoLun(0, EDate, TongLunTrai, 0.25, 0.25, 15, 0)
  ArrDoLunTim = TaoMangDoLun(0, EDate, TongLunTim, 0.25, 0.25, 15, 0)
  ArrDoLunPhai = TaoMangDoLun(0, EDate, TongLunPhai, 0.25, 0.25, 15, 0)
  
  For NumLp = LBound(ArrNgay) To UBound(ArrNgay)
    Dim NgayQT As String, CDMoc As Long, SoMiaMoc As Long, CaoDoMay As Long, DoLunTrai As Long, DoLunTim As Long, DoLunPhai As Long
    NgayQT = ArrNgay(NumLp)
    CDMoc = ArrCDMoc(NumLp): SoMiaMoc = ArrSoMiaMoc(NumLp): CaoDoMay = ArrCDMoc(NumLp) + ArrSoMiaMoc(NumLp)
    DoLunTrai = ArrDoLunTrai(NumLp): DoLunTim = ArrDoLunTim(NumLp): DoLunPhai = ArrDoLunPhai(NumLp)
    
    ArrQT = Array(NgayQT, CDMoc, SoMiaMoc, CaoDoMay, DoLunTrai, DoLunTim, DoLunPhai)
    
    If Not DictData.exists(NgayQT) Then
      DictData.Add Key:=NgayQT, Item:=ArrQT
    End If
  Next NumLp
Next NumList

Dim ArrKeys As Variant, ArrItems As Variant, DLQT As Variant, ArrCanNoi As Variant
ArrKeys = DictData.keys: ArrItems = DictData.items

ArrCanNoi = Me.LBCanNoi.List
ArrDapNen = Me.LBDapNen.List


ReDim ArrDL(0 To UBound(ArrKeys), 0 To 25)
For NumRow = 0 To UBound(ArrKeys)
  DLQT = ArrItems(NumRow)
  ArrDL(NumRow, 1) = DLQT(0)
  ArrDL(NumRow, 2) = DLQT(1)
  ArrDL(NumRow, 3) = DLQT(2)
  ArrDL(NumRow, 4) = DLQT(3)
  
  ArrDL(NumRow, 17) = DLQT(4)
  ArrDL(NumRow, 18) = DLQT(5)
  ArrDL(NumRow, 19) = DLQT(6)
  
  For numCan = LBound(ArrCanNoi, 1) To UBound(ArrCanNoi, 1)
    Dim NgayBD As Date, NgayKT As Date
    NgayQT = dhDate(CStr(DLQT(0)))
    NgayBD = dhDate(CStr(ArrCanNoi(numCan, 0)))
    NgayKT = dhDate(CStr(ArrCanNoi(numCan, 1)))
    
    If NgayQT >= NgayBD And NgayQT <= NgayKT Then
      ArrDL(NumRow, 11) = ArrCanNoi(numCan, 2)
      ArrDL(NumRow, 12) = ArrCanNoi(numCan, 3)
      ArrDL(NumRow, 13) = ArrCanNoi(numCan, 4)
    Exit For
    End If
  Next numCan
  
  For numDN = LBound(ArrDapNen, 1) To UBound(ArrDapNen, 1)
    If DLQT(0) = ArrDapNen(numDN, 0) Then
      ArrDL(NumRow, 24) = ArrDapNen(numDN, 1) / 1000
      Exit For
    End If
  Next numDN
  
Next NumRow
For NumRow = LBound(ArrDL, 1) To UBound(ArrDL, 1)
  If NumRow = LBound(ArrDL, 1) Then
    'Gia Tri Ky Truoc
    BanLunTraiKT = Me.TBBanLunTraiKT.Value
    BanLunTimKT = Me.TBBanLunTimKT.Value
    BanLunPhaiKT = Me.TBBanLunPhaiKT.Value
    
    LunTichLuyTraiKT = Me.TBLunTichLuyTraiKT.Value
    LunTichLuyTimKT = Me.TBLunTichLuyTimKT.Value
    LunTichLuyPhaiKT = Me.TBLunTichLuyPhaiKT.Value
    
    CaoDoNenKT = Me.TBCDNenKT.Value
    TongDapTichLuyKT = Me.TBDapTichLuyKT.Value / 1000
    
    ArrDL(NumRow, 14) = BanLunTraiKT - ArrDL(NumRow, 17)
    ArrDL(NumRow, 15) = BanLunTraiKT - ArrDL(NumRow, 18)
    ArrDL(NumRow, 16) = BanLunTraiKT - ArrDL(NumRow, 19)
    
    ArrDL(NumRow, 20) = LunTichLuyTraiKT - ArrDL(NumRow, 17)
    ArrDL(NumRow, 21) = LunTichLuyTraiKT - ArrDL(NumRow, 18)
    ArrDL(NumRow, 22) = LunTichLuyTraiKT - ArrDL(NumRow, 19)
    
    ArrDL(NumRow, 23) = CaoDoNenKT - ArrDL(NumRow, 18) / 1000 + ArrDL(NumRow, 24)
    ArrDL(NumRow, 25) = TongDapTichLuyKT + ArrDL(NumRow, 24)
      
  Else
    ArrDL(NumRow, 14) = ArrDL(NumRow - 1, 14) - ArrDL(NumRow, 17)
    ArrDL(NumRow, 15) = ArrDL(NumRow - 1, 15) - ArrDL(NumRow, 18)
    ArrDL(NumRow, 16) = ArrDL(NumRow - 1, 16) - ArrDL(NumRow, 19)
    
    ArrDL(NumRow, 20) = ArrDL(NumRow - 1, 20) - ArrDL(NumRow, 17)
    ArrDL(NumRow, 21) = ArrDL(NumRow - 1, 21) - ArrDL(NumRow, 18)
    ArrDL(NumRow, 22) = ArrDL(NumRow - 1, 22) - ArrDL(NumRow, 19)
    
    ArrDL(NumRow, 23) = ArrDL(NumRow - 1, 23) - ArrDL(NumRow, 18) / 1000 + ArrDL(NumRow, 24)
    ArrDL(NumRow, 25) = ArrDL(NumRow - 1, 25) + ArrDL(NumRow, 24)
    
  End If
  
    ArrDL(NumRow, 8) = ArrDL(NumRow, 11) + ArrDL(NumRow, 14)
    ArrDL(NumRow, 9) = ArrDL(NumRow, 12) + ArrDL(NumRow, 15)
    ArrDL(NumRow, 10) = ArrDL(NumRow, 13) + ArrDL(NumRow, 16)
    
    ArrDL(NumRow, 5) = ArrDL(NumRow, 4) - ArrDL(NumRow, 8)
    ArrDL(NumRow, 6) = ArrDL(NumRow, 4) - ArrDL(NumRow, 9)
    ArrDL(NumRow, 7) = ArrDL(NumRow, 4) - ArrDL(NumRow, 10)
    
Next NumRow


DRow = getDRow(wsMC, "B")
With wsMC
  .Range("A" & DRow).Resize(UBound(ArrDL) + 1, 26).Value = ArrDL
End With

Set DataDict = Nothing
End Sub

Private Sub BTNNgayBD_Click()
  Me.TBNgayBD = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub BTNNgayDN_Click()
  Me.TBNgayDN = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub BTNNgayKT_Click()
  Me.TBNgayKT = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub BTNPhanPhoiLun_Click()
Dim DictPP As Object
Dim ArrDN As Variant, ArrPP As Variant, NumLp As Long
Dim SDate As Date, LDate As Date, DayPeriod As Long
Dim TongLunTrai As Long, TongLunTim As Long, TongLunPhai As Long
Set DictPP = CreateObject("Scripting.Dictionary")
DictPP.Add Key:=Me.TBNgayBD.Text, Item:=""
DictPP.Add Key:=Me.TBNgayKT.Text, Item:=""

If Me.TBSoNgayPT.Text = "" Or Me.TBLunDVTrai.Value > 15 Or Me.TBLunDVTim.Value > 15 Or Me.TBLunDVPhai.Value > 15 Then
  Call Msg2User("Du Lieu Khong Hop Le")
  Exit Sub
End If

If Me.LBDapNen.ListCount > 0 Then
  ArrDN = Me.LBDapNen.List
  For NumLp = LBound(ArrDN, 1) To UBound(ArrDN, 1)
    DictPP.Add ArrDN(NumLp, 0), Item:=""
  Next NumLp
End If
ArrPP = QuickSort1D(DictPP.keys)

ReDim ArrPPL(0 To UBound(ArrPP) - 1, 0 To 5)
For NumLp = 0 To UBound(ArrPP) - 1
  If NumLp = UBound(ArrPP) - 1 Then
    LDate = dhDate(CStr(ArrPP(NumLp + 1)))
  Else
    LDate = dhDate(CStr(ArrPP(NumLp + 1))) - 1
  End If
  
  SDate = dhDate(CStr(ArrPP(NumLp)))
  DayPeriod = LDate - SDate + 1
  
  ArrPPL(NumLp, 0) = Date2Str(SDate)
  ArrPPL(NumLp, 1) = Date2Str(LDate)
  ArrPPL(NumLp, 2) = DayPeriod
Next NumLp

TongLunTrai = Me.TBDoLunTrai.Value
TongLunTim = Me.TBDoLunTim.Value
TongLunPhai = Me.TBDoLunPhai.Value

ArrPeriod = GetSubArr(ArrPPL, 2, 2, 0)


PPLTrai = PhanPhoiLun(ArrPeriod, TongLunTrai, 1, 0)
PPLTim = PhanPhoiLun(ArrPeriod, TongLunTim, 1, 0)
PPLPhai = PhanPhoiLun(ArrPeriod, TongLunPhai, 1, 0)


For NumLp = 0 To UBound(ArrPPL, 1)
  ArrPPL(NumLp, 3) = PPLTrai(NumLp)
  ArrPPL(NumLp, 4) = PPLTim(NumLp)
  ArrPPL(NumLp, 5) = PPLPhai(NumLp)
Next NumLp

Me.LBPhanPhoiLun.List = ArrPPL
Set DictPP = Nothing
End Sub

Private Sub BTNPhanTichSB_Click()
Dim SoNgayPT As Long
SoNgayPT = dhDate(Me.TBNgayKT.Text) - dhDate(Me.TBNgayBD.Text) + 1
  Me.TBSoNgayPT.Value = SoNgayPT
  Me.TBDoLunTrai = Text2Num(Me.TBCaoDoKTTrai.Value) - Text2Num(Me.TBCaoDoBDTrai.Value)
  Me.TBDoLunTim = Text2Num(Me.TBCaoDoKTTim.Value) - Text2Num(Me.TBCaoDoBDTim.Value)
  Me.TBDoLunPhai = Text2Num(Me.TBCaoDoKTPhai.Value) - Text2Num(Me.TBCaoDoBDPhai.Value)
  
  TBLunDVTrai.Value = Round(Me.TBDoLunTrai.Value / SoNgayPT, 2)
  TBLunDVTim.Value = Round(Me.TBDoLunTim.Value / SoNgayPT, 2)
  TBLunDVPhai.Value = Round(Me.TBDoLunPhai.Value / SoNgayPT, 2)
End Sub

Private Sub BTNRemoveCanNoi_Click()
If Me.LBCanNoi.ListIndex <> -1 Then
  Me.LBCanNoi.RemoveItem Me.LBCanNoi.ListIndex
End If
End Sub

Private Sub BTNRemoveDN_Click()
If Me.LBDapNen.ListIndex <> -1 Then
  Me.LBDapNen.RemoveItem Me.LBDapNen.ListIndex
End If
End Sub

Private Sub BTNSNoiCan_Click()
  Me.TBSCanNoi = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub CBMaGoiThau_Change()
  Me.CBTenMatCat.List = getFileXLSX(DataPath(Me.CBMaGoiThau.Text))
  Me.LBDSMatCat.List = getFileXLSX(DataPath(Me.CBMaGoiThau.Text))
  
  If NumRowTB(DSGoiThau(True, 1), Me.CBMaGoiThau.Text) > 0 Then
    Me.TBTenGoiThau.Text = wsDSGoiThau.Range("B" & NumRowTB(DSGoiThau(True, 1), Me.CBMaGoiThau.Text)).Value
  End If
End Sub

Private Sub InitUF()
  Call UIInit(Me, UIBasic(True), UIBasic(False))
  Me.CBMaGoiThau.List = GetListSubFolder(DataPath)
End Sub
Private Sub UserForm_Initialize()
Call InitUF
Me.StartUpPosition = 2
With Me
  .Width = .Width * widthRatio / 100
  .Height = .Height * heightRatio / 100
  .Zoom = Application.WorksheetFunction.Min(widthRatio, heightRatio)
End With
End Sub
'<-- Unload Form -->
Private Sub TatForm()
  Call ClearMemory(Me, UIBasic(True))
  Unload Me
End Sub

