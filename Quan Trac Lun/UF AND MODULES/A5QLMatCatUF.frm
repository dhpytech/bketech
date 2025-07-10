VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} A5QLMatCatUF 
   ClientHeight    =   10044
   ClientLeft      =   -150
   ClientTop       =   -570
   ClientWidth     =   19620
   OleObjectBlob   =   "A5QLMatCatUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "A5QLMatCatUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InitUF()
  Call UIInit(Me, UIBasic(True), UIBasic(False))
  Me.CBMaGoiThau.List = DSGTByOpt(False, "Activate", 5)
  Me.Opt3BanDo.Value = True
  
  arrTitle = Array("Th" & ChrW(244) & "ng Tin Chung", "S" & ChrW(7889) & " " & ChrW(272) & ChrW(7885) & "c Mia ", " Cao " & ChrW(272) & ChrW(7897) & " " & ChrW(272) & ChrW(7881) & "nh C" & ChrW(7847) & "n ", " Chi" & ChrW(7873) & "u D" & ChrW(224) & "i C" & ChrW(7847) & "n ", _
                   " Cao " & ChrW(272) & ChrW(7897) & " B" & ChrW(224) & "n L" & ChrW(250) & "n ", " " & ChrW(272) & ChrW(7897) & " L" & ChrW(250) & "n H" & ChrW(7857) & "ng Ng" & ChrW(224) & "y ", " T" & ChrW(7893) & "ng L" & ChrW(250) & "n T" & ChrW(237) & "ch L" & ChrW(361) & "y")
  For numFR = LBound(arrTitle) To UBound(arrTitle)
    Me.Controls("FRLD" & numFR + 1).Caption = arrTitle(numFR)
  Next numFR
End Sub

Private Sub BTNHome_Click()
  Call TatForm
  A2ControllerUF.Show
End Sub

Private Sub BTNNgayLD_Click()
  Me.TBNgayLD = Format(A0CalendarUF.GetDate(), "dd/mm/yyyy")
End Sub

Private Sub BTNTaoMCN_Click()
Dim DictTT As Object
Dim arrGT As Variant, ArrData As Variant, ArrMCN As Variant, ArrTT As Variant
Dim MaGT As String, TTGT As String, TDDuAn As String, TDGoiThau As String, TDNhaThau As String
Dim numGT As Long
'Kiem Tra Dieu Kien Ban Dau
MaGT = Me.CBMaGoiThau.Text
arrGT = DSGoiThau(True, 5)
ArrMCN = DSMatCatByOpt(False, MaGT, 5)

ArrTT = Array(Me.TBTenMCN.Text, Me.CBMaGoiThau.Text)
If CheckEmpty(ArrTT) Then
  Call Msg2User(MsgThieuTT)
  Exit Sub
End If

If CheckArrExists(ArrMCN, Me.TBTenMCN.Text, 1, 1) Then
  Call Msg2User("M" & ChrW(7863) & "t C" & ChrW(7855) & "t Ngang " & ChrW(272) & ChrW(227) & " T" & ChrW(7891) & "n T" & ChrW(7841) & "i. Vui L" & ChrW(242) & "ng Ki" & ChrW(7875) & "m Tra L" & ChrW(7841) & "i!")
  Exit Sub
End If

'Lay Thong Tin Tieu De
For numGT = LBound(arrGT, 1) To UBound(arrGT, 1)
  If MaGT = arrGT(numGT, 1) Then
    TTGT = arrGT(numGT, 5)
    Exit For
  End If
Next numGT

'Tao File MCN
Set DictTT = Str2Dict(TTGT)
  TDDuAn = DictTT("TDDuAn")
  TDGoiThau = DictTT("TDGoiThau")
  TDNhaThau = "NH" & ChrW(192) & " TH" & ChrW(7846) & "U: " & DictTT("NhaThau")
Set DictTT = Nothing
ArrData = Array(1, Me.TBNgayLD.Value, Me.TBCaoDoMoc.Value, Me.TBSoMiaMoc.Value, Me.TBCDMay.Value, Me.TBSDMTrai.Value, Me.TBSDMTim.Value, Me.TBSDMPhai.Value, Me.TBDinhCanTrai.Value, Me.TBDinhCanTim.Value, Me.TBDinhCanPhai.Value, _
                Me.TBChieuDaiCanTrai.Value, Me.TBChieuDaiCanTim.Value, Me.TBChieuDaiCanPhai.Value, Me.TBBanLunTrai.Value, Me.TBBanLunTim.Value, Me.TBBanLunPhai.Value, Me.TBDoLunTrai.Value, Me.TBDoLunTim.Value, Me.TBDoLunPhai.Value, _
                Me.TBTongLunTrai.Value, Me.TBTongLunTim.Value, Me.TBTongLunPhai.Value, Me.TBCDNen.Value, Me.TBChieuDay.Value, Me.TBTongDap.Value)
             
Dim wbMCN As Workbook
Dim wsMCN As Worksheet
Dim TemplatePath As String, SavePath As String, LyTrinh As String, CDMoc As Double, NgayLD As String

LyTrinh = Me.TBTenMCN.Text: CDMoc = Me.TBCaoDoMoc.Value: NgayLD = Me.TBNgayLD.Text
TemplatePath = ThisWorkbook.Path & "\Templates\Template 03.xlsx"
Set wbMCN = Application.Workbooks.Add(TemplatePath)
Set wsMCN = wbMCN.Sheets(1)

wsMCN.Range("A1").Value = TDDuAn
wsMCN.Range("A2").Value = TDGoiThau
wsMCN.Range("A3").Value = TDNhaThau
wsMCN.Range("D7").Value = LyTrinh
wsMCN.Range("T6").Value = CDMoc
wsMCN.Range("T7").Value = NgayLD
wsMCN.Range("A11").Resize(1, UBound(ArrData) + 1).Value = ArrData

SavePath = DataPath(Me.CBMaGoiThau.Text) & "\" & Me.TBTenMCN.Text & ".xlsx"
wbMCN.SaveAs SavePath
wbMCN.Close False
Set wbMCN = Nothing
Set wsMCN = Nothing

'Ghi Du Lieu Vao Ung Dung
Dim IdMCN As String, LoaiMCN As String
Dim ArrDB As Variant, ArrID As Variant

If Me.Opt3BanDo.Value Then
  LoaiMCN = "L03"
Else
  LoaiMCN = "L05"
End If

ArrID = Array(Me.TBTenMCN, Me.CBMaGoiThau, LoaiMCN)
IdMCN = Join(ArrID, "<|>")
ArrDB = Array(IdMCN, Me.TBTenMCN, Me.CBMaGoiThau, LoaiMCN)
wsDSMatCat.Range("A" & getDRow(wsDSMatCat, "A")).Resize(1, UBound(ArrDB) + 1).Value2 = ArrDB

End Sub

Private Sub OptChinhSuaTT_Click()
  Call ChuyenTab
End Sub

Private Sub OptTaoMCN_Click()
  Call ChuyenTab
End Sub

Private Sub TBCaoDoMoc_Change()
  Call TinhCDMay
End Sub

Private Sub TBSoMiaMoc_Change()
  Call TinhCDMay
End Sub

Private Sub UserForm_Initialize()
  Call InitUF
  Call ScaleUF(Me)
End Sub
'<-- Unload Form -->
Private Sub TatForm()
  Call ClearMemory(Me, UIBasic(True))
  Unload Me
End Sub
Private Sub ChuyenTab()
Select Case True
  Case Me.OptTaoMCN.Value
    numTab = 0
  Case Me.OptChinhSuaTT.Value
    numTab = 1
  Case Else
    numTab = 0
End Select
  Me.MultiPage1.Value = numTab
End Sub
Private Sub TinhCDMay()
  Me.TBCDMay.Value = Text2Num(Me.TBCaoDoMoc.Value) + Text2Num(Me.TBSoMiaMoc.Value)
End Sub

