VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} B1TaoGTMoiUF 
   ClientHeight    =   10176
   ClientLeft      =   -60
   ClientTop       =   -180
   ClientWidth     =   19365
   OleObjectBlob   =   "B1TaoGTMoiUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "B1TaoGTMoiUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function ArrName(Optional Name As Boolean = True) As Variant
  If Name Then
    ArrName = Array("TBMaGoiThau", "CBTrangThai", "TBViTri", "TBTenGoiThau", "TBTieuDeDuAn", "TBTieuDeGoiThau", "TBChuDauTu", "TBTuVanGS", "TBNhaThau", "TBThongTinKhac")
  Else
    ArrName = Array(Me.TBMaGoiThau.Text, Me.CBTrangThai.Text, Me.TBViTri.Text, Me.TBTenGoiThau.Text, Me.TBTieuDeDuAn.Text, Me.TBTieuDeGoiThau.Text, Me.TBChuDauTu.Text, Me.TBTuVanGS.Text, Me.TBNhaThau.Text)
  End If
End Function

Private Sub BTNReset_Click()
  Call ClearData(Me, ArrName)
End Sub

Private Sub BTNThoat_Click()
  Unload Me
End Sub

Private Sub BTNTaoMoi_Click()
'1. Kiem Tra Các Dieu Kien Ban Dau
If CheckEmpty(ArrName(False)) Then
  Msg2User (MsgThieuTT)
  Exit Sub
End If

If CheckArrExists(DSGoiThau(False, 1), Me.TBMaGoiThau.Text, 2, 1) Then
  Msg2User ("M" & ChrW(227) & " G" & ChrW(243) & "i Th" & ChrW(7847) & "u " & ChrW(272) & ChrW(227) & " T" & ChrW(7891) & "n T" & ChrW(7841) & "i. Vui L" & ChrW(242) & "ng Ch" & ChrW(7885) & "n M" & ChrW(227) & " Kh" & ChrW(225) & "c!")
  Exit Sub
End If
'2. Ghi Goi Thau Vào Data
Dim TDDuAn As String, TDGoiThau As String, CDT As String, TVGS As String, NhaThau As String, TTKhac As String
Dim ArrTT As Variant, ArrData As Variant

TDDuAn = "TDDuAn|:" & Me.TBTieuDeDuAn.Text: TDGoiThau = "TDGoiThau|:" & Me.TBTieuDeGoiThau.Text: TTKhac = "TTKhac|:" & Me.TBThongTinKhac.Text
CDT = "CDT|:" & Me.TBChuDauTu.Text: TVGS = "TVGS|:" & Me.TBTuVanGS.Text: NhaThau = "NhaThau|:" & Me.TBNhaThau.Text

ArrTT = Array(TDDuAn, TDGoiThau, TTKhac, CDT, TVGS, NhaThau)
ArrData = Array(Me.TBMaGoiThau.Text, Me.TBTenGoiThau.Text, Me.CBTrangThai.Text, Me.TBViTri.Text, Join(ArrTT, "<|>"))

With wsDSGoiThau.Range("A" & getDRow(wsDSGoiThau, "A")).Resize(1, UBound(ArrData) + 1)
  .Value = ArrData
  .WrapText = False
End With

'3. Tao Folder Chua File Mat Cat
Dim DhPath As String
DhPath = DataPath(Me.TBMaGoiThau.Text)
If Not CheckFolderPath(DhPath) Then
  Call CreateFolder(DhPath)
End If

'4. Reset Data
Call BTNReset_Click
End Sub

Private Sub UserForm_Initialize()
  Call TabInit(Me, ArrName(True), True)
  Me.CBTrangThai.List = Array("Activate", "Stop")
End Sub
