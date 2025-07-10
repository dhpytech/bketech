VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} A2ControllerUF 
   ClientHeight    =   9825.001
   ClientLeft      =   -1080
   ClientTop       =   -3930
   ClientWidth     =   14850
   OleObjectBlob   =   "A2ControllerUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "A2ControllerUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub BTNCaiDat_Click()
  'Call TatForm
  'B5CaiDatUF.Show
End Sub

Private Sub BTNGoiThau_Click()
  Call TatForm
  A3DSGoiThauUF.Show
End Sub

Private Sub BTNLuu_Click()
  ThisWorkbook.Save
End Sub

Private Sub BTNMatCat_Click()
  Call TatForm
  A5QLMatCatUF.Show
End Sub

Private Sub BTNQuanTrac_Click()
  Call TatForm
  A4QuanTracUF.Show
End Sub

Private Sub BTNThoat_Click()
  Call TatForm
  'CurrentUser = ""
  A1LoginUF.Show
End Sub

Private Sub BTNThongKe_Click()
  Call TatForm
  B0ThongKeUF.Show
End Sub

Private Sub BTNUser_Click()
  'Call TatForm
  'B7MKAdminUF.Show
End Sub

Private Sub BTNXuatKho_Click()
  Call TatForm
  A4XuatkhoUF.Show
End Sub
'<-- Initial -->
Private Sub UserForm_Initialize()
  Call HideAllUI
  Call UIInit(Me, UIController(True), UIController(False))
  Call ScaleUF(Me)
End Sub
'<-- Unload Form -->
Private Sub TatForm()
  Call ClearMemory(Me, UIController(True))
  Unload Me
End Sub
'<-- Disable X -->
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = False Then
  Cancel = True
  End If
End Sub
