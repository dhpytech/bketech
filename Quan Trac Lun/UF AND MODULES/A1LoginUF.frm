VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} A1LoginUF 
   ClientHeight    =   4410
   ClientLeft      =   -60
   ClientTop       =   -270
   ClientWidth     =   7170
   OleObjectBlob   =   "A1LoginUF.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "A1LoginUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Const WM_SETTEXT = &HC
#If Win64 Then
    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Const WM_SETICON = &H80
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Sub unhideImage_Click()
Dim numClick As Long
If Me.userPassword.PasswordChar = "*" Then
  Me.userPassword.PasswordChar = ""
  Me.unhideImage.Picture = LoadPicture(dhRoot & "\Icons\hide.ico")
Else
  Me.userPassword.PasswordChar = "*"
  Me.unhideImage.Picture = LoadPicture(dhRoot & "\Icons\unhide.ico")
End If
End Sub
Private Sub UserForm_Initialize()
Dim hWnd&, lStyle
Dim strIconPath As String
Dim lngIcon As Long
Dim lngHWnd As Long

    ' Get the icon from the source
    lngIcon = ExtractIcon(0, strIconPath, 0)
    ' Get the window handle of the userform
    lngHWnd = FindWindow("ThunderDFrame", Me.Caption)
    'Set the big (32x32) and small (16x16) icons
    SendMessage lngHWnd, WM_SETICON, True, lngIcon
    SendMessage lngHWnd, WM_SETICON, False, lngIcon
    'SendMessage lnghWnd, WM_SETICON, False, lngIcon '16x16
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    DefWindowProcW hWnd, WM_SETTEXT, 0, StrPtr(ChrW(272) & ChrW(259) & "ng Nh" & ChrW(7853) & "p")
    
    On Error Resume Next
    Me.unhideImage.Picture = LoadPicture(dhRoot & "\Icons\unhide.ico")
    wsCache.Shapes(1).Visible = msoFalse
    
    Me.userName.ControlTipText = "T" & ChrW(224) & "i kho" & ChrW(7843) & "n"
    Me.userPassword.ControlTipText = "M" & ChrW(7853) & "t kh" & ChrW(7849) & "u"
    
    Me.TaiKhoanListBox.List = UserData
    Me.StartUpPosition = 2
    'Resize UF
    With Me
    .Zoom = Int(Application.WorksheetFunction.Min(widthRatio, heightRatio) * 1)
    .Width = .Width * .Zoom / 100
    .Height = .Height * .Zoom / 100
    End With
End Sub
Private Sub loginButton_Click()
Dim arrUser As Variant
If Me.Author.Caption = "Design By DhTech" Then
  Me.noticeLogin = ChrW(272) & "ang x" & ChrW(225) & "c th" & ChrW(7921) & "c T" & ChrW(224) & "i Kho" & ChrW(7843) & "n. Vui l" & ChrW(242) & "ng ch" & ChrW(7901) & " ..."
  Me.noticeLogin.ForeColor = vbBlue
  Me.noticeLogin.Visible = True
  
  arrUser = UserAuth(Me.userName.Text, Me.userPassword.Text)
  If arrUser(0) = False Then
    Me.noticeLogin = "T" & ChrW(234) & "n T" & ChrW(224) & "i Kho" & ChrW(7843) & "n Ho" & ChrW(7863) & "c M" & ChrW(7853) & "t Kh" & ChrW(7849) & "u Kh" & ChrW(244) & "ng " & ChrW(272) & ChrW(250) & "ng"
    Me.noticeLogin.ForeColor = vbRed
    Me.noticeLogin.Visible = True
  Else
    CurrentUser = Me.userName
    Me.userName = "": Me.userPassword = "": Me.noticeLogin.Visible = False
    Unload Me
    A2ControllerUF.Show
  End If
  Erase arrUser
Else
End If
End Sub
Private Sub cancelbutton_Click()
  Unload Me
  ThisWorkbook.Close False
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = False Then
  Cancel = True
  End If
End Sub
Private Function UserAuth(UsName$, UsPass$) As Variant
Dim arrUser As Variant
Dim UsCheck As Boolean, UsPosition$
Dim numUs%
  arrUser = Me.TaiKhoanListBox.List
  UsCheck = False: UsPosition = "Invalid"
  If UBound(arrUser, 1) = 0 Then
  Else
    For numUs = LBound(arrUser, 1) To UBound(arrUser, 1)
      If UsName = arrUser(numUs, 0) And HashSHA256(UsPass) = arrUser(numUs, 1) Then
        UsCheck = True: UsPosition = arrUser(numUs, 2)
        Exit For
      Else
      End If
    Next numUs
  End If
  UserAuth = Array(UsCheck, UsName, UsPosition)
Erase arrUser
End Function
