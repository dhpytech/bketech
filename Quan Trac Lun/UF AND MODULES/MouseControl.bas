Attribute VB_Name = "MouseControl"
#If Win64 Then
  Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
  Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If
  Private Type POINTAPI
    X As Long
    Y As Long
  End Type
Public Function getXmouse()
Dim Pt As POINTAPI, mWnd As Long
  GetCursorPos Pt 'L?y giá tr? X và Y t?i v? trí con tr? chu?t
  mWnd = WindowFromPoint(Pt.X, Pt.Y) 'Luu bi?n mWnd = v?i Handle t?i d?i tu?ng có t?a d? X và Y b?ng hàm WindowFromPoint
  'MsgBox "handle= " & mWnd 'Hi?n Msgbox giá tr? Handle c?a d?i tu?ng dã ch?n
  getXmouse = Pt.X
End Function
Public Function getYmouse()
Dim Pt As POINTAPI, mWnd As Long
  GetCursorPos Pt 'L?y giá tr? X và Y t?i v? trí con tr? chu?t
  mWnd = WindowFromPoint(Pt.X, Pt.Y) 'Luu bi?n mWnd = v?i Handle t?i d?i tu?ng có t?a d? X và Y b?ng hàm WindowFromPoint
  'MsgBox "handle= " & mWnd 'Hi?n Msgbox giá tr? Handle c?a d?i tu?ng dã ch?n
  getYmouse = Pt.Y
End Function
