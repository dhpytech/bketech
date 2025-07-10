Attribute VB_Name = "UXControl"
Public Sub Msg2User(notice As String, Optional Title As String = "Thông Báo")
Dim dhNotice As Object
  Set dhNotice = CreateObject("WScript.Shell")
  dhNotice.Popup notice, , Title, 0 + 64
Set dhNotice = Nothing
End Sub

Sub HideUI(UF As UserForm, Status, ParamArray Values())
If ThisWorkbook.ReadOnly Then
  For NumLp = LBound(Values) To UBound(Values)
    If Status = "Visible" Then
      UF.Controls(Values(NumLp)).Visible = False
    Else
      UF.Controls(Values(NumLp)).Enabled = False
    End If
  Next NumLp
End If
End Sub

Sub UIInit(UF As UserForm, arrLabel As Variant, arrImageName As Variant)
  Dim NumLp%
  For NumLp = LBound(arrLabel, 1) To UBound(arrLabel, 1)
    UF.Controls(arrLabel(NumLp)).Picture = LoadPicture(dhRoot & "\Icons\" & arrImageName(NumLp))
  Next NumLp
End Sub
