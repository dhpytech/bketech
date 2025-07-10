Attribute VB_Name = "B_FormModule"
Option Explicit
'<-- Clear Du Lieu Cua TextBox hay  -->
Sub ClearData(UF As Object, ControlNames As Variant)
Dim CtrlName As Variant
Dim Ctrl As Object
    
For Each CtrlName In ControlNames
  On Error Resume Next
  Set Ctrl = UF.Controls(CtrlName)
  On Error GoTo 0
    
  If Not Ctrl Is Nothing Then
    If TypeOf Ctrl Is MSForms.TextBox Or TypeOf Ctrl Is MSForms.ComboBox Then
      Ctrl.Value = ""
    End If
  End If
    
  Set Ctrl = Nothing
Next CtrlName
End Sub

'Khoi Tao Tab cho UserForm
Sub TabInit(UF As UserForm, ArrName As Variant, Optional dhClear As Boolean = True)
Dim NumLp As Long, numTab As Long

For NumLp = LBound(ArrName) To UBound(ArrName)
  UF.Controls(ArrName(NumLp)).TabIndex = NumLp
  If TypeName(UF.Controls(ArrName(NumLp))) = "CheckBox" Then
    UF.Controls(ArrName(NumLp)).Value = False
  ElseIf TypeName(UF.Controls(ArrName(NumLp))) = "TextBox" Then
    If dhClear = True Then
      UF.Controls(ArrName(NumLp)).Text = ""
    End If
  Else
  End If
Next NumLp
End Sub
Sub ScaleUF(UF As Object)
With UF
  .StartUpPosition = 2
  .Width = .Width * widthRatio / 100
  .Height = .Height * heightRatio / 100
  .Zoom = Application.WorksheetFunction.Min(widthRatio, heightRatio)
End With
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
