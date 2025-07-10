Attribute VB_Name = "DateMoudule"
Option Explicit
Function dhDate(dateString As String, Optional ByVal style As String = "dd/mm/yyyy") As Date
Dim dateArr As Variant
Dim dhDay%, dhMonth%, dhYear%
  If dateString Like "??/??/????" Then
    dateArr = Split(dateString, "/", , vbTextCompare)
    dhDay = Int(dateArr(0)): dhMonth = Int(dateArr(1)): dhYear = Int(dateArr(2))
    dhDate = DateSerial(dhYear, dhMonth, dhDay)
  ElseIf dateString Like "??-??-????" Then
    dateArr = Split(dateString, "-", , vbTextCompare)
    dhDay = Int(dateArr(0)): dhMonth = Int(dateArr(1)): dhYear = Int(dateArr(2))
    dhDate = DateSerial(dhYear, dhMonth, dhDay)
  Else
    dhDate = Date
  End If
End Function
Function HeaderDate2Str(Ngay As Date, Optional Location As String = "")
  If Location = "" Then
    HeaderDate2Str = "Ngày " & Format(Ngay, "dd") & " Tháng " & Format(Ngay, "mm") & " N" & ChrW(259) & "m " & Format(Ngay, "yyyy")
  Else
    HeaderDate2Str = Location & ", Ngày " & Format(Ngay, "dd") & " Tháng " & Format(Ngay, "mm") & " N" & ChrW(259) & "m " & Format(Ngay, "yyyy")
  End If
End Function

Function Date2Str(Ngay As Date, Optional ByVal style As String = "dd/mm/yyyy") As String
  Date2Str = Format(Ngay, style)
End Function


