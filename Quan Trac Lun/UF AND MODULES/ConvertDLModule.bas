Attribute VB_Name = "ConvertDLModule"

Function DocSo(ByVal Number, Optional ByVal Font = 1) As String
Dim MyArray
Dim Str
If Number = "" Then
  DocSo = ""
Else
  Str = Format(Abs(Number), "000000000000000000")
  Select Case Font
  Case 1
  MyArray = Array("không ", "m" & ChrW(7897) & "t ", "hai ", "ba ", "b" & ChrW(7889) & "n ", "n" & ChrW(259) & "m ", "sáu ", "b" & ChrW(7843) & "y ", "tám ", _
            "chín ", "tri" & ChrW(7879) & "u ", "nghìn ", "t" & ChrW(7927) & " " & ", ", "tri" & ChrW(7879) & "u ", "nghìn ", "", "tr" & ChrW(259) & "m ", _
            "m" & ChrW(432) & ChrW(417) & "i ", "không " & "m" & ChrW(432) & ChrW(417) & "i" & " không ", "không " & "m" & ChrW(432) & ChrW(417) & "i", _
            "l" & ChrW(7867), "m" & ChrW(432) & ChrW(417) & "i" & " không", "m" & ChrW(432) & ChrW(417) & "i", "m" & ChrW(432) & ChrW(417) & "i" & " n" & ChrW(259) & "m", _
            "m" & ChrW(432) & ChrW(417) & "i" & " l" & ChrW(259) & "m", "m" & ChrW(7897) & "t " & "m" & ChrW(432) & ChrW(417) & "i", "m" & ChrW(432) & ChrW(7901) & "i", "m" & ChrW(432) & ChrW(417) & "i" & " m" & ChrW(7897) & "t", "m" & ChrW(432) & ChrW(417) & "i" & " m" & ChrW(7889) & "t", "Âm ")
  Case 2
  MyArray = Array("khoâng ", "moät ", "hai ", "ba ", "boán ", "naêm ", "saùu ", "baûy ", "taùm ", "chín ", "trieäu ", "ngaøn ", "tyû ", "trieäu ", "ngaøn ", "", "traêm ", "möôi ", _
            "khoâng möôi khoâng ", "khoâng möôi", "leû", "möôi khoâng", "möôi", "möôi naêm", "möôi laêm", "moät möôi", "möôøi", "möôi moät", "möôi moát", "AÂm ")
  Case 3
  MyArray = Array("kh«ng ", "mét ", "hai ", "ba ", "bèn ", "n¨m ", "s¸u ", "b¶y ", "t¸m ", "chÝn ", "triÖu ", "ngµn ", "tû ", "triÖu ", "ngµn ", "", "tr¨m ", "m­¬i ", "kh«ng m­¬i kh«ng ", _
            "kh«ng m­¬i", "lÎ", "m­¬i kh«ng", "m­¬i", "m­¬i n¨m", "m­¬i l¨m", "mét m­¬i", "m­êi", "m­¬i mét", "m­¬i mèt", "¢m ")
  End Select
  
  If Str = "000000000000000000" Then
    DocSo = UCase(Left(MyArray(0), 1)) & Trim(Mid(MyArray(0), 2)) & "."
    Exit Function
  End If
  
  For i = 1 To Len(Str)
    If Left(Str, i) <> 0 And Mid(Str, (Int((i + 2) / 3) - 1) * 3 + 1, 3) <> 0 Then
        DocSo = DocSo & MyArray(Mid(Str, i, 1)) & MyArray(-(9 + i / 3) * (i Mod 3 = 0) - (15 + i Mod 3) * (i Mod 3 <> 0))
    ElseIf i = 9 And Mid(Str, 7, 3) = 0 And Left(Str, 6) <> 0 Then
        DocSo = DocSo & MyArray(12)
    End If
  Next i
  
  DocSo = Trim(Replace(Replace(Replace(Replace(Replace(Replace(DocSo, MyArray(18), MyArray(15)), MyArray(19), MyArray(20)), MyArray(21), MyArray(22)), MyArray(23), MyArray(24)), MyArray(25), MyArray(26)), MyArray(27), MyArray(28)))
  
  If Number < 0 Then
    DocSo = MyArray(29) & DocSo
  End If
    DocSo = Replace(UCase(Left(DocSo, 1)) & Mid(DocSo, 2) & " ", ", ", "") & ChrW(273) & ChrW(7891) & "ng."
End If
End Function

Function V2E(ByVal sContent As String) As String
Dim i As Long
Dim intCode As Long
Dim sChar As String
Dim sConvert As String
V2E = AscW(sContent)
  For i = 1 To Len(sContent)
    sChar = Mid(sContent, i, 1)
    If sChar <> "" Then
      intCode = AscW(sChar)
    End If
    Select Case intCode
      Case 273
        sConvert = sConvert & "d"
      Case 272
        sConvert = sConvert & "D"
      Case 224, 225, 226, 227, 259, 7841, 7843, 7845, 7847, 7849, 7851, 7853, 7855, 7857, 7859, 7861, 7863
        sConvert = sConvert & "a"
      Case 192, 193, 194, 195, 258, 7840, 7842, 7844, 7846, 7848, 7850, 7852, 7854, 7856, 7858, 7860, 7862
        sConvert = sConvert & "A"
      Case 232, 233, 234, 7865, 7867, 7869, 7871, 7873, 7875, 7877, 7879
        sConvert = sConvert & "e"
      Case 200, 201, 202, 7864, 7866, 7868, 7870, 7872, 7874, 7876, 7878
        sConvert = sConvert & "E"
      Case 236, 237, 297, 7881, 7883
        sConvert = sConvert & "i"
      Case 204, 205, 296, 7880, 7882
        sConvert = sConvert & "I"
      Case 242, 243, 244, 245, 417, 7885, 7887, 7889, 7891, 7893, 7895, 7897, 7899, 7901, 7903, 7905, 7907
        sConvert = sConvert & "o"
      Case 210, 211, 212, 213, 416, 7884, 7886, 7888, 7890, 7892, 7894, 7896, 7898, 7900, 7902, 7904, 7906
        sConvert = sConvert & "O"
      Case 249, 250, 361, 432, 7909, 7911, 7913, 7915, 7917, 7919, 7921
        sConvert = sConvert & "u"
      Case 217, 218, 360, 431, 7908, 7910, 7912, 7914, 7916, 7918, 7920
        sConvert = sConvert & "U"
      Case 253, 7923, 7925, 7927, 7929
        sConvert = sConvert & "y"
      Case 221, 7922, 7924, 7926, 7928
        sConvert = sConvert & "Y"
      Case Else
        sConvert = sConvert & sChar
    End Select
  Next
V2E = sConvert
End Function
