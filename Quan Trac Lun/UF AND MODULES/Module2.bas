Attribute VB_Name = "Module2"
Function PhanPhoiLunF(weights As Variant, _
                      TotalValue As Long, _
                      Optional MinValue As Long = 0, _
                      Optional Base As Long = 0, _
                      Optional WeightCol As Variant = Empty) As Variant
    Dim i As Long, j As Long, n As Long
    Dim is2D As Boolean
    Dim rLBound As Long, rUBound As Long, actualCol As Long
    Dim totalWeight As Double, wVal As Double
    
    '- Xác d?nh kích thu?c m?ng -'
    On Error Resume Next
    is2D = (Not IsEmpty(LBound(weights, 2)))
    On Error GoTo 0
    
    If is2D Then
        rLBound = LBound(weights, 1): rUBound = UBound(weights, 1)
        n = rUBound - rLBound + 1
        If IsEmpty(WeightCol) Then
            actualCol = LBound(weights, 2)     'm?c d?nh c?t d?u
        Else
            actualCol = WeightCol
        End If
    Else
        rLBound = LBound(weights): rUBound = UBound(weights)
        n = rUBound - rLBound + 1
    End If
    
    '- Ki?m tra TotalValue d? d? c?p MinValue -'
    If MinValue * n > TotalValue Then
        Err.Raise vbObjectError + 1, , "MinValue × n vu?t TotalValue"
    End If
    
    '--- Kh?i t?o m?ng k?t qu? ---'
    Dim result() As Long, residual() As Double, idx() As Long
    ReDim result(Base To Base + n - 1)
    ReDim residual(Base To Base + n - 1)
    ReDim idx(Base To Base + n - 1)
    
    '- T?ng tr?ng s? -'
    For i = rLBound To rUBound
        If is2D Then wVal = weights(i, actualCol) Else wVal = weights(i)
        totalWeight = totalWeight + wVal
    Next i
    
    'Ph?n giá tr? còn l?i sau khi c?p MinValue'
    Dim remainTotal As Long: remainTotal = TotalValue - MinValue * n
    Dim remainUnits As Long
    
    '- Phân b? ph?n còn l?i theo tr?ng s? -'
    Dim rawShare As Double, floorShare As Long, sumFloor As Long
    For i = rLBound To rUBound
        j = i - rLBound + Base
        If is2D Then wVal = weights(i, actualCol) Else wVal = weights(i)
        
        rawShare = wVal / totalWeight * remainTotal
        floorShare = Int(rawShare)                 'Floor v? s? nguyên
        residual(j) = rawShare - floorShare        'ph?n l?
        result(j) = floorShare + MinValue          'c?ng MinValue ngay t?i dây
        sumFloor = sumFloor + floorShare
        idx(j) = j
    Next i
    
    remainUnits = remainTotal - sumFloor          's? don v? chua phân
    '- S?p x?p residual gi?m d?n (bong bóng don gi?n) -'
    Dim tempIdx As Long, k As Long
    For i = Base To Base + n - 2
        For k = i + 1 To Base + n - 1
            If residual(idx(i)) < residual(idx(k)) Then
                tempIdx = idx(i): idx(i) = idx(k): idx(k) = tempIdx
            End If
        Next k
    Next i
    
    '- Phát n?t ph?n du -'
    For i = Base To Base + remainUnits - 1
        result(idx(i)) = result(idx(i)) + 1
    Next i
    
    PhanPhoiLun = result
End Function

Function PhanPhoiLun(weights As Variant, _
                      TotalValue As Long, _
                      Optional MinValue As Long = 0, _
                      Optional Base As Long = 0, _
                      Optional WeightCol As Variant = Empty) As Variant
    Dim i As Long, j As Long, n As Long
    Dim is2D As Boolean
    Dim rLBound As Long, rUBound As Long, actualCol As Long
    Dim totalWeight As Double, wVal As Double
    
    '- Xác d?nh kích thu?c m?ng -'
    On Error Resume Next
    is2D = (Not IsEmpty(LBound(weights, 2)))
    On Error GoTo 0
    
    If is2D Then
        rLBound = LBound(weights, 1): rUBound = UBound(weights, 1)
        n = rUBound - rLBound + 1
        If IsEmpty(WeightCol) Then
            actualCol = LBound(weights, 2)
        Else
            actualCol = WeightCol
        End If
    Else
        rLBound = LBound(weights): rUBound = UBound(weights)
        n = rUBound - rLBound + 1
    End If
    
    '- Ki?m tra TotalValue d? d? c?p MinValue -'
    If MinValue * n > TotalValue Then
        Err.Raise vbObjectError + 1, , "MinValue × n vu?t TotalValue"
    End If
    
    '--- Kh?i t?o m?ng k?t qu? ---'
    Dim result() As Long, residual() As Double, idx() As Long
    ReDim result(Base To Base + n - 1)
    ReDim residual(Base To Base + n - 1)
    ReDim idx(Base To Base + n - 1)
    
    '- Tính t?ng tr?ng s? -'
    For i = rLBound To rUBound
        If is2D Then wVal = weights(i, actualCol) Else wVal = weights(i)
        If wVal < 0 Then Err.Raise vbObjectError + 2, , "Tr?ng s? không du?c âm"
        totalWeight = totalWeight + wVal
    Next i
    
    'Ph?n giá tr? còn l?i sau khi c?p MinValue'
    Dim remainTotal As Long: remainTotal = TotalValue - MinValue * n
    Dim remainUnits As Long
    
    '- Phân b? ph?n còn l?i theo tr?ng s? -'
    Dim rawShare As Double, floorShare As Long, sumFloor As Long
    For i = rLBound To rUBound
        j = i - rLBound + Base
        If is2D Then wVal = weights(i, actualCol) Else wVal = weights(i)
        
        rawShare = wVal / totalWeight * remainTotal
        
        '--- S?a l?i lún âm do Int ---'
        floorShare = Fix(rawShare)
        
        residual(j) = rawShare - floorShare
        result(j) = floorShare + MinValue
        sumFloor = sumFloor + floorShare
        idx(j) = j
    Next i
    
    remainUnits = remainTotal - sumFloor
    
    '- S?p x?p residual gi?m d?n -'
    Dim tempIdx As Long, k As Long
    For i = Base To Base + n - 2
        For k = i + 1 To Base + n - 1
            If residual(idx(i)) < residual(idx(k)) Then
                tempIdx = idx(i): idx(i) = idx(k): idx(k) = tempIdx
            End If
        Next k
    Next i
    
    '- Phát n?t ph?n du -'
    For i = Base To Base + remainUnits - 1
        result(idx(i)) = result(idx(i)) + 1
    Next i
    
    PhanPhoiLun = result
End Function



Sub TestDistribute()
    Dim weights As Variant
    Dim Luns As Variant
    Dim i As Long
    
    weights = Array(28)
    Luns = PhanPhoiLun(weights, 300, 1, 0) ' T?ng 300mm, t?i thi?u m?i giai do?n 1mm

    For i = LBound(Luns) To UBound(Luns)
        MsgBox "Giai do?n " & (i - LBound(Luns) + 1) & ": " & Luns(i) & " mm"
    Next i
End Sub

Function TaoMangNgay(ByVal NgayBD As Date, ByVal NgayKT As Date, Optional Base As Long = 0) As Variant
    Dim SoNgay As Long, i As Long
    Dim KQ() As String
    
    If NgayKT < NgayBD Then
        Err.Raise vbObjectError + 1, , "Ngày k?t thúc ph?i >= ngày b?t d?u"
    End If
    
    SoNgay = NgayKT - NgayBD + 1
    ReDim KQ(Base To Base + SoNgay - 1)
    
    For i = 0 To SoNgay - 1
        KQ(Base + i) = Date2Str(NgayBD + i)
    Next i
    
    TaoMangNgay = KQ
End Function

Function TaoMangRandom(SoLuong As Long, MinVal As Long, MaxVal As Long, Optional Base As Long = 0) As Variant
    Dim KQ() As Long
    Dim i As Long
    
    If MaxVal < MinVal Then
        Err.Raise vbObjectError + 1, , "MaxVal ph?i >= MinVal"
    End If
    If SoLuong <= 0 Then
        Err.Raise vbObjectError + 2, , "S? lu?ng ph?i > 0"
    End If
    
    ReDim KQ(Base To Base + SoLuong - 1)
    Randomize ' Kh?i t?o b? sinh s? ng?u nhiên

    For i = Base To Base + SoLuong - 1
        KQ(i) = Int((MaxVal - MinVal + 1) * Rnd + MinVal)
    Next i

    TaoMangRandom = KQ
End Function

Function MangPT(SoLuong As Long, GiaTri As Variant, Optional Base As Long = 0) As Variant
    Dim KQ() As Variant
    Dim i As Long
    
    If SoLuong <= 0 Then
        Err.Raise vbObjectError + 1, , "S? lu?ng ph?i > 0"
    End If
    
    ReDim KQ(Base To Base + SoLuong - 1)
    
    For i = Base To Base + SoLuong - 1
        KQ(i) = GiaTri
    Next i
    
    MangPT = KQ
End Function

