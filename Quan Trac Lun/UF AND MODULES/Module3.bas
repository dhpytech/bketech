Attribute VB_Name = "Module3"
Function PhanPhoiLun_GiamDan_WithNoise(SDate As Long, EDate As Long, TongLun As Double, _
                              kMin As Double, kMax As Double, MaxDaily As Long, _
                              Optional EarlyStopDays As Long = 0, _
                              Optional MaxNoise As Long = 0) As Variant
    Dim nDays As Long: nDays = EDate - SDate + 1
    Dim StopAtIndex As Long: StopAtIndex = nDays - 1 - EarlyStopDays
    If StopAtIndex <= 0 Then
        MsgBox "Không d? s? ngày d? phân ph?i lún!", vbCritical
        Exit Function
    End If

    Dim weights() As Double: ReDim weights(1 To StopAtIndex)
    Dim result() As Long: ReDim result(0 To nDays - 1)
    Dim sumWeights As Double, dhscale As Double
    Dim i As Long, remaining As Double
    Dim roundedTotal As Long, adjust As Long
    Dim k As Double

    ' === Ng?u nhiên h? s? k
    Randomize
    k = kMin + Rnd() * (kMax - kMin)

    ' === Tr?ng s? gi?m d?n theo hàm mu
    For i = 1 To StopAtIndex
        weights(i) = 1 - Exp(-k * (StopAtIndex - i + 1))
        sumWeights = sumWeights + weights(i)
    Next i

    ' === Scale theo t?ng lún
    dhscale = TongLun / sumWeights
    For i = 1 To StopAtIndex
        weights(i) = weights(i) * dhscale
        If weights(i) > MaxDaily Then weights(i) = MaxDaily
    Next i

    ' === Bù ph?n thi?u
    sumWeights = 0
    For i = 1 To StopAtIndex
        sumWeights = sumWeights + weights(i)
    Next i

    remaining = TongLun - sumWeights
    For i = 1 To StopAtIndex
        If remaining <= 0 Then Exit For
        Dim canAdd As Double: canAdd = MaxDaily - weights(i)
        If canAdd > 0 Then
            Dim addVal As Double: addVal = Application.Min(remaining, canAdd)
            weights(i) = weights(i) + addVal
            remaining = remaining - addVal
        End If
    Next i

    ' === Làm tròn + c?ng nhi?u nhung không giãn n?
    Dim lastCumulative As Long: lastCumulative = 0
    For i = 1 To StopAtIndex
        Dim raw As Long: raw = Round(weights(i), 0)
        Dim noise As Long: noise = Int((2 * MaxNoise + 1) * Rnd()) - MaxNoise
        raw = raw + noise
        If raw < 0 Then raw = 0
        If raw > MaxDaily Then raw = MaxDaily
        If raw + lastCumulative > TongLun Then raw = TongLun - lastCumulative
        result(i - 1) = raw
        lastCumulative = lastCumulative + raw
    Next i

    ' === Ði?u ch?nh sai s? do nhi?u
    adjust = TongLun - lastCumulative
    result(StopAtIndex - 1) = result(StopAtIndex - 1) + adjust
    If result(StopAtIndex - 1) < 0 Then result(StopAtIndex - 1) = 0

    ' === Gán 0 cho các ngày d?ng
    For i = StopAtIndex To nDays - 1
        result(i) = 0
    Next i

    PhanPhoiLun_GiamDan = result
End Function

Sub TestPhanPhoiLun_WithNoise()
    Dim delta As Variant
    Dim cum() As Long
    Dim SDate As Long: SDate = 0
    Dim EDate As Long: EDate = 6
    Dim TongLun As Double: TongLun = 40
    Dim kMin As Double: kMin = 0.2
    Dim kMax As Double: kMax = 0.4
    Dim MaxDaily As Long: MaxDaily = 15
    Dim EarlyStopDays As Long: EarlyStopDays = 1
    Dim MaxNoise As Long: MaxNoise = 1
    Dim i As Long

    delta = PhanPhoiLun_GiamDan_WithNoise(SDate, EDate, TongLun, kMin, kMax, MaxDaily, EarlyStopDays, MaxNoise)
    If IsEmpty(delta) Then Exit Sub

    ReDim cum(0 To UBound(delta))
    cum(0) = delta(0)
    For i = 1 To UBound(delta)
        cum(i) = cum(i - 1) + delta(i)
    Next i

    With Sheet3
        .Cells.Clear
        .Cells(1, 1).Resize(1, 3).Value = Array("Ngày", "Lún t?ng ngày", "Lún c?ng d?n")
        For i = 0 To UBound(delta)
            .Cells(i + 2, 1).Value = SDate + i
            .Cells(i + 2, 2).Value = delta(i)
            .Cells(i + 2, 3).Value = cum(i)
        Next i
    End With
End Sub


