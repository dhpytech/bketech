Attribute VB_Name = "Module4"
Function TaoMangDoLun(SDate As Long, EDate As Long, TongLun As Double, kMin As Double, kMax As Double, MaxDaily As Long, Optional EarlyStopDays As Long = 0) As Variant
    Dim nDays As Long: nDays = EDate - SDate + 1
    Dim StopAtIndex As Long: StopAtIndex = nDays - 1 - EarlyStopDays

    If StopAtIndex <= 0 Then
        MsgBox "Không d? s? ngày d? phân ph?i lún"
        Exit Function
    End If

    Dim weights() As Double: ReDim weights(1 To StopAtIndex)
    Dim result() As Long: ReDim result(0 To nDays - 1)
    Dim sumWeights As Double, dhscale As Double
    Dim i As Long, remaining As Double
    Dim roundedTotal As Long, adjust As Long
    Dim k As Double

    Randomize
    k = kMin + Rnd() * (kMax - kMin)

    For i = 1 To StopAtIndex
        weights(i) = 1 - Exp(-k * (StopAtIndex - i + 1))
        sumWeights = sumWeights + weights(i)
    Next i

    dhscale = TongLun / sumWeights
    For i = 1 To StopAtIndex
        weights(i) = weights(i) * dhscale
        If weights(i) > MaxDaily Then weights(i) = MaxDaily
    Next i

    ' Bù l?i ph?n dã b? c?t b?i MaxDaily
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

    ' Làm tròn và tính t?ng
    For i = 1 To StopAtIndex
        result(i - 1) = Round(weights(i), 0)
        roundedTotal = roundedTotal + result(i - 1)
    Next i

    ' Ði?u ch?nh ph?n du
    adjust = TongLun - roundedTotal
    If adjust <> 0 Then
        ' Phân b? du/thi?u vào các di?m phù h?p
        Dim sign As Long: sign = Sgn(adjust)
        Dim adjRemain As Long: adjRemain = Abs(adjust)

        Dim idxList() As Long: ReDim idxList(0 To StopAtIndex - 1)
        For i = 0 To StopAtIndex - 1: idxList(i) = i: Next i

        ' Có th? randomize ho?c ch?n logic khác, t?m th?i theo th? t? ngu?c l?i
        For i = StopAtIndex - 1 To 0 Step -1
            If adjRemain = 0 Then Exit For
            If sign > 0 Then
                result(idxList(i)) = result(idxList(i)) + 1
                adjRemain = adjRemain - 1
            ElseIf sign < 0 And result(idxList(i)) > 0 Then
                result(idxList(i)) = result(idxList(i)) - 1
                adjRemain = adjRemain - 1
            End If
        Next i
    End If

    ' Gán các ngày còn l?i = 0
    For i = StopAtIndex To nDays - 1
        result(i) = 0
    Next i

    TaoMangDoLun = result
End Function


Sub TestPhanPhoiLun_RandomK()
    Dim delta As Variant
    Dim cum() As Long
    Dim SDate As Long: SDate = 0
    Dim EDate As Long: EDate = 6
    Dim TongLun As Double: TongLun = 40
    Dim kMin As Double: kMin = 0.2
    Dim kMax As Double: kMax = 0.4
    Dim MaxDaily As Long: MaxDaily = 15
    Dim EarlyStopDays As Long: EarlyStopDays = 1
    Dim i As Long

    delta = TaoMangDoLun(SDate, EDate, TongLun, kMin, kMax, MaxDaily, EarlyStopDays)
    If IsEmpty(delta) Then Exit Sub

    ReDim cum(0 To UBound(delta))
    cum(0) = delta(0)
    For i = 1 To UBound(delta)
        cum(i) = cum(i - 1) + delta(i)
    Next i

    With Sheet4
        .Cells.Clear
        .Cells(1, 1).Resize(1, 3).Value = Array("Ngày", "Lún t?ng ngày", "Lún c?ng d?n")
        For i = 0 To UBound(delta)
            .Cells(i + 2, 1).Value = SDate + i
            .Cells(i + 2, 2).Value = delta(i)
            .Cells(i + 2, 3).Value = cum(i)
        Next i
    End With
End Sub


