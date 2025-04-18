Attribute VB_Name = "Ä£¿é1"
Function DinttDeltaCpmdT(a As Double, b As Double, c As Double, T As Double) As Double
    DinttDeltaCpmdT = (a * (-298.15 + T)) + (b * (-44.4467 + 0.0005 * T ^ (2))) + (c * (-8.83452 + (3.33333 * 10 ^ -7) * T ^ 3))
End Function

Function DinttDeltaCpmRatioTdT(a, b, c, T)
    DinttDeltaCpmRatioTdT = a * (Log(T / 298.15)) + (b / 1000) * (T - 298.15) + (c / (2 * (10 ^ 6))) * (T ^ 2 - 88893.4225)
End Function

Function DeltaCpm(a As Double, b As Double, c As Double, T As Double) As Double
    DeltaCpm = a + b * 10 ^ -3 * T + c * 10 ^ -6 * T ^ 2
End Function
Function Interval(c As Double, down As Double, up As Double, reg As Double, rtu As Double) As Double
    If (c > (down - reg)) And (c < (up + reg)) Then
        Interval = rtu
    Else
        Interval = -1
    End If
End Function
