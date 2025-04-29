Attribute VB_Name = "模块4"
Sub GenerateAdvancedPhaseDiagramW()
    Dim wsMain As Worksheet, wsData As Worksheet, cht As Chart
    Dim critTemp As Double, critPress As Double
    Dim tripleT As Double, tripleP As Double
    Dim deltaHfus0 As Double, deltaVfus As Double
    Dim deltaHvap0 As Double, deltaHsub0 As Double, deltaVvap As Double
    Dim CpA As Double, CpB As Double, CpC As Double
    Dim R As Double: R = 8.314
    On Error GoTo ErrorHandler
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    Set wsData = ThisWorkbook.Sheets("Data")
    
    ' 从Main表获取物性参数
    With wsMain
        tripleT = .Range("H5").Value    ' 三相点温度
        tripleP = .Range("I5").Value    ' 三相点压力
        deltaHfus0 = .Range("E8").Value ' 基准熔化焓
        deltaVfus = .Range("G8").Value * 10 ^ -6  ' 熔化体积变(m3/mol)
        deltaHvap0 = .Range("H8").Value ' 基准气化焓
        critTemp = .Range("J5").Value   ' 临界温度
        critPress = .Range("K5").Value  ' 临界压力
        deltaVvap = .Range("J8").Value
        
        
        ' 获取热容系数
        CpA = .Range("E5").Value        ' a参数
        CpB = .Range("F5").Value * 0.001 ' b参数转换为J/(mol·K2)
        CpC = .Range("G5").Value * 0.000001 ' c参数转换为J/(mol·K3)
    End With
    
    ' 创建数据表
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("PhaseData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Dim wsChart As Worksheet
    Set wsChart = Sheets.Add
    wsChart.name = "PhaseData"
    
    ' 生成温度数组 (0K - 三相点×1.2)
    Dim T() As Double, P_sl() As Double, P_lg() As Double, P_sg() As Double
    Dim n As Long: n = 2000
    Dim T_min As Double: T_min = 1
    Dim T_max As Double: T_max = critTemp * 1.3
    ReDim T(1 To n), P_sl(1 To n), P_lg(1 To n), P_sg(1 To n)
    Dim min_T As Double
    
    If (tripleT = 273.15) Then min_T = tripleT - 25 Else: min_T = tripleT
    
    
    Dim stepT As Double
    stepT = (T_max - T_min) / (n - 1)
    
    deltaHvapP = deltaHvap0 + CpA * (critTemp - tripleT) + _
                   (CpB / 2) * (critTemp ^ 2 - tripleT ^ 2) + _
                   (CpC / 3) * (critTemp ^ 3 - tripleT ^ 3)
     Dim rt_const As Double: rt_const = tripleP * Exp(-deltaHvapP / R * (1 / critTemp - 1 / tripleT))
    
    ' 计算温度相关焓值
    Dim i As Long
    For i = 1 To n
        T(i) = T_min + (i - 1) * stepT
        
        ' 计算ΔH(T) = ΔH0 + ∫ΔCp dT
        Dim deltaHfus As Double, deltaHvap As Double
        If T(i) >= tripleT Then
            deltaHfus = deltaHfus0 + CpA * (T(i) - tripleT) + _
                       (CpB / 2) * (T(i) ^ 2 - tripleT ^ 2) + _
                       (CpC / 3) * (T(i) ^ 3 - tripleT ^ 3)
        Else
            deltaHfus = deltaHfus0 ' 低于三相点时不考虑固态变化
        End If
        
        deltaHvap = deltaHvap0 + CpA * (T(i) - tripleT) + _
                   (CpB / 2) * (T(i) ^ 2 - tripleT ^ 2) + _
                   (CpC / 3) * (T(i) ^ 3 - tripleT ^ 3)
        ' 固-液线 (克拉伯龙方程)
        If T(i) > min_T And T(i) < critTemp Then
            ' 考虑ΔVfus为负值：dP/dT = ΔH/(TΔV)
            P_sl(i) = CalculatePressure(T(i), tripleT, tripleP, deltaHfus, deltaVfus, CpA, CpB, CpC, tripleT)
            If (T(i) > tripleT - 2 And T(i) < tripleT + 2) Then
                P_sl(i) = tripleP
            End If
            If P_sl(i) < 0 Then P_sl(i) = 0
        Else
            P_sl(i) = 0
        End If
        
        ' 液-气线 (克劳修斯-克拉伯龙方程)
        If T(i) > tripleT And T(i) < critTemp Then
            P_lg(i) = tripleP * Exp(-deltaHvap / R * (1 / T(i) - 1 / tripleT))
        Else
            P_lg(i) = 0
        End If
        
        ' 固-气线 (升华线)
        If T(i) <= tripleT Then
            deltaHsub0 = deltaHfus0 + deltaHvap0
            P_sg(i) = tripleP * Exp(-deltaHsub0 / R * (1 / T(i) - 1 / tripleT))
        Else
            P_sg(i) = 0
        End If
        
    Next i
    
    ' 写入数据
    wsChart.Range("A1:D1") = Array("Temperature", "Solid-Liquid", "Liquid-Gas", "Solid-Gas")
    For i = 1 To n
        wsChart.Cells(i + 1, 1) = T(i)
        wsChart.Cells(i + 1, 2) = P_sl(i)
        wsChart.Cells(i + 1, 3) = P_lg(i)
        wsChart.Cells(i + 1, 4) = P_sg(i)
    Next i
    
    ' 创建图表
    Set cht = wsChart.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Chart
    With cht
        .SetSourceData Source:=wsChart.Range("A1:D" & n + 1)
        .HasTitle = True
        .ChartTitle.Text = wsMain.Range("C3").Value & " 相图（理论计算）"
        .Parent.Width = 500 * 2
        .Parent.Height = 300 * 2
        .Parent.Left = 0
        .Parent.Top = 0
        ' 设置坐标轴
        With .Axes(xlCategory)
            .MinimumScale = T_min
            .MaximumScale = T_max
            .HasTitle = True
            .AxisTitle.Text = "Temperature (K)"
        End With
        
        With .Axes(xlValue)
            .ScaleType = xlScaleLogarithmic
            .MinimumScale = 1
            .MaximumScale = 10000000000#
            .HasTitle = True
            .AxisTitle.Text = "Pressure (Pa)"
            .TickLabels.NumberFormat = "0.E+00"
        End With
        
        ' 设置曲线格式
        FormatSeries .FullSeriesCollection(1), "Solid-Liquid", RGB(0, 112, 192)
        FormatSeries .FullSeriesCollection(2), "Liquid-Gas", RGB(255, 0, 0)
        FormatSeries .FullSeriesCollection(3), "Solid-Gas", RGB(0, 176, 80)
        
        ' 添加关键点
        AddCriticalPoint cht, "Triple Point", tripleT, tripleP, 8, RGB(0, 0, 0)
        AddCriticalPoint cht, "Critical Point", critTemp, critPress, 8, RGB(255, 0, 255)
        If (wsMain.Range("D11").Value <> 0 And wsMain.Range("D12").Value <> 0) Then AddCriticalPoint cht, "MyPoint1", wsMain.Range("D11").Value, wsMain.Range("D12").Value, 8, RGB(65, 54, 186)
        
        
        If (wsMain.Range("E11").Value <> 0 And wsMain.Range("E12").Value <> 0) Then AddCriticalPoint cht, "MyPoint2", wsMain.Range("E11").Value, wsMain.Range("E12").Value, 8, RGB(65, 54, 186)
        
        
        If (wsMain.Range("F11").Value <> 0 And wsMain.Range("F12").Value <> 0) Then AddCriticalPoint cht, "MyPoint3", wsMain.Range("F11").Value, wsMain.Range("F12").Value, 8, RGB(65, 54, 186)
        
        
        If (wsMain.Range("G11").Value <> 0 And wsMain.Range("G12").Value <> 0) Then AddCriticalPoint cht, "MyPoint4", wsMain.Range("G11").Value, wsMain.Range("G12").Value, 8, RGB(65, 54, 186)
        
        
        If (wsMain.Range("H11").Value <> 0 And wsMain.Range("H12").Value <> 0) Then AddCriticalPoint cht, "MyPoint5", wsMain.Range("H11").Value, wsMain.Range("H12").Value, 8, RGB(65, 54, 186)
        
        
    End With
    
    AddReturnToMainButton
    ' 移动图表
    wsChart.Move After:=Sheets(Sheets.Count)
    Exit Sub
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' 辅助函数：格式化曲线
Private Sub FormatSeries(series As series, name As String, color As Long)
    With series
        .name = name
        .Format.Line.ForeColor.RGB = color
        .Format.Line.Weight = 2
        .Smooth = True
    End With
End Sub

' 辅助函数：添加关键点
Private Sub AddCriticalPoint(cht As Chart, name As String, x As Double, y As Double, size As Long, color As Long)
    cht.SeriesCollection.NewSeries
    With cht.FullSeriesCollection(cht.FullSeriesCollection.Count)
        .name = name
        .Values = y
        .XValues = x
        .ChartType = xlXYScatter
        .MarkerStyle = 8
        .MarkerSize = size
        .Format.Fill.ForeColor.RGB = color
    End With
End Sub

Function CalculatePressure(T As Double, T1 As Double, P1 As Double, _
                           deltaHfus0 As Double, deltaVfus As Double, _
                           CpA As Double, CpB As Double, CpC As Double, _
                           T0 As Double) As Double
    ' 计算压强 P(T) 的函数
    ' 使用用户提供的变量和公式

    ' 计算各项
    Dim term1 As Double
    term1 = deltaHfus0 * Log(T / T1)
    
    Dim term2 As Double
    term2 = CpA * (T - T1 - T0 * Log(T / T1))
    
    Dim term3 As Double
    term3 = (CpB * 10 ^ -3) / 4 * (T ^ 2 - T1 ^ 2) - (CpB * 10 ^ -3) / 2 * T0 ^ 2 * Log(T / T1)
    
    Dim term4 As Double
    term4 = (CpC * 10 ^ -6) / 9 * (T ^ 3 - T1 ^ 3) - (CpC * 10 ^ -6) / 3 * T0 ^ 3 * Log(T / T1)
    
    ' 计算总压强
    CalculatePressure = P1 + (term1 + term2 + term3 + term4) / deltaVfus
End Function

