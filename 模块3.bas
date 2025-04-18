Attribute VB_Name = "模块3"
Sub AddReturnToMainButton()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim btnWidth As Double, btnHeight As Double
    Dim marginRight As Double, marginBottom As Double
    
    ' 参数设置
    btnWidth = 80      ' 按钮宽度（单位：磅）
    btnHeight = 30     ' 按钮高度
    marginRight = 20   ' 距离右边界的边距
    marginBottom = 20  ' 距离下边界的边距
    
    ' 获取当前活动工作表（假设是新生成的表格所在工作表）
    Set ws = ActiveSheet
    
    ' 删除已存在的同名按钮（避免重复）
    On Error Resume Next
    ws.Shapes("ReturnToMainBtn").Delete
    On Error GoTo 0
    
    ' 计算按钮位置（右下角）
    With ws.UsedRange
        Dim maxRight As Double
        maxRight = .Columns(.Columns.Count).Left + .Columns(.Columns.Count).Width
    End With
    
    Dim btnLeft As Double, btnTop As Double
    btnLeft = 0 'ws.Cells(1, 1).Left + maxRight - btnWidth - marginRight
    btnTop = 0 'ws.Cells(ws.Rows.Count, 1).Top - btnHeight - marginBottom
    
    ' 添加按钮
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btn
        .name = "ReturnToMainBtn"
        .TextFrame.Characters.Text = "返回主界面"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .Fill.ForeColor.RGB = RGB(91, 155, 213)  ' 蓝色背景
        .Line.ForeColor.RGB = RGB(0, 0, 0)       ' 黑色边框
        ' 绑定点击事件
        .OnAction = "ReturnToMain"
    End With
End Sub

' 返回Main工作表的宏
Sub ReturnToMain()
    On Error Resume Next
    Worksheets("Main").Activate
    If Err.Number <> 0 Then
        MsgBox "未找到Main工作表！", vbExclamation
    End If
    On Error GoTo 0
End Sub
