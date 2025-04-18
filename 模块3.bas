Attribute VB_Name = "ģ��3"
Sub AddReturnToMainButton()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim btnWidth As Double, btnHeight As Double
    Dim marginRight As Double, marginBottom As Double
    
    ' ��������
    btnWidth = 80      ' ��ť��ȣ���λ������
    btnHeight = 30     ' ��ť�߶�
    marginRight = 20   ' �����ұ߽�ı߾�
    marginBottom = 20  ' �����±߽�ı߾�
    
    ' ��ȡ��ǰ������������������ɵı�����ڹ�����
    Set ws = ActiveSheet
    
    ' ɾ���Ѵ��ڵ�ͬ����ť�������ظ���
    On Error Resume Next
    ws.Shapes("ReturnToMainBtn").Delete
    On Error GoTo 0
    
    ' ���㰴ťλ�ã����½ǣ�
    With ws.UsedRange
        Dim maxRight As Double
        maxRight = .Columns(.Columns.Count).Left + .Columns(.Columns.Count).Width
    End With
    
    Dim btnLeft As Double, btnTop As Double
    btnLeft = 0 'ws.Cells(1, 1).Left + maxRight - btnWidth - marginRight
    btnTop = 0 'ws.Cells(ws.Rows.Count, 1).Top - btnHeight - marginBottom
    
    ' ��Ӱ�ť
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btn
        .name = "ReturnToMainBtn"
        .TextFrame.Characters.Text = "����������"
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .Fill.ForeColor.RGB = RGB(91, 155, 213)  ' ��ɫ����
        .Line.ForeColor.RGB = RGB(0, 0, 0)       ' ��ɫ�߿�
        ' �󶨵���¼�
        .OnAction = "ReturnToMain"
    End With
End Sub

' ����Main������ĺ�
Sub ReturnToMain()
    On Error Resume Next
    Worksheets("Main").Activate
    If Err.Number <> 0 Then
        MsgBox "δ�ҵ�Main������", vbExclamation
    End If
    On Error GoTo 0
End Sub
