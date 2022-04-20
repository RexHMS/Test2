Attribute VB_Name = "Multiple"
Public MyArray(10) As String
Public OF As Integer
Public Z As Integer
Public wtf As Integer
Sub Multiple_Click()

Dim INP(10) As Object
Dim Rename(10) As String
Dim MainSheet As String
Dim Filename As String
Dim Huawei As String

    MultiFile.Show
    If wtf <> 1 Then
        End
    End If
    
    Application.ScreenUpdating = False '不要更新螢幕
    
    N = ActiveWorkbook.Sheets(1).Range("A1")
    'Analytical_options.Show
    'ActiveWorkbook.Sheets(1).Range("A1").Clear
    QQ = "Huawei SNR test"
    
    Worksheets(2).Select
    
    Set WW = Worksheets(2)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            Huawei = 0
            Call Multiple.RV(N)
            Call Multiple.NOISE(N)
            Call Multiple.SNR(N)
        Else
            Huawei = 1
        End If
    
    Worksheets(1).Select
    
    
    ns = N + 1
    Z = 0
    For I = ns To 2 Step -1
        'ActiveWorkbook.Sheets(1).Cells(44, K) = Sheets(I).Name
        QQ = "Huawei SNR test"
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        If WW.Range("I1") = "Nothing" Then
            Z = Z + 1
            WW.Range("I1").Clear
        End If
    Next I
    
    
    If Z = 0 Then
        Call Multiple.HuaweiSNR(N, Huawei)
    End If
    
    Sheets(1).name = "Standard"
    
    Sheets(1).Activate
    
    If Range("b1") > 0 Then
        Sheets.Add before:=ActiveSheet
        Sheets(2).Select
        Sheets(2).Range("b1").Cut Destination:=Sheets(1).Range("b1")
        Sheets(2).Range(Cells(1, 1), Cells(10, 1)).Cut Destination:=Sheets(1).Range("a1")
        N = Sheets(1).Range("A1").Value
        Sheets(1).Select
        Call Multiple.Other(N)
        Range(Cells(1, 1), Cells(15, 1)).Clear
        Range("B1").Clear
        Sheets(1).name = "OtherOptions"
    '/////////////////////////////////////////////////////////////////////////////////////////////
    End If
    
    Range(Cells(1, 1), Cells(15, 1)).Clear
    Range("B1").Clear
    ActiveWorkbook.Save

 End Sub

Sub RV(N)
Dim R(10) As Integer
    ns = N + 1
    ActiveWorkbook.Sheets(1).Range("B2") = "RV"
    ActiveWorkbook.Sheets(1).Range("B3") = "Max"
    ActiveWorkbook.Sheets(1).Range("B4") = "Avg"
    ActiveWorkbook.Sheets(1).Range("B5") = "Min"
    J = 3
    K = J

    L = 2
    M = 1
    O = 1

    For I = ns To 2 Step -1
        ActiveWorkbook.Sheets(1).Cells(2, K) = Sheets(I).name
        QQ = "Ridge-Valley Value"
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            QQ = "Signal(RV)"
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
        A = WW.Range(Cells(first, ZZ), Cells(Last, ZZ))
        WS.Cells(3, K).Value = Application.WorksheetFunction.Max(A)
        WS.Cells(4, K).Value = Application.WorksheetFunction.Average(A)
        WS.Cells(4, K).Value = WorksheetFunction.Round(WS.Cells(4, K), 2)
        WS.Cells(5, K).Value = Application.WorksheetFunction.Min(A)
        R(I) = ZZ
        K = K + 1
    Next

    Sheets(1).Select
    Range(Cells(2, 3), Cells(2, K - 1)).Copy Destination:=Cells(2, K + 2)
    Cells(9, K + 1) = "Total"
        
    Range(Cells(2, 2), Cells(5, K - 1)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
    
    RVL = Application.WorksheetFunction.Min(Range(Cells(3, 3), Cells(5, K - 1)))
    RVH = Application.WorksheetFunction.Max(Range(Cells(3, 3), Cells(5, K - 1)))
        'Range("B6").Value = RVL
        'Range("B7").Value = RVH
    RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 0)
    RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 0)
    P = RVL + RVspacing
    Cells(3, K + 1).Value = RVL & "~" & P
        
    MM = P
    Do Until L > 6
        L = L + 1
        P = P + M
        Q = P + RVspacing
        Cells(L + O, K + 1).Value = P & "~" & Q
        P = Q
    Loop
        
    S = K + 2
        
    For I = ns To 2 Step -1
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        ZZ = R(I)
        Call Multi區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, S, Huawei)
        MM = RVL + RVspacing
    Next
        
    Range(Cells(2, K + 1), Cells(8, K + 1 + N)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range(Cells(2, K + 1), Cells(8, K + 1 + N)), PlotBy:=xlColumns
    ActiveChart.ChartTitle.Text = "RV 分佈圖"
    ActiveChart.Parent.Top = Cells(2, K + 3 + N).Top
    ActiveChart.Parent.Left = Cells(2, K + 3 + N).Left
    
    Cells(2, K + 1) = "RV"
    
    Range(Cells(2, K + 1), Cells(9, K + 1 + N)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
     
'QQ = "Ridge-Valley Value"
'Set WW = Worksheets(NS)
'Set WS = Worksheets(1)
'Call Positioning.Position(WW, QQ, G, SS, ZZ, LAST, FIRST)
'A = WW.Range(Cells(FIRST, ZZ), Cells(LAST, ZZ))
'    WS.Cells(3, 3).Value = Application.WorksheetFunction.Max(A)
'    WS.Cells(4, 3).Value = Application.WorksheetFunction.Average(A)
'    WS.Cells(4, 3).Value = WorksheetFunction.Round(WS.Cells(4, 3), 2)
'    WS.Cells(5, 3).Value = Application.WorksheetFunction.Min(A)

End Sub

Sub NOISE(N)
Dim R(10) As Integer
    ns = N + 1
    ActiveWorkbook.Sheets(1).Range("B16") = "Noise"
    ActiveWorkbook.Sheets(1).Range("B17") = "Max"
    ActiveWorkbook.Sheets(1).Range("B18") = "Avg"
    ActiveWorkbook.Sheets(1).Range("B19") = "Min"
    J = 3
    K = J

    L = 16
    M = 0.001
    O = 1
    
    For I = ns To 2 Step -1
        ActiveWorkbook.Sheets(1).Cells(16, K) = Sheets(I).name
        QQ = "Noise"
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        A = WW.Range(Cells(first, ZZ), Cells(Last, ZZ))
        WS.Cells(17, K).Value = Application.WorksheetFunction.Max(A)
        WS.Cells(18, K).Value = Application.WorksheetFunction.Average(A)
        WS.Cells(18, K).Value = WorksheetFunction.Round(WS.Cells(18, K), 3)
        WS.Cells(19, K).Value = Application.WorksheetFunction.Min(A)
        R(I) = ZZ

        K = K + 1
    Next

    Sheets(1).Select
    Range(Cells(16, 3), Cells(16, K - 1)).Copy Destination:=Cells(16, K + 2)
    Cells(23, K + 1) = "Total"
        
    Range(Cells(16, 2), Cells(19, K - 1)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
        
    RVL = Application.WorksheetFunction.Min(Range(Cells(17, 3), Cells(19, K - 1)))
    RVH = Application.WorksheetFunction.Max(Range(Cells(17, 3), Cells(19, K - 1)))
        'Range("B20").Value = RVL
        'Range("B21").Value = RVH
    RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
    RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
    P = RVL + RVspacing
    Cells(17, K + 1).Value = RVL & "~" & P
        
    MM = P
    Do Until L > 20
        L = L + 1
        P = P + M
        Q = P + RVspacing
        Cells(L + O, K + 1).Value = P & "~" & Q
        P = Q
    Loop
        
    S = K + 2
        
    For I = ns To 2 Step -1
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        ZZ = R(I)
        Call Multi區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, S, Huawei)
        MM = RVL + RVspacing
    Next
        
    Range(Cells(16, K + 1), Cells(22, K + 1 + N)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range(Cells(16, K + 1), Cells(22, K + 1 + N)), PlotBy:=xlColumns
    ActiveChart.ChartTitle.Text = "Noise 分佈圖"
    ActiveChart.Parent.Top = Cells(16, K + 3 + N).Top
    ActiveChart.Parent.Left = Cells(16, K + 3 + N).Left
    
    Cells(16, K + 1) = "Noise"
    
    Range(Cells(16, K + 1), Cells(23, K + 1 + N)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With

End Sub

Sub SNR(N)
Dim R(10) As Integer
    ns = N + 1
    ActiveWorkbook.Sheets(1).Range("B30") = "SNR"
    ActiveWorkbook.Sheets(1).Range("B31") = "Max"
    ActiveWorkbook.Sheets(1).Range("B32") = "Avg"
    ActiveWorkbook.Sheets(1).Range("B33") = "Min"
    J = 3
    K = J

    L = 30
    M = 0.001
    O = 1
    
    For I = ns To 2 Step -1
        ActiveWorkbook.Sheets(1).Cells(30, K) = Sheets(I).name
        QQ = "SNR(RV)"
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            QQ = "SNR"
            Set WW = Worksheets(I)
            Set WS = Worksheets(1)
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
         
        A = WW.Range(Cells(first, ZZ), Cells(Last, ZZ))
        WS.Cells(31, K).Value = Application.WorksheetFunction.Max(A)
        WS.Cells(32, K).Value = Application.WorksheetFunction.Average(A)
        WS.Cells(32, K).Value = WorksheetFunction.Round(WS.Cells(32, K), 3)
        WS.Cells(33, K).Value = Application.WorksheetFunction.Min(A)
        R(I) = ZZ

        K = K + 1

    Next

    Sheets(1).Select
    Range(Cells(30, 3), Cells(30, K - 1)).Copy Destination:=Cells(30, K + 2)
    Cells(37, K + 1) = "Total"
        
    Range(Cells(30, 2), Cells(33, K - 1)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
        
    RVL = Application.WorksheetFunction.Min(Range(Cells(31, 3), Cells(33, K - 1)))
    RVH = Application.WorksheetFunction.Max(Range(Cells(31, 3), Cells(33, K - 1)))
        'Range("B34").Value = RVL
        'Range("B35").Value = RVH
    RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
    RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
    P = RVL + RVspacing
    Cells(31, K + 1).Value = RVL & "~" & P
        
    MM = P
    Do Until L > 34
        L = L + 1
        P = P + M
        Q = P + RVspacing
        Cells(L + O, K + 1).Value = P & "~" & Q
        P = Q
    Loop
        
    S = K + 2
        
    For I = ns To 2 Step -1
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        ZZ = R(I)
        Call Multi區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, S, Huawei)
        MM = RVL + RVspacing
    Next
        
    Range(Cells(30, K + 1), Cells(36, K + 1 + N)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range(Cells(30, K + 1), Cells(36, K + 1 + N)), PlotBy:=xlColumns
    ActiveChart.ChartTitle.Text = "SNR 分佈圖"
    ActiveChart.Parent.Top = Cells(30, K + 3 + N).Top
    ActiveChart.Parent.Left = Cells(30, K + 3 + N).Left

    Cells(30, K + 1) = "SNR"
    
    Range(Cells(30, K + 1), Cells(37, K + 1 + N)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
End Sub

Sub HuaweiSNR(N, Huawei)
Dim R(10) As Integer
    ns = N + 1
    ActiveWorkbook.Sheets(1).Range("B2") = "HuaweiSNR"
    ActiveWorkbook.Sheets(1).Range("B3") = "Max"
    ActiveWorkbook.Sheets(1).Range("B4") = "Avg"
    ActiveWorkbook.Sheets(1).Range("B5") = "Min"
    
    J = 3
    K = J

    L = 2
    M = 0.001
    O = 1
    
    For I = ns To 2 Step -1
        ActiveWorkbook.Sheets(1).Cells(2, K) = Sheets(I).name
        QQ = "Huawei SNR test"
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        If WW.Range("I1") = "" Then
            'WW.Range("I1").Clear
            QQ = "SNR"
            Set WW = Worksheets(I)
            Set WS = Worksheets(1)
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            
            A = WW.Range(Cells(first, ZZ), Cells(Last, ZZ))
            WS.Cells(3, K).Value = Application.WorksheetFunction.Max(A)
            WS.Cells(4, K).Value = Application.WorksheetFunction.Average(A)
            WS.Cells(4, K).Value = WorksheetFunction.Round(WS.Cells(4, K), 3)
            WS.Cells(5, K).Value = Application.WorksheetFunction.Min(A)
            R(I) = ZZ
            K = K + 1
        End If

    Next

    Sheets(1).Select
    Range(Cells(2, 3), Cells(2, K - 1)).Copy Destination:=Cells(2, K + 2)
    Cells(9, K + 1) = "Total"
        
    Range(Cells(2, 2), Cells(5, K - 1)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
        
    RVL = Application.WorksheetFunction.Min(Range(Cells(3, 3), Cells(5, K - 1)))
    RVH = Application.WorksheetFunction.Max(Range(Cells(3, 3), Cells(5, K - 1)))
    RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
    RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
    P = RVL + RVspacing
    Cells(3, K + 1).Value = RVL & "~" & P
        
    MM = P
    Do Until L > 6
        L = L + 1
        P = P + M
        Q = P + RVspacing
        Cells(L + O, K + 1).Value = P & "~" & Q
        P = Q
    Loop
        
    S = K + 2
        
    For I = ns To 2 Step -1
        Set WW = Worksheets(I)
        Set WS = Worksheets(1)
        ZZ = R(I)
        Call Multi區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, S, Huawei)
        MM = RVL + RVspacing
    Next
        
    Range(Cells(2, K + 1), Cells(9, K + 1 + N)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range(Cells(2, K + 1), Cells(8, K + 1 + N)), PlotBy:=xlColumns
    ActiveChart.ChartTitle.Text = "Huawei SNR 分佈圖"
    ActiveChart.Parent.Top = Cells(2, K + 3 + N).Top
    ActiveChart.Parent.Left = Cells(2, K + 3 + N).Left

    Cells(2, K + 1) = "Huawei SNR"
    
    Range(Cells(2, K + 1), Cells(9, K + 1 + N)).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
    End With
    

End Sub

Sub OpFcount()

Dim L1, L2, L3, L4, L5
Dim R1, R2, R3, R4, R5
 
'Dim MyArray(10) As String
    I = 1
    If Analytical_options.CheckBox1.Value = True Then      '使用先前已有群組
        L1 = 1
        MyArray(I) = "Placement Test"
        I = I + 1
        'Me.ComboBox1.Visible = True     '顯示現有群組名稱組合框
        'Me.Label4.Visible = True        '顯示相應說明文字標籤
    Else            '否則
        L1 = 0
        'Me.ComboBox1.Visible = False        '隱藏現有群組名稱組合框
        'Me.Label4.Visible = False       '隱藏相應說明文字標籤

    End If
    
    If Analytical_options.CheckBox2.Value = True Then
        L2 = 1
        MyArray(I) = "Imaging Current Test(VCC)"
        I = I + 1
    Else
        L2 = 0
    End If
    
    If Analytical_options.CheckBox3.Value = True Then
        L3 = 1
        MyArray(I) = "Imaging Current Test(VDD)"
        I = I + 1
    Else
        L3 = 0
    End If

    If Analytical_options.CheckBox4.Value = True Then
        L4 = 1
        MyArray(I) = "FOD Current Test(VCC)"
        I = I + 1
    Else
        L4 = 0
    End If

    If Analytical_options.CheckBox5.Value = True Then
        L5 = 1
        MyArray(I) = " FOD Current Test(VDD)"
        I = I + 1
    Else
        L5 = 0
    End If

    If Analytical_options.CheckBox6.Value = True Then
        R1 = 1
        MyArray(I) = "PowerDown Current Test(VCC)"
        I = I + 1
    Else
        R1 = 0
    End If

    If Analytical_options.CheckBox7.Value = True Then
        R2 = 1
        MyArray(I) = " PowerDown Current Test(VDD)"
        I = I + 1
    Else
        R2 = 0
    End If

    'If FRRForm.CheckBox8.Value = True Then
    '    R3 = 1
    '    MyArray(I) = "PWD Current VCC"
    '    I = I + 1
    'Else
    '    R3 = 0
    'End If

    'If FRRForm.CheckBox9.Value = True Then
    '    R4 = 1
    '    MyArray(I) = "PWD Current VCC"
    '    I = I + 1
    'Else
    '    R4 = 0
    'End If

    OF = L1 + L2 + L3 + L4 + L5 + R1 + R2 '+ R3 + R4
    
End Sub

Sub Multi區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, S, Huawei)
Dim I, J, K, L, M
    WS.Select
    If Huawei = 0 Then
        If QQ = "Signal(RV)" Or QQ = "Ridge-Valley Value" Then
            J = 3
            K = 4
            M = 1
        ElseIf QQ = "Noise" Then
            J = 17
            K = 18
            M = 0.001
        ElseIf QQ = "SNR(RV)" Then
            J = 31
            K = 32
            M = 0.001
        ElseIf QQ = "SNR" Then
            J = 31
            K = 32
            M = 0.001
        End If
    Else
        QQ = "SNR"
        J = 3
        K = 4
        M = 0.001
    End If
    
    WS.Cells(J, S) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), "<=" & MM)
    I = 3
    L = 1
    Do Until I > 7
        Caculate = MM + L * RVspacing + M
        Caculate2 = Application.WorksheetFunction.Sum(WS.Range(Cells(J, S), Cells(K - 1, S)))
        WS.Cells(K, S) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), _
            "<=" & Caculate) - Caculate2
        MM = MM + M
        I = I + 1
        L = L + 1
        K = K + 1
    Loop
    WS.Cells(K, S) = Application.WorksheetFunction.Sum(WS.Range(Cells(J, S), Cells(K - 1, S)))
    S = S + 1

End Sub

Sub Other(N)
Dim I, J, K, L, M, O, P, Q, R, S, T, U, V, X, Y As Integer
Dim A As Range
Dim Z(20) As Integer
    
    ns = N + 2
    R = 2
    L = 2
    X = Range("B1").Value
    V = 1
    Q = 1
    Y = 1
    
    Do While V <= X
        Cells(L, R) = Cells(V + 1, 1).Value
        Cells(L + 1, R) = "Max"
        Cells(L + 2, R) = "Avg"
        Cells(L + 3, R) = "Min"
        J = 3
        K = J

        M = 0.001
        O = 1
    
        For I = ns To 3 Step -1
            ActiveWorkbook.Sheets(1).Cells(L, K) = Sheets(I).name
            QQ = Cells(L, R)
            Set WW = Worksheets(I)
            Set WS = Worksheets(1)
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            Set A = WW.Range(Cells(first, ZZ), Cells(Last, ZZ))
            Set WS = Sheets(1)
            WS.Select
            With WS
                Cells(L + 1, K).Value = Application.WorksheetFunction.Max(A)
                Cells(L + 2, K).Value = Application.WorksheetFunction.Average(A)
                Cells(L + 2, K).Value = WorksheetFunction.Round(WS.Cells(L + 2, K), 3)
                Cells(L + 3, K).Value = Application.WorksheetFunction.Min(A)
            End With
            Z(I) = ZZ

            K = K + 1

        Next I

        If QQ = "Placement Test" Then
            M = 1
        Else
            M = 0.001
        End If
        
        Sheets(1).Select
        Range(Cells(L, 3), Cells(L, K - 1)).Copy Destination:=Cells(L, K + 2)
        Cells(L + 7, K + 1) = "Total"
        Range(Cells(L, 2), Cells(L + 3, K - 1)).Select  '將要複製的範圍先選取起來
        Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        End With
        
        RVL = Application.WorksheetFunction.Min(Range(Cells(L + 1, 3), Cells(L + 3, K - 1)))
        RVH = Application.WorksheetFunction.Max(Range(Cells(L + 1, 3), Cells(L + 3, K - 1)))
        
        If QQ = "Placement Test" Then
            RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 0)
            RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 0)
        Else
            RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
            RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
        End If
        
        P = RVL + RVspacing
        Cells(L + 1, K + 1).Value = RVL & "~" & P
        U = L
        MM = P
        T = L + 4
        Do Until L > T
            L = L + 1
            P = P + M
            Q = P + RVspacing
            Cells(L + O, K + 1).Value = P & "~" & Q
            P = Q
        Loop
        
        S = K + 2
        
        For I = ns To 3 Step -1
            Set WW = Worksheets(I)
            Set WS = Worksheets(1)
            ZZ = Z(I)
            Call MultiOther數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, U, S)
            
            MM = RVL + RVspacing
        Next I
        
        Range(Cells(U, K + 1), Cells(U + 6, K + 1 + N)).Select
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData Source:=Range(Cells(U, K + 1), Cells(U + 6, K + 1 + N)), PlotBy:=xlColumns
        ActiveChart.ChartTitle.Text = QQ & "分佈圖"
        ActiveChart.Parent.Top = Cells(U, K + 3 + N).Top
        ActiveChart.Parent.Left = Cells(U, K + 3 + N).Left
        Cells(U, K + 1) = QQ
    
        Range(Cells(U, K + 1), Cells(U + 7, K + 1 + N)).Select  '將要複製的範圍先選取起來
        Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        End With
        L = U + 14
        V = V + 1
        
    Loop
    
    
End Sub

Sub MultiOther數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, U, S)
Dim I, J, K, L, M
    WS.Select
    I = 16
    If QQ = "Placement Test" Then
        M = 1
    Else
        M = 0.001
    End If
    J = U + 1
    K = U + 2
    WS.Cells(J, S) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), "<=" & MM)
    I = 3
    L = 1
    Do Until I > 7
        Caculate = MM + L * RVspacing + M
        Caculate2 = Application.WorksheetFunction.Sum(WS.Range(Cells(J, S), Cells(K - 1, S)))
        WS.Cells(K, S) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), _
            "<=" & Caculate) - Caculate2
        MM = MM + M
        I = I + 1
        L = L + 1
        K = K + 1
    Loop
    WS.Cells(K, S) = Application.WorksheetFunction.Sum(WS.Range(Cells(J, S), Cells(K - 1, S)))
    S = S + 1
End Sub


Sub CheckOp()
    
    I = Range("A1").Value
    I = I + 1
    For I = I To 1 Step -1
        For OPN = 1 To Multiple.OF
            QQ = Sheets(1).Cells(OPN + 1, 1)
            Set WW = Worksheets(I)
            Set WS = Worksheets(1)
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            If WW.Range("i1").Value = "Nothing" Then
                MsgBox "檔案" & WW.name & "沒有" & QQ & vbNewLine & "請關閉檔案，重新載入!!", 0 + 64
                ActiveWorkbook.Save
                End
            End If
        Next OPN
    Next I
End Sub

