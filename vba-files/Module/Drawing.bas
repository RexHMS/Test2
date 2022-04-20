Attribute VB_Name = "Drawing"
Sub RNS_Distribution(QQ, RN, Hua, HuaSNR)
Dim I, J, K, N, P, R, G, B As Integer
Dim L, M, O As String
    
    If QQ = "Signal(RV)" Or QQ = "Ridge-Valley Value" Then
        I = 2
        J = 7
        K = 2
        L = "RV分佈圖"
        M = "RV分佈圖(%)"
        N = 1
        O = "RV"
        P = 8
        R = 91
        G = 155
        B = 213
    ElseIf QQ = "Noise" Then
        I = 15
        J = 20
        K = 16
        L = "Noise 分佈圖"
        M = "Noise 分佈圖(%)"
        N = 14
        O = "Noise"
        P = 21
        R = 255
        G = 102
        B = 0
    ElseIf Hua = 0 Then
        If HuaSNR = 0 Then
            If QQ = "SNR(RV)" Or QQ = "SNR" Then
                I = 28
                J = 33
                K = 30
                L = "SNR 分佈圖"
                M = "SNR 分佈圖(%)"
                N = 27
                O = "SNR"
                P = 34
                R = 155
                G = 205
                B = 155
            End If
        End If
    
    ElseIf Hua = 1 Then
        'If HuaSNR = 0 Then
        '    If QQ = "SNR(RV)" Or QQ = "SNR" Then
        '        I = 2
        '        J = 7
        '        K = 2
        '        L = "SNR 分佈圖"
        '        M = "SNR 分佈圖(%)"
        '        N = 1
        '        O = "SNR"
        '        P = 8
        '        R = 91
        '        G = 155
        '        B = 213
        '    End If
       
        'ElseIf HuaSNR = 1 Then
        '        I = 15
        '        J = 20
        '        K = 16
        '        L = "Huawei SNR 分佈圖"
        '        M = "Huawei SNR 分佈圖(%)"
        '        N = 14
        '        O = "Huawei SNR"
        '        P = 21
        '        R = 255
        '        G = 102
        '        B = 0
        'End If
        I = 2
        J = 7
        K = 2
        L = "Huawei SNR 分佈圖"
        M = "Huawei SNR 分佈圖(%)"
        N = 1
        O = "RV"
        P = 8
        R = 91
        G = 155
        B = 213
        
    End If
    
    RN.Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=ActiveSheet.Range(Cells(I, 7), Cells(J, 8)), PlotBy:=xlColumns

    ActiveChart.ApplyLayout (5)
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Orientation = xlHorizontal

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = L '抓取圖表的TITLE名稱
    ActiveChart.Parent.Top = Cells(K, 11).Top
    ActiveChart.Parent.Left = Cells(K, 11).Left
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = False 'Y軸名稱隱藏
    ActiveChart.Axes(xlValue).HasMajorGridlines = True    '圖表格線
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    ActiveChart.SeriesCollection(1).Interior.Color = RGB(R, G, B)

    Range("G" & I & ":G" & J & ",I" & I & ":I" & J).Select
    
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range( _
        "RV_Noise!G" & I & ":G" & J & ",RV_Noise!I" & I & ":I" & J)
    ActiveChart.ApplyLayout (5)
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Orientation = xlHorizontal

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = M '抓取圖表的TITLE名稱
    
    Selection.Format.TextFrame2.TextRange.Characters.Text = M
    ActiveChart.Parent.Top = Cells(K, 18).Top
    ActiveChart.Parent.Left = Cells(K, 18).Left
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = False 'Y軸名稱隱藏
    ActiveChart.Axes(xlValue).HasMajorGridlines = True    '圖表格線
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    ActiveChart.SeriesCollection(1).Interior.Color = RGB(R, G, B)
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    RN.Range("G" & N) = O
    RN.Range("H" & N) = "Pcs"
    RN.Range("I" & N) = "%"
    RN.Range("G" & P) = "Total Pcs"
    
    RN.Range("G" & N & ":I" & P).HorizontalAlignment = xlCenter
    
    RN.Columns("G").ColumnWidth = 15.5
    RN.Columns("G").HorizontalAlignment = xlCenter

    Range("G" & N & ":I" & P).Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        End With

End Sub


Sub IC_Form()

    Worksheets.Add ' 新增工作表
    Worksheets(1).name = "IC_information"
    Worksheets("IC_information").Select
    Range("A1:D1").Select
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
    Range("A12:D12").Select   '合併儲存格
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
    Range("A13:D13").Select
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
    Range("A1:D13").Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        
'.LineStyle = xlContinuous '實線
'.Weight = xlThick  '粗線
'.Borders(xlEdgeRight).ColorIndex = 3 '紅色
        
'xlContinuous = 細線
'xlThick = 粗線
'-----------------------------------------------------------------------------------------------
        End With

     ' 將A,B,C,D 欄變寬
    With Worksheets("IC_information").Columns("A")
        .ColumnWidth = .ColumnWidth * 2
    End With
        
    With Worksheets("IC_information").Columns("B")
        .ColumnWidth = .ColumnWidth * 2
    End With
        
    With Worksheets("IC_information").Columns("C")
        .ColumnWidth = .ColumnWidth * 2
    End With
        
    With Worksheets("IC_information").Columns("D")
        .ColumnWidth = .ColumnWidth * 2
    End With
        
    Range("A1").Formula = "IC_information"
          
    With Worksheets("IC_information")
        .Cells(1, 1).Font.Size = 14                         '  設定Title字體大小
        .Cells(1, 1).Font.Bold = True                        '   設定粗體字
    End With
           
    With Worksheets("IC_information")
        .Cells(13, 4).Font.Size = 12         '  設定字體大小
          
        Range("A2").Formula = "Product Name:"
        Range("A3").Formula = "Date:"
        Range("A4").Formula = "Type:"
        Range("A5").Formula = "Die:"
        Range("A6").Formula = "Package Size:"
        Range("A7").Formula = "Mold Clearance:"
        Range("A8").Formula = "Coating house:"
        Range("A9").Formula = "Coating type/ color:"
        Range("A10").Formula = "Coting thickness:"
        Range("A11").Formula = "DK:"
        Range("C2").Formula = "Wafer house:"
        Range("C3").Formula = "Package house:"
        Range("C4").Formula = "Module house:"
        Range("C5").Formula = "Lot No. :"
        Range("C6").Formula = "Data code:"
        Range("C7").Formula = "LDO on/ off:"
        Range("C8").Formula = "Substrate:"
        Range("C9").Formula = "Module connecter:"
        Range("C10").Formula = "Pcs:"
        Range("C11").Formula = "PM:"
        Range("A12").Formula = "Major purpose: "
        Range("A13").Formula = "Remark: "
        .Rows(12).HorizontalAlignment = -4131
        .Rows(13).HorizontalAlignment = -4131
    End With
         
    With Worksheets("IC_information").Cells.Font
        .name = "Calibri"                '  設定Calibri字型
      ' .Size = 8
    End With

End Sub

Sub Report_form()


    Sheets("IC_information").Select    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    Sheets.Add before:=ActiveSheet ' 新增工作表
    Worksheets(1).name = "Report"
    
    Range("A1:B14").Select    '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
    With Selection
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        
        '.LineStyle = xlContinuous '實線
        '.Weight = xlThick  '粗線
        '.Borders(xlEdgeRight).ColorIndex = 3 '紅色
        'xlContinuous = 細線
        'xlThick = 粗線
'-----------------------------------------------------------------------------------------------
    End With

     ' A欄寬設定8.5 , B 欄寬 60
    With Worksheets("Report").Columns("A")
        .ColumnWidth = 14.5
    End With
        
            
    With Worksheets("Report").Columns("B")
        .ColumnWidth = 110
    End With
    
    Range("A3:B3").Select   '合併儲存格
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
        
    With Worksheets("Report").Rows("4")
        .RowHeight = 409.5
    End With
        
    Range("A1").Formula = "Assigner"
         '     .Font.ColorIndex = 7
         
           
    'With Worksheets("IC_information")
     '     .Cells(13, 4).Font.Size = 12         '  設定字體大小
          
    Range("A2").Formula = "Purpose"
    Range("A3").Formula = "Comment"
    Range("A4").Formula = "Result"
    Range("A5").Formula = "Issue"
    Range("A6").Formula = "Criteria"
    Range("A7").Formula = "測試日期"
    Range("A8").Formula = "版號"
    Range("A9").Formula = "參數"
    Range("A10").Formula = "Arduino Voltage"
    Range("A11").Formula = "Arduino Numbe"
    Range("A12").Formula = "File link"
    Range("A13").Formula = "Next Step"
    Range("A14").Formula = "測試人員"
          
         ' .Rows(13).HorizontalAlignment = -4131
         
     '   End With

    Range("A1:A14").Interior.ColorIndex = 49
    Range("A1:A14").Font.ColorIndex = 2
    Range("A1:A14").HorizontalAlignment = xlCenter
  '  With Worksheets("IC_information").Cells.Font
  '  .Font.ColorIndex = 7
     '  .Name = "Calibri"                '  設定Calibri字型
      ' .Size = 8
  '   End With
End Sub

Sub ImangeCurrentMap(cs, B)

    cs.Select
    If B = "Imaging Current Test(VCC)" Or B = "Imaging Current Test(3.3V)" Then
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlColumnClustered
        ActiveChart.SetSourceData Source:=ActiveSheet.Range("H3:i8"), PlotBy:=xlColumns
        C = 2
    Else
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlColumnClustered
        ActiveChart.SetSourceData Source:=ActiveSheet.Range("H17:i22"), PlotBy:=xlColumns
        C = 16
    End If

    ActiveChart.ApplyLayout (5)
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Orientation = xlHorizontal

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = B & " 分佈圖" '抓取圖表的TITLE名稱
    ActiveChart.Parent.Top = Range("L" & C).Top
    ActiveChart.Parent.Left = Range("L" & C).Left
 
    If B = "Imaging Current Test(VDD)" Or B = "Imaging Current Test(1.8V)" Then
        ActiveChart.SeriesCollection(1).Interior.Color = RGB(255, 102, 0)
    End If
 
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = False 'Y軸名稱隱藏
    ActiveChart.Axes(xlValue).HasMajorGridlines = True    '圖表格線
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
 'ActiveChart.SeriesCollection(1).Interior.Color = 37


'*******************************************************************************************************************************


    If B = "Imaging Current Test(VCC)" Or B = "Imaging Current Test(3.3V)" Then
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData Source:=Range( _
        "H2:H8,J2:J8")
    Else
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData Source:=Range( _
        "H17:H22,J17:J22")
    End If

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = B & " 分佈圖(%)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = B & " 分佈圖(%)"
    ActiveChart.Parent.Top = Range("S" & C).Top
    ActiveChart.Parent.Left = Range("S" & C).Left
    If B = "Imaging Current Test(VDD)" Or B = "Imaging Current Test(1.8V)" Then
        ActiveChart.SeriesCollection(1).Interior.Color = RGB(255, 102, 0)
    End If
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    With Selection.Format.TextFrame2.TextRange.Characters(1, 8).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
End Sub

Sub FODCurrentMap(cs, B)

    cs.Select
    If B = "FOD Current Test(VCC)" Or B = "FOD Current Test(3.3V)" Then
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlColumnClustered
        ActiveChart.SetSourceData Source:=ActiveSheet.Range("H31:i36"), PlotBy:=xlColumns
        C = 30
    Else
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlColumnClustered
        ActiveChart.SetSourceData Source:=ActiveSheet.Range("H44:i49"), PlotBy:=xlColumns
        C = 44
    End If
    ActiveChart.ApplyLayout (5)
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Orientation = xlHorizontal

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = B & " 分佈圖" '抓取圖表的TITLE名稱
    ActiveChart.Parent.Top = Range("L" & C).Top
    ActiveChart.Parent.Left = Range("L" & C).Left
 
    If B = " FOD Current Test(VDD)" Or B = " FOD Current Test(1.8V)" Then
        ActiveChart.SeriesCollection(1).Interior.Color = RGB(255, 102, 0)
    End If
    
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = False 'Y軸名稱隱藏
    ActiveChart.Axes(xlValue).HasMajorGridlines = True    '圖表格線
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    'ActiveChart.SeriesCollection(1).Interior.Color = 37

'*******************************************************************************************************************************

    If B = "FOD Current Test(VCC)" Or B = "FOD Current Test(3.3V)" Then
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData Source:=Range( _
        "H31:H36,J31:J36")
    Else
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData Source:=Range( _
        "H44:H49,J44:J49")
    End If

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = B & " 分佈圖(%)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = B & " 分佈圖(%)"
    ActiveChart.Parent.Top = Range("S" & C).Top
    ActiveChart.Parent.Left = Range("S" & C).Left
    If B = " FOD Current Test(VDD)" Or B = " FOD Current Test(1.8V)" Then
        ActiveChart.SeriesCollection(1).Interior.Color = RGB(255, 102, 0)
    End If
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    With Selection.Format.TextFrame2.TextRange.Characters(1, 8).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
End Sub

Sub PowerDownCurrentMap(cs, B)

Dim I, J

    cs.Select
    If B = "PowerDown Current Test(VCC)" Or B = "PowerDown Current Test(3.3V)" And cs.Range("a1") = "有Imaging Current Test" Then
        If cs.Range("A2") = "無FOD Current Test" Then
            I = 31
            J = 35
            ActiveSheet.Shapes.AddChart.Select
            ActiveChart.ChartType = xlColumnClustered
            ActiveChart.SetSourceData Source:=ActiveSheet.Range("H" & I & ": I" & J), PlotBy:=xlColumns
            C = 30
        Else
            I = 59
            J = 64
            ActiveSheet.Shapes.AddChart.Select
            ActiveChart.ChartType = xlColumnClustered
            ActiveChart.SetSourceData Source:=ActiveSheet.Range("H" & I & ": I" & J), PlotBy:=xlColumns
            C = 58
        End If
    ElseIf B = " PowerDown Current Test(VDD)" Or B = " PowerDown Current Test(1.8V)" And cs.Range("a1") = "有Imaging Current Test" Then
        If cs.Range("A2") = "無FOD Current Test" Then
            I = 44
            J = 49
            ActiveSheet.Shapes.AddChart.Select
            ActiveChart.ChartType = xlColumnClustered
            ActiveChart.SetSourceData Source:=ActiveSheet.Range("H" & I & ": I" & J), PlotBy:=xlColumns
            C = 44
        Else
            I = 72
            J = 77
            ActiveSheet.Shapes.AddChart.Select
            ActiveChart.ChartType = xlColumnClustered
            ActiveChart.SetSourceData Source:=ActiveSheet.Range("H" & I & ": I" & J), PlotBy:=xlColumns
            C = 72
    
        End If
    End If

    ActiveChart.ApplyLayout (5)
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Orientation = xlHorizontal

    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = B & " 分佈圖" '抓取圖表的TITLE名稱
    ActiveChart.Parent.Top = Range("L" & C).Top
    ActiveChart.Parent.Left = Range("L" & C).Left
    If B = " PowerDown Current Test(VDD)" Or B = " PowerDown Current Test(1.8V)" Then
        ActiveChart.SeriesCollection(1).Interior.Color = RGB(255, 102, 0)
    End If
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = False 'Y軸名稱隱藏
    ActiveChart.Axes(xlValue).HasMajorGridlines = True    '圖表格線
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    'ActiveChart.SeriesCollection(1).Interior.Color = 37
'*******************************************************************************************************************************
    If B = "PowerDown Current Test(VCC)" Or B = "PowerDown Current Test(3.3V)" And cs.Range("a1") = "有Imaging Current Test" Then
        If cs.Range("A2") = "無FOD Current Test" Then
            I = 31
            J = 35
            ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
            ActiveChart.SetSourceData Source:=Range("H" & I & ": H" & J & " , J" & I & ": J" & J)
        Else
            I = 59
            J = 64
            ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
            ActiveChart.SetSourceData Source:=Range("H" & I & ": H" & J & " , J" & I & ": J" & J)
        End If
 '///////////////////////////////////////////////////////////////////////////////////////////////////
    ElseIf B = " PowerDown Current Test(VDD)" Or B = " PowerDown Current Test(1.8V)" And cs.Range("a1") = "有Imaging Current Test" Then
        If cs.Range("A2") = "無FOD Current Test" Then
            I = 44
            J = 49
            ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
            ActiveChart.SetSourceData Source:=Range("H" & I & ": H" & J & " , J" & I & ": J" & J)
        Else
            I = 72
            J = 77
            ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
            ActiveChart.SetSourceData Source:=Range("H" & I & ": H" & J & " , J" & I & ": J" & J)
        End If
    End If
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = B & " 分佈圖(%)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = B & " 分佈圖(%)"
    ActiveChart.Parent.Top = Range("S" & C).Top
    ActiveChart.Parent.Left = Range("S" & C).Left
    If B = " PowerDown Current Test(VDD)" Or B = " PowerDown Current Test(1.8V)" Then
        ActiveChart.SeriesCollection(1).Interior.Color = RGB(255, 102, 0)
    End If
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableNone)
    With Selection.Format.TextFrame2.TextRange.Characters(1, 8).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
End Sub


Sub Current_form(WS, B)

Dim I, J, K, L, M, N, O, P, R

    If B = "Imaging Current Test(VDD)" Or B = "Imaging Current Test(1.8V)" Then
        I = 2
        J = 5
        K = 2
        L = 9
        M = 16
        N = 23
        O = 24.4
        P = 24.4
        R = 24.4

    ElseIf B = " FOD Current Test (VDD)" Or B = " FOD Current Test (1.8V)" And WS.Range("a1") = "無Imaging Current Test" Then
        I = 2
        J = 5
        K = 2
        L = 9
        M = 16
        N = 23
        O = 24.4
        P = 24.4
        R = 24.4
        
    ElseIf B = " FOD Current Test (VDD)" Or B = "FOD Current Test (1.8V)" And WS.Range("a1") = "有Imaging Current Test" Then
        I = 30
        J = 33
        K = 30
        L = 37
        M = 43
        N = 50
        O = 24.4
        P = 24.4
        R = 24.4

    ElseIf B = " PowerDown Current Test (VDD)" Or B = " PowerDown Current Test (1.8V)" And WS.Range("a1") = "無Imaging Current Test" And WS.Range("A2") = "無FOD Current Test" Then
        I = 2
        J = 5
        K = 2
        L = 9
        M = 16
        N = 23
        O = 27.4
        P = 28.4
        R = 28.4

    ElseIf B = " PowerDown Current Test (VDD)" Or B = " PowerDown Current Test (1.8V)" And WS.Range("a1") = "有Imaging Current Test" And WS.Range("A2") = "無FOD Current Test" Then
        I = 30
        J = 33
        K = 30
        L = 37
        M = 43
        N = 50
        O = 27.4
        P = 28.4
        R = 28.4

    ElseIf B = " PowerDown Current Test (VDD)" Or B = " PowerDown Current Test (1.8V)" And WS.Range("a1") = "有Imaging Current Test" And WS.Range("A2") = "有FOD Current Test" Then
        I = 58
        J = 61
        K = 58
        L = 65
        M = 71
        N = 78
        O = 27.4
        P = 28.4
        R = 28.4

    End If

    WS.Select
    WS.Columns("A").ColumnWidth = 20.63
    WS.Columns("D").ColumnWidth = O
    WS.Columns("E").ColumnWidth = P
    WS.Columns("H").ColumnWidth = R

    Range("c" & I, "e" & J).Select '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        End With
    Range("H" & K, "j" & L).Select '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        End With
    Range("H" & M, "j" & N).Select '將要複製的範圍先選取起來
    Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
        End With

End Sub
