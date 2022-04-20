Attribute VB_Name = "Count"
Sub RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)
Dim I As Integer
Dim J, K, L, M As Integer
Dim A

    If QQ = "Signal(RV)" Or QQ = "Ridge-Valley Value" Then
        WS.Range("C2") = "RV"
        I = 3
        J = 2
        K = 3
        L = 4
        M = 5
    ElseIf QQ = "Noise" Then
        WS.Range("D2") = "Noise"
        I = 4
        J = 2
        K = 3
        L = 4
        M = 5
        
    ElseIf Hua = 0 Then
        If HuaSNR = 0 Then
            If QQ = "SNR(RV)" Or QQ = "SNR" Then
                WS.Range("E2") = "SNR"
                I = 5
                J = 2
                K = 3
                L = 4
                M = 5
            End If
        End If
    
    ElseIf Hua = 1 Then
        'If HuaSNR = 0 Then
        '    If QQ = "SNR(RV)" Or QQ = "SNR" Then
        '        WS.Range("C2") = "SNR"
        '        I = 3
        '        J = 2
        '        K = 3
        '        L = 4
        '        M = 5
        '    End If
        'ElseIf HuaSNR = 1 Then
        '    WS.Range("D2") = "Huawei SNR"
        '    I = 4
        '    J = 2
        '    K = 3
        '    L = 4
        '    M = 5
        WS.Range("C2") = "Huawei SNR"
        I = 3
        J = 2
        K = 3
        L = 4
        M = 5
        
        'End If
    End If
    
    A = WW.Range(Cells(II, ZZ), Cells(Last, ZZ))
    WS.Cells(K, I).Value = Application.WorksheetFunction.Max(A)
    WS.Cells(L, I).Value = Application.WorksheetFunction.Average(A)
    WS.Cells(L, I).Value = WorksheetFunction.Round(WS.Cells(L, I), 2)
    WS.Cells(M, I).Value = Application.WorksheetFunction.Min(A)

End Sub

Sub RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)
'************************************************************************************************************************

    If QQ = "Signal(RV)" Or QQ = "Ridge-Valley Value" Then
        RVL = RN.Range("C5").Value
        RVH = RN.Range("C3").Value
        I = 1
        J = 1
        M = 1
        RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 0)
        RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 0)
        K = RVL + RVspacing
        RN.Range("G2").Value = RVL & "~" & K
    ElseIf QQ = "Noise" Then
        RVL = RN.Range("D5").Value
        RVH = RN.Range("D3").Value
        I = 1
        J = 0.001
        M = 14
        RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
        RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
        K = RVL + RVspacing
        RN.Range("G15").Value = RVL & "~" & K
        
    ElseIf Hua = 0 Then
        If HuaSNR = 0 Then
            If QQ = "SNR(RV)" Or QQ = "SNR" Then
                RVL = RN.Range("E5").Value
                RVH = RN.Range("E3").Value
                I = 1
                J = 0.001
                M = 27
                RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
                RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
                K = RVL + RVspacing
                RN.Range("G28").Value = RVL & "~" & K
            End If
        End If


    ElseIf Hua = 1 Then
        'If HuaSNR = 0 Then
        '    If QQ = "SNR(RV)" Or QQ = "SNR" Then
        '        RVL = RN.Range("C5").Value
        '        RVH = RN.Range("C3").Value
        '        I = 1
        '        J = 0.001
        '        M = 1
        '        RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
        '        RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
        '        K = RVL + RVspacing
        '        RN.Range("G2").Value = RVL & "~" & K
        '    End If
        'ElseIf HuaSNR = 1 Then
        '    RVL = RN.Range("D5").Value
        '    RVH = RN.Range("D3").Value
        '    I = 1
        '    J = 0.001
        '    M = 14
        '    RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
        '    RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
        '    K = RVL + RVspacing
        '    RN.Range("G15").Value = RVL & "~" & K
        'End If
        
        'With RN.Range("B2:D5").Borders
        '.LineStyle = 6
        '.ColorIndex = 1
        '.Weight = 2
        'End With
        
            RVL = RN.Range("C5").Value
            RVH = RN.Range("C3").Value
            I = 1
            J = 0.001
            M = 1
            RVspacing = WorksheetFunction.Round((RVH - RVL) / 6, 3)
            RVspacing = WorksheetFunction.RoundUp((RVH - RVL) / 6, 3)
            K = RVL + RVspacing
            RN.Range("G2").Value = RVL & "~" & K
        'End If
        
        With RN.Range("B2:C5").Borders
        .LineStyle = 6
        .ColorIndex = 1
        .Weight = 2
        End With
        
    End If
    
    If Hua = 0 Then
        With RN.Range("B2:E5").Borders
            .LineStyle = 6
            .ColorIndex = 1
            .Weight = 2
        End With
    Else
        With RN.Range("B2:C5").Borders
            .LineStyle = 6
            .ColorIndex = 1
            .Weight = 2
        End With
    End If

    MM = K
'RV LOOP

    Do Until I > 5
        I = I + 1
        K = K + J
        L = K + RVspacing
        RN.Cells(I + M, 7).Value = K & "~" & L
        K = L
    Loop
End Sub

Sub 區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
Dim I, J, K, L, M
    WS.Select
    If QQ = "Signal(RV)" Or QQ = "Ridge-Valley Value" Then
        J = 2
        K = 3
        M = 1
    ElseIf QQ = "Noise" Then
        J = 15
        K = 16
        M = 0.001
        
    ElseIf Hua = 0 Then
        If HuaSNR = 0 Then
            If QQ = "SNR(RV)" Or QQ = "SNR" Then
                J = 28
                K = 29
                M = 0.001
            End If
        End If
        
    ElseIf Hua = 1 Then
        'If HuaSNR = 0 Then
        '    If QQ = "SNR(RV)" Or QQ = "SNR" Then
        '        J = 2
        '        K = 3
        '        M = 0.001
        '    End If
        'ElseIf HuaSNR = 1 Then
        '    J = 15
        '    K = 16
        '    M = 0.001
        'End If
        J = 2
        K = 3
        M = 1
        
    End If
        
    WS.Cells(J, 8) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), "<=" & MM)
    I = 3
    L = 1
    Do Until I > 7
        Caculate = MM + L * RVspacing + M
        Caculate2 = Application.WorksheetFunction.Sum(WS.Range(Cells(J, 8), Cells(K - 1, 8)))
        WS.Cells(K, 8) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), _
            "<=" & Caculate) - Caculate2
        MM = MM + M
        I = I + 1
        L = L + 1
        K = K + 1
    Loop

'計算ICT3V 區間數量/////////////////////////////////////////////////////////////////////////////////////////////////////////

    WS.Cells(K, 8) = Application.WorksheetFunction.Sum(WS.Range(Cells(J, 8), Cells(K - 1, 8)))
    If QQ = "Signal(RV)" Or QQ = "Ridge-Valley Value" Then
        J = 2
        K = 8
    ElseIf QQ = "Noise" Then
        J = 15
        K = 21
    ElseIf Hua = 0 Then
        If HuaSNR = 0 Then
            If QQ = "SNR(RV)" Or QQ = "SNR" Then
                J = 28
                K = 34
            End If
        End If
    
    ElseIf Hua = 1 Then
        'If HuaSNR = 0 Then
        '    If QQ = "SNR(RV)" Or QQ = "SNR" Then
        '        J = 2
        '        K = 8
        '    End If
        'ElseIf HuaSNR = 1 Then
        '    J = 15
        '    K = 21
        'End If
        J = 2
        K = 8
    End If
      
    I = 1
    L = J
    Do Until I > 7
        WS.Cells(J, 9) = WS.Cells(J, 8) / WS.Cells(K, 8)
        J = J + 1
        I = I + 1
    Loop
    WS.Range(Cells(L, 9), Cells(K, 9)).Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"
End Sub

Sub Current區間數量(QQ, WW, WS, II, III, Last, MM, ZZ, ICT3Vspacing)
Dim I, J, K, L, M, O, P, Q, R As Integer

    WS.Select
    If QQ = "Imaging Current Test(3.3V)" Or QQ = "Imaging Current Test(VCC)" Then
        J = 3
        K = 4
        M = 0.001
        O = 3
        P = 9
    ElseIf QQ = "Imaging Current Test(1.8V)" Or QQ = "Imaging Current Test(VDD)" Then
        J = 17
        K = 18
        M = 0.001
        O = 17
        P = 23
    ElseIf II = "無Imaging Current Test" And QQ = "FOD Current Test(VCC)" Or QQ = "FOD Current Test(3.3V)" Then
        J = 3
        K = 4
        M = 0.001
        O = 3
        P = 9
    ElseIf II = "無Imaging Current Test" And QQ = " FOD Current Test(VDD)" Or QQ = " FOD Current Test(1.8V)" Then
        J = 17
        K = 18
        M = 0.001
        O = 17
        P = 23
    ElseIf II = "有Imaging Current Test" And QQ = "FOD Current Test(VCC)" Or QQ = "FOD Current Test(3.3V)" Then
        J = 31
        K = 32
        M = 0.001
        O = 31
        P = 37
    ElseIf II = "有Imaging Current Test" And QQ = " FOD Current Test(VDD)" Or QQ = " FOD Current Test(1.8V)" Then
        J = 44
        K = 45
        M = 0.001
        O = 44
        P = 50
    ElseIf II = "無Imaging Current Test" And III = "無FOD Current Test" And QQ = "PowerDown Current Test(VCC)" Or QQ = "PowerDown Current Test(3.3V)" Then
        J = 3
        K = 4
        M = 0
        O = 3
        P = 9
    ElseIf II = "無Imaging Current Test" And III = "無FOD Current Test" And QQ = " PowerDown Current Test(VDD)" Or QQ = " PowerDown Current Test(1.8V)" Then
        J = 17
        K = 18
        M = 0.001
        O = 17
        P = 23
    ElseIf II = "有Imaging Current Test" And III = "無FOD Current Test" And QQ = "PowerDown Current Test(VCC)" Or QQ = "PowerDown Current Test(3.3V)" Then
        J = 31
        K = 32
        M = 0.001
        O = 31
        P = 37
    ElseIf II = "有Imaging Current Test" And III = "無FOD Current Test" And QQ = " PowerDown Current Test(VDD)" Or QQ = " PowerDown Current Test(1.8V)" Then
        J = 44
        K = 45
        M = 0.001
        O = 44
        P = 50
    ElseIf II = "有Imaging Current Test" And III = "有FOD Current Test" And QQ = "PowerDown Current Test(VCC)" Or QQ = "PowerDown Current Test(3.3V)" Then
        J = 59
        K = 60
        M = 0.001
        O = 59
        P = 65
    ElseIf II = "有Imaging Current Test" And III = "有FOD Current Test" And QQ = " PowerDown Current Test(VDD)" Or QQ = " PowerDown Current Test(1.8V)" Then
        J = 72
        K = 73
        M = 0.001
        O = 72
        P = 78
    End If
        
    WS.Cells(J, 9) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), _
             "<=" & MM)
    I = 3
    L = 1
    Do Until I > 7
        Caculate = MM + L * ICT3Vspacing + M
        Caculate2 = Application.WorksheetFunction.Sum(WS.Range(Cells(J, 9), Cells(K - 1, 9)))
        WS.Cells(K, 9) = Application.WorksheetFunction.CountIf(WW.Columns(ZZ), _
        "<=" & Caculate) - Caculate2
        MM = MM + M
        I = I + 1
        L = L + 1
        K = K + 1
    Loop

'計算ICT3V 區間數量/////////////////////////////////////////////////////////////////////////////////////////////////////////
    WS.Cells(K, 9) = Application.WorksheetFunction.Sum(WS.Range(Cells(J, 9), Cells(K - 1, 9)))
    Q = 1
    R = O
    Do Until Q > 7
        WS.Cells(O, 10) = WS.Cells(O, 9) / WS.Cells(P, 9)
        O = O + 1
        Q = Q + 1
    Loop
    WS.Range(Cells(R, 10), Cells(P, 10)).Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.0%"
End Sub

Sub HWSW(II, BinSS, BinZZ, BinLast, BINAME)

    Set HSB = Worksheets("HW_SW_BIN")
    Set all = Worksheets("all_log")
    HSB.Select
    BINAME = 200
    K = BINAME
    L = 2
'縱向排序////////////////////////////////////////////////////////////////////////
    HSB.Cells(L, 1) = K
'********************************BIN 定位******************************************************************************

    Set WW = Worksheets("all_log")
    M = BinSS
    ZZ = BinZZ
    Last = BinLast
    all.Select
    'Call Positioning.Position(WW, QQ, G, SS, ZZ, LAST, FIRST)     'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row

    HSB.Cells(L, 5) = Application.WorksheetFunction.CountIf(all.Range(Cells(M, ZZ), Cells(Last, ZZ)), "=" & HSB.Cells(L, 1))

    Do While K < 299
        K = K + 1
        L = L + 1
        HSB.Cells(L, 1) = K
        HSB.Cells(L, 5) = Application.WorksheetFunction.CountIf(all.Range(Cells(M, ZZ), Cells(Last, ZZ)), "=" & HSB.Cells(L, 1))
    Loop
        
    HSB.Select
    P = 2
    N = 1
    O = 5
    Do Until Cells(P, N) = ""
        If Cells(P, O) = 0 Then
            Cells(P, O).Delete (xlShiftUp)
            Cells(P, N).Delete (xlShiftUp)
        Else
            P = P + 1
        End If
    Loop


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Set H = Worksheets("HW_SW_BIN")
    Set A = Worksheets("all_log")

    Set WW = Worksheets("all_log")
    QQ = "HW_BIN"
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    WW.Range(Cells(first, ZZ), Cells(Last, ZZ)).Copy Destination:=H.Range("H2")
                       
    QQ = " SW_BIN"
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    WW.Range(Cells(first, ZZ), Cells(Last, ZZ)).Copy Destination:=H.Range("I2")
            
    QQ = " BIN"
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    WW.Range(Cells(first, ZZ), Cells(Last, ZZ)).Copy Destination:=H.Range("G2")

Dim Table1
    H.Select
    Last = Cells(65536, 8).End(xlUp).Row
    Table1 = H.Range(Cells(2, 7), Cells(Last, 9))        '利用all log裡面 BIN 、SHWkey 的座標的"位置" 做一個Table
    M = 2
    N = 9
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Do Until Cells(M, 1) = ""
        If H.Cells(M, 1) > 0 Then
            H.Cells(M, 4) = WorksheetFunction.VLookup(H.Cells(M, 1), Table1, 3, False)
            H.Cells(M, 3) = WorksheetFunction.VLookup(H.Cells(M, 1), Table1, 2, False)
            'h.Range(Cells(m, 7)) = WorksheetFunction.VLookup(h.Cells(m, 4), Worksheets("all_log").Range(Cells(6, r), Cells(3005, s)), 5, False)
        Else
            H.Cells(M, 3) = ""
            H.Cells(M, 4) = ""
        End If
        M = M + 1
    Loop

    Worksheets("HW_SW_BIN").Select
    Last = Cells(65536, 5).End(xlUp).Row
    H.Range("E" & Last + 1) = Application.WorksheetFunction.Sum(Range("E2:E" & Last))                       '計算Bin 的總數量
    Set H = Worksheets("HW_SW_BIN")
    O = 2

    Do While Cells(O, 5) > 0
        If H.Cells(O, 5) > 0 Then
            H.Cells(O, 6) = H.Cells(O, 5) / Range("E" & Last + 1)
            O = O + 1
        Else
            Exit Do
        End If
    Loop
    
    Columns("G:I").Clear
    
    H.Range("C1") = "HW"
    H.Range("C" & Last + 1) = "Total"
    H.Range("D1") = "SW"
    H.Range("E1") = "Pcs"
    H.Range("F1") = "%"
    H.Range("F1:F" & Last + 1).Style = "Percent"
    H.Range("F1:F" & Last + 1).NumberFormatLocal = "0.00%"
    H.Range(Cells(Last + 1, 3), Cells(Last + 1, 4)).Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    H.Range(Cells(1, 2), Cells(Last + 1, 2)).Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
'/////////////////////////


  '轉置
 
Dim Arr
    Range("B1:F" & Last + 1).Copy
    Range("H1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    Arr = Range("B1:F" & Last + 1)
    Range("H1").Resize(UBound(Arr, 2), UBound(Arr)) = Application.WorksheetFunction.Transpose(Arr)
    
    Z = 1
    Do While Z < 8
        Worksheets("HW_SW_BIN").Columns(1).Delete
        Z = Z + 1
    Loop

End Sub
Sub CurrentCount()

Dim cs, bin1log
Dim B
Dim RNG
Dim sr
Dim Table1
Dim ICT3VH, ICT3VL, ICT3Vspacing
Dim ICT1VH, ICT1VL, ICT1Vspacing
Dim Caculate

    Set cs = Worksheets("Current_statistics")
    Set bin1log = Worksheets("Bin1_log")
    If (Range("A1") = "有Imaging Current Test" And Range("A2") = "" And Range("A3") = "") Then
        QQ = "Imaging Current Test(3.3V)"
        Set WW = Worksheets("all_log")
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    Else
        Range("B1") = "2"
    End If


End Sub

Sub Imaging_Current_Test(B)

'Current_statistics//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim cs, bin1log
Dim C, D
Dim S, T, U, V, W, X
Dim RNG
Dim sr
Dim Table1
Dim ICT3VH, ICT3VL, ICT3Vspacing
Dim ICT1VH, ICT1VL, ICT1Vspacing
Dim Caculate

    Set cs = Worksheets("Current_statistics")
    Set bin1log = Worksheets("Bin1_log")
    C = B
    Q = 1
    Do Until Q > 2
'********************************BIN 定位******************************************************************************
        QQ = B
        Set WW = Worksheets("all_log")
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            B = "Imaging Current Test(VDD)"
            QQ = B
            D = B
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
'**********************************************************************************************************************
        cs.Select
        cs.Range("C3") = "Max"
        cs.Range("C4") = "Average"
        cs.Range("C5") = "Min"
        bin1log.Select
        Table1 = bin1log.Range(Cells(first, ZZ), Cells(Last, ZZ))
  
 'CS Imaging Current Test(3.3V)
        cs.Select
        S = Application.WorksheetFunction.Max(Table1)
        T = Application.WorksheetFunction.Average(Table1)
        T = WorksheetFunction.Round(T, 2)
        U = Application.WorksheetFunction.Min(Table1)
  
'CS Imaging Current Test(3.3V) 區間
            ICT3VL = U
            ICT3VH = S
            ICT3Vspacing = WorksheetFunction.Round((ICT3VH - ICT3VL) / 6, 3)
            ICT3Vspacing = WorksheetFunction.RoundUp((ICT3VH - ICT3VL) / 6, 3)
            I = 1
            V = ICT3VL & "~" & ICT3VL + ICT3Vspacing
            Caculate = ICT3VL + I * ICT3Vspacing
            If Q = 1 Then
                cs.Range("D2") = B
                cs.Range("D3") = S
                cs.Range("D4") = T
                cs.Range("D5") = U
                cs.Range("H3") = V
                W = 3
                X = 2
            Else
                cs.Range("E2") = B
                cs.Range("E3") = S
                cs.Range("E4") = T
                cs.Range("E5") = U
                cs.Range("H17") = V
                W = 17
                X = 16
            End If
            
            cs.Range("H" & X) = QQ
            cs.Range("I" & X) = "PCS"
            cs.Range("J" & X) = "%"
            
            If Range("D2") > Range("D5") Or Range("E2") > Range("E5") Then
                cs.Cells(W, 9) = Application.WorksheetFunction.CountIf(bin1log.Columns(ZZ), _
                "<=" & Caculate)
                If ICT3Vspacing = 0 Then
                    ICT3Vspacing = 0
                Else
                    ICT3Vspacing = ICT3Vspacing
                End If

                K = ICT3VL + ICT3Vspacing
   
                MM = K
                If X < 6 Then
                    Do Until X > 6
                        X = X + 1
                        K = K + 0.001
                        L = K + ICT3Vspacing
                        cs.Cells(X + 1, 8).Value = K & "~" & L
                        K = L
                    Loop
                Else
                    Do Until X > 20
                        X = X + 1
                        K = K + 0.001
                        L = K + ICT3Vspacing
                        cs.Cells(X + 1, 8).Value = K & "~" & L
                        K = L
                    Loop
                End If
                cs.Range("H" & X + 2) = "Total"
            Else
                cs.Range("H18") = "Total"
            End If
   
'************************************************************************
            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("Current_statistics")
            Call Count.Current區間數量(QQ, WW, WS, II, III, Last, MM, ZZ, ICT3Vspacing)
            Call Drawing.ImangeCurrentMap(cs, B)


            Q = Q + 1
            B = "Imaging Current Test(1.8V)"
        Loop
        
        '/////image Current 結束 2019/01/21

'A08ICT3Vmap()/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
    Call Drawing.Current_form(WS, B)
    Worksheets("all_log").Select
    Worksheets("all_log").Range("I1").Clear
    
End Sub
Sub FOD_Current_Test(B)

'Current_statistics/////////////////////////////////////////////////////////////////////////////////////////////
Dim cs, bin1log
Dim RNG
Dim sr
Dim Table1
Dim ICT3VH, ICT3VL, ICT3Vspacing
Dim ICT1VH, ICT1VL, ICT1Vspacing
Dim Caculate
Dim I, J
Dim C, D, E, F, G, H, K, L
Dim M, N, O, P, Q, R, S, T
Dim U, V
    Set cs = Worksheets("Current_statistics")
    Set bin1log = Worksheets("Bin1_log")
    O = B
    W = 1
        Do Until W > 2
'********************************BIN 定位*******************************
            QQ = B
            Set WW = Worksheets("all_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
            If WW.Range("I1") = "Nothing" Then
                B = " FOD Current Test(VDD)"
                QQ = B
                P = B
                Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            End If
'*****************************************************************
            If cs.Range("A1") = "無Imaging Current Test" Then
                C = 3
            Else
                C = 30
            End If
            D = C + 1
            E = C + 2
            F = C + 3
            H = C + 4
            G = C + 5
            K = C + 6
            L = C + 7
            M = C + 14
            N = C + 15
            O = C + 16
            P = C + 17
            Q = C + 18
            R = C + 19
            S = C + 20
            T = C + 21

  'CS.Range("b1").Value = sr.ColumN
            cs.Range("C" & D) = "Max"
            cs.Range("C" & E) = "Average"
            cs.Range("C" & F) = "Min"
            II = cs.Range("a1")
            
            bin1log.Select
            Table1 = bin1log.Range(Cells(first, ZZ), Cells(Last, ZZ))
            
 'CS Imaging Current Test(3.3V)
            cs.Select
            S = Application.WorksheetFunction.Max(Table1)
            T = Application.WorksheetFunction.Average(Table1)
            T = WorksheetFunction.Round(T, 2)
            U = Application.WorksheetFunction.Min(Table1)

'CS Imaging Current Test(3.3V) 區間
            ICT3VL = U
            ICT3VH = S
            ICT3Vspacing = WorksheetFunction.Round((ICT3VH - ICT3VL) / 6, 3)
            ICT3Vspacing = WorksheetFunction.RoundUp((ICT3VH - ICT3VL) / 6, 3)
            I = 1
            V = ICT3VL & "~" & ICT3VL + ICT3Vspacing
            Caculate = ICT3VL + I * ICT3Vspacing
            If W = 1 Then
                cs.Range("D" & C) = B
                cs.Range("D" & D) = S
                cs.Range("D" & E) = T
                cs.Range("D" & F) = U
                cs.Range("H" & D) = V
                X = D
                D = D
            Else
                cs.Range("E" & C) = B
                cs.Range("E" & D) = S
                cs.Range("E" & E) = T
                cs.Range("E" & F) = U
                cs.Range("H" & M) = V
                D = M
                W = 17
                X = D
            End If
            cs.Range("H" & X - 1) = QQ
            cs.Range("I" & X - 1) = "PCS"
            cs.Range("J" & X - 1) = "%"
            cs.Cells(D, 9) = Application.WorksheetFunction.CountIf(bin1log.Columns(ZZ), _
            "<=" & Caculate)
            If ICT3Vspacing = 0 Then
                ICT3Vspacing = 0
            Else
                ICT3Vspacing = ICT3Vspacing
            End If

            K = ICT3VL + ICT3Vspacing
            MM = K
            If X = 31 Then
                Do Until X > 35
                    X = X + 1
                    K = K + 0.001
                    L = K + ICT3Vspacing
                    cs.Cells(X, 8).Value = K & "~" & L
                    K = L
                Loop
            ElseIf X = 44 Then
                Do Until X > 48
                    K = K + 0.001
                    X = X + 1
                    L = K + ICT3Vspacing
                    cs.Cells(X, 8).Value = K & "~" & L
                    K = L
                Loop
            End If
            cs.Range("H" & X + 1) = "Total"
'************************************************************************
            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("Current_statistics")
            Call Count.Current區間數量(QQ, WW, WS, II, III, Last, MM, ZZ, ICT3Vspacing)
            Call Drawing.FODCurrentMap(cs, B)

'A08ICT3Vmap()/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            W = W + 1
            B = "FOD Current Test (1.8V)"
        Loop
    Call Drawing.Current_form(WS, B)
    Worksheets("all_log").Select
    Worksheets("all_log").Range("I1").Clear
End Sub

Sub PowerDown_Current_Test(B)

'Current_statistics//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim cs, bin1log
Dim RNG
Dim sr
Dim Table1
Dim ICT3VH, ICT3VL, ICT3Vspacing
Dim ICT1VH, ICT1VL, ICT1Vspacing
Dim Caculate
Dim I, J
Dim C, D, E, F, G, H, K, L
Dim M, N, O, P, Q, R, S, T
Dim U, V

    Set cs = Worksheets("Current_statistics")
    Set bin1log = Worksheets("Bin1_log")
    O = B
    W = 1
        Do Until W > 2
'********************************BIN 定位******************************************************************************
            QQ = B
            Set WW = Worksheets("all_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=篩選目標的名稱M,G=位置,SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
            If WW.Range("I1") = "Nothing" Then
                B = " PowerDown Current Test(VDD)"
                QQ = B
                P = B
                Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            End If
'**********************************************************************************************************************

            If cs.Range("A1") = "無Imaging Current Test" And cs.Range("A2") = "無FOD Current Test" Then
                C = 3
            ElseIf cs.Range("A1") = "有Imaging Current Test" And cs.Range("A2") = "無FOD Current Test" Then
                C = 30
            Else
                C = 58
            End If
            D = C + 1
            E = C + 2
            F = C + 3
            H = C + 4
            G = C + 5
            K = C + 6
            L = C + 7
            N = C + 15
            O = C + 16
            P = C + 17
            Q = C + 18
            R = C + 19
            M = C + 14
            S = C + 20
            T = C + 21
  'CS.Range("b1").Value = sr.ColumN
            cs.Range("C" & D) = "Max"
            cs.Range("C" & E) = "Average"
            cs.Range("C" & F) = "Min"
            II = cs.Range("a1")
            III = cs.Range("a2")
            bin1log.Select
            Table1 = bin1log.Range(Cells(first, ZZ), Cells(Last, ZZ))
  
 'CS Imaging Current Test(3.3V)
            cs.Select
            S = Application.WorksheetFunction.Max(Table1)
            T = Application.WorksheetFunction.Average(Table1)
            T = WorksheetFunction.Round(T, 2)
            U = Application.WorksheetFunction.Min(Table1)
'CS Imaging Current Test(3.3V) 區間
            ICT3VL = U
            ICT3VH = S
            ICT3Vspacing = WorksheetFunction.Round((ICT3VH - ICT3VL) / 6, 3)
            ICT3Vspacing = WorksheetFunction.RoundUp((ICT3VH - ICT3VL) / 6, 3)
            I = 1
            V = ICT3VL & "~" & ICT3VL + ICT3Vspacing
            Caculate = ICT3VL + I * ICT3Vspacing

            If W = 1 Then
                cs.Range("D" & C) = B
                cs.Range("D" & D) = S
                cs.Range("D" & E) = T
                cs.Range("D" & F) = U
                cs.Range("H" & D) = V
                X = D
                D = D
            Else
                cs.Range("E" & C) = B
                cs.Range("E" & D) = S
                cs.Range("E" & E) = T
                cs.Range("E" & F) = U
                cs.Range("H" & M) = V
                D = M
                W = 17
                X = D
            End If

            cs.Range("H" & X - 1) = QQ
            cs.Range("I" & X - 1) = "PCS"
            cs.Range("J" & X - 1) = "%"
            cs.Cells(D, 9) = Application.WorksheetFunction.CountIf(bin1log.Columns(ZZ), _
            "<=" & Caculate)
            If ICT3Vspacing = 0 Then
                ICT3Vspacing = 0
            Else
                ICT3Vspacing = ICT3Vspacing
            End If

            K = ICT3VL + ICT3Vspacing
            MM = K
   
            If X < 6 Then
                Do Until X > 6
                    X = X + 1
                    K = K + 0.001
                    L = K + ICT3Vspacing
                    cs.Cells(X + 1, 8).Value = K & "~" & L
                    K = L
                Loop
            ElseIf X = 31 Then
                Do Until X > 35
                    X = X + 1
                    K = K + 0.001
                    L = K + ICT3Vspacing
                    cs.Cells(X, 8).Value = K & "~" & L
                    K = L
                Loop
            ElseIf X = 44 Then
                Do Until X > 48
                    X = X + 1
                    K = K + 0.001
                    L = K + ICT3Vspacing
                    cs.Cells(X, 8).Value = K & "~" & L
                    K = L
                Loop
            ElseIf X = 59 Then
                Do Until X > 63
                    X = X + 1
                    K = K + 0.001
                    L = K + ICT3Vspacing
                    cs.Cells(X, 8).Value = K & "~" & L
                    K = L
                Loop
            ElseIf X = 72 Then
                Do Until X > 76
                    X = X + 1
                    K = K + 0.001
                    L = K + ICT3Vspacing
                    cs.Cells(X, 8).Value = K & "~" & L
                    K = L
                Loop
            End If
            cs.Range("H" & X + 1) = "Total"
'************************************************************************
            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("Current_statistics")
            Call Count.Current區間數量(QQ, WW, WS, II, III, Last, MM, ZZ, ICT3Vspacing)
            Call Drawing.PowerDownCurrentMap(cs, B)
'A08ICT3Vmap()/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            W = W + 1
            B = " PowerDown Current Test (1.8V)"
        Loop
    Call Drawing.Current_form(WS, B)
    Worksheets("all_log").Select
    Worksheets("all_log").Range("I1").Clear
    
End Sub

Sub Finger_count(FN, ByRef MyArray(), I)
Dim L1, L2, L3, L4, L5
Dim R1, R2, R3, R4, R5
    I = 0
    If FRRFORM.CheckBox1.Value = True Then      '使用先前已有群組
        L1 = 1
        MyArray(I) = "L1"
        I = I + 1
        'Me.ComboBox1.Visible = True     '顯示現有群組名稱組合框
        'Me.Label4.Visible = True        '顯示相應說明文字標籤
    Else            '否則
        L1 = 0
        'Me.ComboBox1.Visible = False        '隱藏現有群組名稱組合框
        'Me.Label4.Visible = False       '隱藏相應說明文字標籤

    End If
    
    If FRRFORM.CheckBox2.Value = True Then
        L2 = 1
        MyArray(I) = "L2"
        I = I + 1
    Else
        L2 = 0
    End If
    
    If FRRFORM.CheckBox3.Value = True Then
        L3 = 1
        MyArray(I) = "L3"
        I = I + 1
    Else
        L3 = 0
    End If

    If FRRFORM.CheckBox4.Value = True Then
        L4 = 1
        MyArray(I) = "L4"
        I = I + 1
    Else
        L4 = 0
    End If

    If FRRFORM.CheckBox5.Value = True Then
        L5 = 1
        MyArray(I) = "L5"
        I = I + 1
    Else
        L5 = 0
    End If

    If FRRFORM.CheckBox6.Value = True Then
        R1 = 1
        MyArray(I) = "R1"
        I = I + 1
    Else
        R1 = 0
    End If

    If FRRFORM.CheckBox7.Value = True Then
        R2 = 1
        MyArray(I) = "R2"
        I = I + 1
    Else
        R2 = 0
    End If

    If FRRFORM.CheckBox8.Value = True Then
        R3 = 1
        MyArray(I) = "R3"
        I = I + 1
    Else
        R3 = 0
    End If

    If FRRFORM.CheckBox9.Value = True Then
        R4 = 1
        MyArray(I) = "R4"
        I = I + 1
    Else
        R4 = 0
    End If

    If FRRFORM.CheckBox10.Value = True Then
        R5 = 1
        MyArray(I) = "R5"
        I = I + 1
    Else
        R5 = 0
    End If

    FN = L1 + L2 + L3 + L4 + L5 + R1 + R2 + R3 + R4 + R5
    
End Sub



