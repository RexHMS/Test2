Attribute VB_Name = "AutoSorting"
Sub AutoSortingReport()

Dim UID                '*********UID �w��*********
Dim G As Range         '*********UID �w��*********
Dim S, T, U            '*********UID �w��*********
Dim V                  '*********UID �w��*********
Dim W
Dim Hua, HuaCol, HuaRow
Dim HuaSNR
Dim B As String
Dim sr As Range
Dim RNG As Range        '�۰ʿz�ﵲ�G�d��
Dim theRow As Range     '�U�ϰ쪺��ƦC
Dim theArea As Range        '�U�ϰ�d��
Dim I, J, K, L
Dim RVL, RVH, NOL, NOH
Dim RVspacing
Dim NOspacing
Dim RN As Worksheet
Dim bin1log As Worksheet
Dim myChart As ChartObject
Dim HSB As Worksheet
Dim all As Worksheet
Dim QQ
Dim HH
Dim WW
Dim BINAME
Dim BinSS, BinZZ, BinLast


    
    TextBox1 = Application.GetOpenFilename("(*.csv),*.csv", , Title:="���s���ÿ�������J���ɮ�")
    If TextBox1 = False Then
        End
    End If
    Workbooks.Open Filename:=TextBox1
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.path & "\" & Left(ActiveWorkbook.name, Len(ActiveWorkbook.name) - 3) & "xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False
        
    Application.ScreenUpdating = False
        
    Sheets.Add before:=ActiveSheet '�b�̫e���s�W�@��Sheet
    Sheets.Add before:=ActiveSheet '�b�̫e���s�W�@��Sheet
    Sheets.Add before:=ActiveSheet
    Sheets.Add before:=ActiveSheet

    Sheets(1).name = "HW_SW_BIN"   'Sheets(1) �W�r�s HW_SW BIN
    Sheets(2).name = "RV_Noise"    'Sheets(2) �W�r�s RV_Noise
    Sheets(3).name = "all_log"     'Sheets(3) �W�r�s all log
    Sheets(4).name = "Bin1_log"    'Sheets(4) �W�r�s Bin1 log


'********************************Test Sequence �w��******************************************************************************
    QQ = "Test Sequence"
    Set WW = Worksheets(Worksheets.Count)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)   'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
    HH = Last
    II = first
'�ˬdUID ����//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'********************************UID �w��******************************************************************************
    Worksheets(5).Activate
    QQ = "UID"
    Set WW = Worksheets(5)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first) 'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
        If Range("i1") = "Nothing" Then
            QQ = " Sensor UID"
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
    T = SS - 1
    V = ZZ + 2
    W = Cells(SS, V)
'********************************�ˬdUID����******************************************************************************
    Call Check.CheckUID(WW, SS, ZZ, HH, Last)
    Worksheets("all_log").Columns(1).Clear
    Worksheets("all_log").Range("E1").Clear
    Worksheets("all_log").Range("F1").Clear
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Worksheets(5).Activate
    Worksheets(5).Range(Cells(T, 1), Cells(HH, 74)).Copy Destination:=Worksheets("all_log").Range("A1")
'********************************BIN �w��B�ƻs******************************************************************************

    QQ = " BIN"
    Set WW = Worksheets("all_log")
    Call Positioning.PositionBin(WW, QQ, G, SS, ZZ, Last, first, II, BINAME) 'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
    BinSS = SS
    BinZZ = ZZ
    BinLast = Last
'********************************�M�� UID ��m******************************************************************************

'********************************RV �w��******************************************************************************
    Worksheets("RV_Noise").Activate
    Worksheets("RV_Noise").Range("B3") = "Max"
    Worksheets("RV_Noise").Range("B4") = "Average"
    Worksheets("RV_Noise").Range("B5") = "Min"
    
    QQ = "Huawei SNR test"
    Set WW = Worksheets("Bin1_log")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)       '�ˬd�O�_���ج� case�AQQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
        If WW.Range("I1") = "Nothing" Then                          '�p�G���O�AI1 = Nothing
            WW.Range("I1").Clear                                    '�M�� I1
            QQ = "Signal(RV)"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
                If WW.Range("I1") = "Nothing" Then
                    WW.Range("I1").Clear
                    QQ = "Ridge-Valley Value"
                    Set WW = Worksheets("Bin1_log")
                    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
                End If
 
            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("RV_Noise")
            Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

            Set RN = Worksheets("RV_Noise")
            Set bin1log = Worksheets("Bin1_log")
            Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("RV_Noise")
            Call Count.�϶��ƶq(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
            Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
        End If

'********************************Noise �w��******************************************************************************
    QQ = "Huawei SNR test"
    Set WW = Worksheets("Bin1_log")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            QQ = "Noise"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
  
            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("RV_Noise")
            Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

            Set RN = Worksheets("RV_Noise")
            Set bin1log = Worksheets("Bin1_log")
            Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("RV_Noise")
            Call Count.�϶��ƶq(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
            Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
        End If

'********************************SNR �w��******************************************************************************

    QQ = "Huawei SNR test"
    Set WW = Worksheets("Bin1_log")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     ''�ˬd�O�_���ج� case�AQQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
    HuaCol = ZZ
    HuaRow = SS
    If QQ = "Huawei SNR test" And WW.Range("i1").Value = "Nothing" Then
        Hua = 0
    Else
        Hua = 1
    End If
    
    '���s�b�ج�SNR�A�p�ⴶ�qSNR�A��ø��
    If Hua = 0 Then
        HuaSNR = 0
        QQ = "SNR(RV)"
        Set WW = Worksheets("Bin1_log")
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
        If WW.Range("i1").Value = "Nothing" Then
            QQ = "SNR"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�
        End If
        
        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

        Set RN = Worksheets("RV_Noise")
        Set bin1log = Worksheets("Bin1_log")
        Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.�϶��ƶq(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
        Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
    
    '�s�b�ج�SNR�A�p�ⴶ�qSNR�A��ø��
    'ElseIf Hua = 1 Then
        'HuaSNR = 0
        'QQ = "SNR(RV)"
        'Set WW = Worksheets("Bin1_log")
        'Call Positioning.PositionSNRbeforeHua(WW, QQ, G, SS, ZZ, LAST, FIRST, HuaCol, HuaRow)   'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
        'If WW.Range("i1").Value = "Nothing" Then
        '    QQ = "SNR"
        '    Set WW = Worksheets("Bin1_log")
        '    Call Positioning.PositionSNRbeforeHua(WW, QQ, G, SS, ZZ, LAST, FIRST, HuaCol, HuaRow)    'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�
        'End If
        
        'Set WW = Worksheets("Bin1_log")
        'Set WS = Worksheets("RV_Noise")
        'Call Count.RN_MAX(QQ, WW, WS, II, LAST, ZZ, W, Hua, HuaSNR)

        'Set RN = Worksheets("RV_Noise")
        'Set bin1log = Worksheets("Bin1_log")
        'Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

        'Set WW = Worksheets("Bin1_log")
        'Set WS = Worksheets("RV_Noise")
        'Call Count.�϶��ƶq(QQ, WW, WS, II, LAST, MM, ZZ, RVspacing, W, Hua, HuaSNR)
        'Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
    End If


'********************************Huawei SNR �w��******************************************************************************

    '�s�b�ج�SNR�A�p��ج�SNR�A��ø��
    If Hua = 1 Then
        HuaSNR = 1
        QQ = "SNR"
        Set WW = Worksheets("Bin1_log")
        Call Positioning.PositionSNRafterHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)
        If WW.Range("i1").Value = "Nothing" Then
            QQ = "SNR(RV)"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.PositionSNRafterHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)    'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�
        End If
        
        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

        Set RN = Worksheets("RV_Noise")
        Set bin1log = Worksheets("Bin1_log")
        Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.�϶��ƶq(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
        Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
    
    Else
        WW.Range("I1").Clear
            
    End If

    Worksheets("Bin1_log").Range("I1").Clear

'*************************��q�y*********************************************************************************************
    
    Sheets.Add after:=Worksheets("RV_Noise")
    Sheets(3).name = "Current_statistics"    '�W�r�s Current_statistics"

    Set cs = Worksheets("Current_statistics")
    Set bin1log = Worksheets("Bin1_log")
        With Worksheets("Current_statistics")
            Call Check.CheckCurrent(cs, bin1log)
        End With


'**********************************************************************************************************************
    Call Count.HWSW(II, BinSS, BinZZ, BinLast, BINAME)
    Call Drawing.IC_Form
    Sheets("HW_SW_BIN").Activate
    Range("A1").Formula = "=IC_information!R[1]C[1]"
    Call Drawing.Report_form


'*************************�x�s����ɶ�*********************************************************************************************
    ActiveWorkbook.Save

End Sub


