Attribute VB_Name = "AutoSorting"
Sub AutoSortingReport()

Dim UID                '*********UID 定位*********
Dim G As Range         '*********UID 定位*********
Dim S, T, U            '*********UID 定位*********
Dim V                  '*********UID 定位*********
Dim W
Dim Hua, HuaCol, HuaRow
Dim HuaSNR
Dim B As String
Dim sr As Range
Dim RNG As Range        '自動篩選結果範圍
Dim theRow As Range     '各區域的資料列
Dim theArea As Range        '各區域範圍
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


    
    TextBox1 = Application.GetOpenFilename("(*.csv),*.csv", , Title:="請瀏覽並選取欲載入的檔案")
    If TextBox1 = False Then
        End
    End If
    Workbooks.Open Filename:=TextBox1
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.path & "\" & Left(ActiveWorkbook.name, Len(ActiveWorkbook.name) - 3) & "xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False
        
    Application.ScreenUpdating = False
        
    Sheets.Add before:=ActiveSheet '在最前面新增一個Sheet
    Sheets.Add before:=ActiveSheet '在最前面新增一個Sheet
    Sheets.Add before:=ActiveSheet
    Sheets.Add before:=ActiveSheet

    Sheets(1).name = "HW_SW_BIN"   'Sheets(1) 名字叫 HW_SW BIN
    Sheets(2).name = "RV_Noise"    'Sheets(2) 名字叫 RV_Noise
    Sheets(3).name = "all_log"     'Sheets(3) 名字叫 all log
    Sheets(4).name = "Bin1_log"    'Sheets(4) 名字叫 Bin1 log


'********************************Test Sequence 定位******************************************************************************
    QQ = "Test Sequence"
    Set WW = Worksheets(Worksheets.Count)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)   'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
    HH = Last
    II = first
'檢查UID 重複//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'********************************UID 定位******************************************************************************
    Worksheets(5).Activate
    QQ = "UID"
    Set WW = Worksheets(5)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first) 'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If Range("i1") = "Nothing" Then
            QQ = " Sensor UID"
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
    T = SS - 1
    V = ZZ + 2
    W = Cells(SS, V)
'********************************檢查UID重複******************************************************************************
    Call Check.CheckUID(WW, SS, ZZ, HH, Last)
    Worksheets("all_log").Columns(1).Clear
    Worksheets("all_log").Range("E1").Clear
    Worksheets("all_log").Range("F1").Clear
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Worksheets(5).Activate
    Worksheets(5).Range(Cells(T, 1), Cells(HH, 74)).Copy Destination:=Worksheets("all_log").Range("A1")
'********************************BIN 定位、複製******************************************************************************

    QQ = " BIN"
    Set WW = Worksheets("all_log")
    Call Positioning.PositionBin(WW, QQ, G, SS, ZZ, Last, first, II, BINAME) 'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
    BinSS = SS
    BinZZ = ZZ
    BinLast = Last
'********************************尋找 UID 位置******************************************************************************

'********************************RV 定位******************************************************************************
    Worksheets("RV_Noise").Activate
    Worksheets("RV_Noise").Range("B3") = "Max"
    Worksheets("RV_Noise").Range("B4") = "Average"
    Worksheets("RV_Noise").Range("B5") = "Min"
    
    QQ = "Huawei SNR test"
    Set WW = Worksheets("Bin1_log")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)       '檢查是否為華為 case，QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If WW.Range("I1") = "Nothing" Then                          '如果不是，I1 = Nothing
            WW.Range("I1").Clear                                    '清空 I1
            QQ = "Signal(RV)"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
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
            Call Count.區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
            Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
        End If

'********************************Noise 定位******************************************************************************
    QQ = "Huawei SNR test"
    Set WW = Worksheets("Bin1_log")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            QQ = "Noise"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
  
            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("RV_Noise")
            Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

            Set RN = Worksheets("RV_Noise")
            Set bin1log = Worksheets("Bin1_log")
            Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

            Set WW = Worksheets("Bin1_log")
            Set WS = Worksheets("RV_Noise")
            Call Count.區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
            Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
        End If

'********************************SNR 定位******************************************************************************

    QQ = "Huawei SNR test"
    Set WW = Worksheets("Bin1_log")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     ''檢查是否為華為 case，QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
    HuaCol = ZZ
    HuaRow = SS
    If QQ = "Huawei SNR test" And WW.Range("i1").Value = "Nothing" Then
        Hua = 0
    Else
        Hua = 1
    End If
    
    '不存在華為SNR，計算普通SNR，並繪圖
    If Hua = 0 Then
        HuaSNR = 0
        QQ = "SNR(RV)"
        Set WW = Worksheets("Bin1_log")
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If WW.Range("i1").Value = "Nothing" Then
            QQ = "SNR"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)     'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的
        End If
        
        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

        Set RN = Worksheets("RV_Noise")
        Set bin1log = Worksheets("Bin1_log")
        Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
        Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
    
    '存在華為SNR，計算普通SNR，並繪圖
    'ElseIf Hua = 1 Then
        'HuaSNR = 0
        'QQ = "SNR(RV)"
        'Set WW = Worksheets("Bin1_log")
        'Call Positioning.PositionSNRbeforeHua(WW, QQ, G, SS, ZZ, LAST, FIRST, HuaCol, HuaRow)   'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        'If WW.Range("i1").Value = "Nothing" Then
        '    QQ = "SNR"
        '    Set WW = Worksheets("Bin1_log")
        '    Call Positioning.PositionSNRbeforeHua(WW, QQ, G, SS, ZZ, LAST, FIRST, HuaCol, HuaRow)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的
        'End If
        
        'Set WW = Worksheets("Bin1_log")
        'Set WS = Worksheets("RV_Noise")
        'Call Count.RN_MAX(QQ, WW, WS, II, LAST, ZZ, W, Hua, HuaSNR)

        'Set RN = Worksheets("RV_Noise")
        'Set bin1log = Worksheets("Bin1_log")
        'Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

        'Set WW = Worksheets("Bin1_log")
        'Set WS = Worksheets("RV_Noise")
        'Call Count.區間數量(QQ, WW, WS, II, LAST, MM, ZZ, RVspacing, W, Hua, HuaSNR)
        'Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
    End If


'********************************Huawei SNR 定位******************************************************************************

    '存在華為SNR，計算華為SNR，並繪圖
    If Hua = 1 Then
        HuaSNR = 1
        QQ = "SNR"
        Set WW = Worksheets("Bin1_log")
        Call Positioning.PositionSNRafterHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)
        If WW.Range("i1").Value = "Nothing" Then
            QQ = "SNR(RV)"
            Set WW = Worksheets("Bin1_log")
            Call Positioning.PositionSNRafterHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的
        End If
        
        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.RN_MAX(QQ, WW, WS, II, Last, ZZ, W, Hua, HuaSNR)

        Set RN = Worksheets("RV_Noise")
        Set bin1log = Worksheets("Bin1_log")
        Call Count.RENG(RN, bin1log, RVspacing, MM, QQ, W, Hua, HuaSNR)

        Set WW = Worksheets("Bin1_log")
        Set WS = Worksheets("RV_Noise")
        Call Count.區間數量(QQ, WW, WS, II, Last, MM, ZZ, RVspacing, W, Hua, HuaSNR)
        Call Drawing.RNS_Distribution(QQ, RN, Hua, HuaSNR)
    
    Else
        WW.Range("I1").Clear
            
    End If

    Worksheets("Bin1_log").Range("I1").Clear

'*************************算電流*********************************************************************************************
    
    Sheets.Add after:=Worksheets("RV_Noise")
    Sheets(3).name = "Current_statistics"    '名字叫 Current_statistics"

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


'*************************儲存結算時間*********************************************************************************************
    ActiveWorkbook.Save

End Sub


