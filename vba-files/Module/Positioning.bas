Attribute VB_Name = "Positioning"
Sub Position(WW, QQ, G, SS, ZZ, Last, first)

    WW.Select
        W = Range("A1").Address

'********************************尋找 UID 位置******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range("A1:ZZ70")
        Set G = .Cells.Find(What:=QQ, after:=WW.Range(W), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
  
        Else
            SS = G.Row             '********************************找到 UID 位置******************************************************
            ZZ = G.Column
            'Range("G1") = S
            ' Range("H1") = T
            ' u = S + 3006
            first = G.Row + 4
            Last = Cells(65536, ZZ).End(xlUp).Row
            'A欄資料複製到B欄後，排序B欄
            'Worksheets(5).Range(Cells(S, T), Cells(u, T)).Copy Destination:=Worksheets("all_log").Range("A1")     '複製 UID整排到 ALL LOG 的 A1
            'Worksheets("all_log").Columns(1).Sort key1:=Worksheets("all_log").Range("A1")                         '排序 ALL LOG 的 A
            '設定A1為現在的儲存格位置
            'Sheets("all_log").Select
            'Set currentCell = Range("A1")
        End If
    End With

End Sub

Sub PositionSNRbeforeHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)

    WW.Select
    W = Range("A1").Address
    
'********************************尋找 UID 位置******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range(Cells(1, 1), Cells(HuaRow + 4, HuaCol - 1))
        Set G = .Cells.Find(What:=QQ, after:=WW.Range(W), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
  
        Else
            SS = G.Row             '********************************找到 UID 位置******************************************************
            ZZ = G.Column
            first = G.Row + 4
            Last = Cells(65536, ZZ).End(xlUp).Row
            Range("i1").Clear
        End If
    End With
    
End Sub
Sub PositionSNRafterHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)

    WW.Select
    W = Cells(HuaRow, HuaCol).Address
    
'********************************尋找 UID 位置******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range(Cells(HuaRow, HuaCol), Cells(HuaRow + 4, HuaCol + 10))
        Set G = .Cells.Find(What:=QQ, after:=WW.Range(W), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
  
        Else
            SS = G.Row             '********************************找到 UID 位置******************************************************
            ZZ = G.Column
            first = G.Row + 4
            Last = Cells(65536, ZZ).End(xlUp).Row

        End If
    End With
    
End Sub



Sub PositionBin(WW, QQ, G, SS, ZZ, Last, first, II, BINAME)
Dim RNG As Object

    T = SS - 1

'********************************尋找 UID 位置******************************************************************************
    WW.Select
'********************************尋找 UID 位置******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range("A1:BV50")
        Set G = .Cells.Find(What:=QQ, after:=WW.Range("A1"), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
        Else
            SS = G.Row             '********************************找到 UID 位置******************************************************
            ZZ = G.Column
            first = G.End(xlDown).Row
            Last = Cells(65536, ZZ).End(xlUp).Row
            I = first
            
            With WW       '在Orders工作表中
                II = I - 1
                BINAME = 201
                
                '*************排序******************
                Rows(II & ":" & II).Select
                Selection.AutoFilter
                ActiveSheet.AutoFilter.sort. _
                SortFields.Add Key:=Cells(II, ZZ), SortOn:=xlSortOnValues, Order:= _
                xlAscending, DataOption:=xlSortNormal
            
                With ActiveSheet.AutoFilter. _
                    sort
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                '***********************************
                
                
                Range(Cells(II, ZZ), Cells(Last, ZZ)).Select
                T = 0
                For Each RNG In Selection
                    If RNG.Text Like "*" & BINAME & "*" Then
                        BINAME = 201
                        T = T + 1
                    End If
                Next
            
                If T = 0 Then
                    ChangeBin.Show
                    Unload ChangeBin
                    BINAME = Range("A1")
                    Range("A1").Clear
                End If
            
                Set RNG = .Range("A" & II & " :BU" & Last & "")   '所有資料範圍
                RNG.AutoFilter Field:=ZZ, Criteria1:=BINAME        '篩選出值為201者
                '設定篩選結果範圍
                Set RNG = RNG.Resize(RNG.Rows.Count - 1).Offset(1, 0).SpecialCells(xlCellTypeVisible)
            End With
        
            '遍歷篩選結果範圍各AREA
            For Each theArea In RNG.Areas
            '遍歷各AREA的各列
            WW.Range("A1 :BU" & Last & "").Copy Destination:=Worksheets("Bin1_log").Range("A1")
            Next
            RNG.AutoFilter
        End If
    End With

End Sub
