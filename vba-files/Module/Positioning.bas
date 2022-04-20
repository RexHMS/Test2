Attribute VB_Name = "Positioning"
Sub Position(WW, QQ, G, SS, ZZ, Last, first)

    WW.Select
        W = Range("A1").Address

'********************************�M�� UID ��m******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range("A1:ZZ70")
        Set G = .Cells.Find(What:=QQ, after:=WW.Range(W), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
  
        Else
            SS = G.Row             '********************************��� UID ��m******************************************************
            ZZ = G.Column
            'Range("G1") = S
            ' Range("H1") = T
            ' u = S + 3006
            first = G.Row + 4
            Last = Cells(65536, ZZ).End(xlUp).Row
            'A���ƽƻs��B���A�Ƨ�B��
            'Worksheets(5).Range(Cells(S, T), Cells(u, T)).Copy Destination:=Worksheets("all_log").Range("A1")     '�ƻs UID��ƨ� ALL LOG �� A1
            'Worksheets("all_log").Columns(1).Sort key1:=Worksheets("all_log").Range("A1")                         '�Ƨ� ALL LOG �� A
            '�]�wA1���{�b���x�s���m
            'Sheets("all_log").Select
            'Set currentCell = Range("A1")
        End If
    End With

End Sub

Sub PositionSNRbeforeHua(WW, QQ, G, SS, ZZ, Last, first, HuaCol, HuaRow)

    WW.Select
    W = Range("A1").Address
    
'********************************�M�� UID ��m******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range(Cells(1, 1), Cells(HuaRow + 4, HuaCol - 1))
        Set G = .Cells.Find(What:=QQ, after:=WW.Range(W), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
  
        Else
            SS = G.Row             '********************************��� UID ��m******************************************************
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
    
'********************************�M�� UID ��m******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range(Cells(HuaRow, HuaCol), Cells(HuaRow + 4, HuaCol + 10))
        Set G = .Cells.Find(What:=QQ, after:=WW.Range(W), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
  
        Else
            SS = G.Row             '********************************��� UID ��m******************************************************
            ZZ = G.Column
            first = G.Row + 4
            Last = Cells(65536, ZZ).End(xlUp).Row

        End If
    End With
    
End Sub



Sub PositionBin(WW, QQ, G, SS, ZZ, Last, first, II, BINAME)
Dim RNG As Object

    T = SS - 1

'********************************�M�� UID ��m******************************************************************************
    WW.Select
'********************************�M�� UID ��m******************************************************************************
    Set G = WW.Cells.Find(QQ)
    With WW.Range("A1:BV50")
        Set G = .Cells.Find(What:=QQ, after:=WW.Range("A1"), LookIn:=xlFormulas, _
        LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If G Is Nothing Then
            WW.Range("i1").Value = "Nothing"
        Else
            SS = G.Row             '********************************��� UID ��m******************************************************
            ZZ = G.Column
            first = G.End(xlDown).Row
            Last = Cells(65536, ZZ).End(xlUp).Row
            I = first
            
            With WW       '�bOrders�u�@��
                II = I - 1
                BINAME = 201
                
                '*************�Ƨ�******************
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
            
                Set RNG = .Range("A" & II & " :BU" & Last & "")   '�Ҧ���ƽd��
                RNG.AutoFilter Field:=ZZ, Criteria1:=BINAME        '�z��X�Ȭ�201��
                '�]�w�z�ﵲ�G�d��
                Set RNG = RNG.Resize(RNG.Rows.Count - 1).Offset(1, 0).SpecialCells(xlCellTypeVisible)
            End With
        
            '�M���z�ﵲ�G�d��UAREA
            For Each theArea In RNG.Areas
            '�M���UAREA���U�C
            WW.Range("A1 :BU" & Last & "").Copy Destination:=Worksheets("Bin1_log").Range("A1")
            Next
            RNG.AutoFilter
        End If
    End With

End Sub
