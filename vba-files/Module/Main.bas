Attribute VB_Name = "Main"

Sub AutoSortingReport_Click()
    Start = Timer
    Call AutoSorting.AutoSortingReport
    Finish = Timer
    TotalTime = Finish - Start    ' Calculate total time.
    MsgBox "Total for " & TotalTime & " seconds"
    ActiveWorkbook.Save
End Sub


Sub MP_Click()
Dim QQ
Dim A
Dim B
Dim C
Dim SingalRange
Dim SRMax
Dim SRMin
Dim SRAvg

    Start = Timer
    Call AutoSorting.AutoSortingReport
    Sheets(7).Copy before:=Sheets(6)
    Sheets(6).name = "FRR_X-section_Sample"
    
    QQ = "Signal(RV)"
    Set WW = Worksheets("FRR_X-section_Sample")
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)    'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If WW.Range("I1") = "Nothing" Then
            WW.Range("I1").Clear
            QQ = "Ridge-Valley Value"
            Set WW = Worksheets("FRR_X-section_Sample")
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            
            If WW.Range("I1") = "Nothing" Then
                WW.Range("I1").Clear
                QQ = "SignalOut"
                Set WW = Worksheets("FRR_X-section_Sample")
                Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
            End If
        End If
    SingalRange = Worksheets("FRR_X-section_Sample").Range(Cells(first, ZZ), Cells(Last, ZZ))
    
    Rows("5:5").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields.Add Key:= _
        Range("H5"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "BIN1 Sequence"
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("I6:I7").Select
    Selection.AutoFill Destination:=Range("I" & first & ":I" & Last)
    Application.CutCopyMode = False
    
    I = first
    II = I - 1
    ZZ = ZZ + 1
    Rows(II & ":" & II).Select
    'Selection.AutoFilt
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Add Key:=Cells(II, ZZ), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '****************************************************
    '上色排序篩選RV 高中低各兩顆
    Rows(I + 1 & ":" & I).Select      '低
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    
    
    SRMax = Application.WorksheetFunction.Max(SingalRange)
    SRAvg = Application.WorksheetFunction.Average(SingalRange)
    SRAvg = WorksheetFunction.Round(SRAvg, 2)
    SRMin = Application.WorksheetFunction.Min(SingalRange)
    
    QQ = "Huawei SNR test"
    Set WW = Worksheets("FRR_X-section_Sample")
        WW.Range("I1").Clear
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    
    
     If WW.Range("I1") = "Nothing" Then
    
    C = Worksheets("RV_Noise").Range("C4")
    E = WorksheetFunction.Round(C, 0)
    F = E
    Else
    C = SRAvg
    E = WorksheetFunction.Round(C, 0)
    F = E
    ZZ = ZZ + 1
    End If
    'ZZ = ZZ + 1
    Set WW = Worksheets("FRR_X-section_Sample")
    WW.Range(Cells(I, ZZ), Cells(Last, ZZ)).Select
        Set sr = ActiveSheet.Cells.Find(F)
        With Selection
            Set sr = .Cells.Find(What:=F, after:=Cells(I, ZZ), LookIn:=xlValues, _
            LookAt:=1, SearchOrder:=1, SearchDirection:=xlNext, _
            MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        End With
        
        If sr Is Nothing Then
            F = E + 1
            G = 0
            GoTo Line1
        Else
            SRR = sr.Row
            SRC = sr.Column
            GoTo Line2
        End If
Line1:
    Set WW = Worksheets("FRR_X-section_Sample")
    WW.Range(Cells(I, ZZ), Cells(Last, ZZ)).Select
        Set sr = ActiveSheet.Cells.Find(F)
        With Selection
            Set sr = .Cells.Find(What:=F, after:=Cells(I, ZZ), LookIn:=xlValues, _
            LookAt:=1, SearchOrder:=1, SearchDirection:=xlNext, _
            MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        End With
        
        If sr Is Nothing Then
            G = G + 1
            F = F - G - 1
            GoTo Line3
        Else
            SRR = sr.Row
            SRC = sr.Column
            GoTo Line2
        End If
        
Line3:
    Set WW = Worksheets("FRR_X-section_Sample")
    WW.Range(Cells(I, ZZ), Cells(Last, ZZ)).Select
        Set sr = ActiveSheet.Cells.Find(F)
        With Selection
            Set sr = .Cells.Find(What:=F, after:=Cells(I, ZZ), LookIn:=xlValues, _
            LookAt:=1, SearchOrder:=1, SearchDirection:=xlNext, _
            MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        End With
        
        If sr Is Nothing Then
            G = G + 1
            F = F + G + 1
            GoTo Line1
        Else
            SRR = sr.Row
            SRC = sr.Column
            GoTo Line2
        End If
        
        

Line2:
    Rows(SRR & ":" & SRR + 1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With


    Rows(Last - 1 & ":" & Last).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-171
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Add(Range("m5"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue. _
        Color = RGB(226, 239, 218)
    With ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("m5").Select
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Add(Range("m5"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue. _
        Color = RGB(255, 242, 204)
    With ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort.SortFields. _
        Add(Range("m5"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue. _
        Color = RGB(252, 228, 214)
    With ActiveWorkbook.Worksheets("FRR_X-section_Sample").AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    
    
    Finish = Timer
    TotalTime = Finish - Start    ' Calculate total time.
    MsgBox "Total for " & TotalTime & " seconds"
    ActiveWorkbook.Save
End Sub

Sub FRR_Form_Click()

    FRRFORM.Show
    Unload FRRFORM

End Sub





