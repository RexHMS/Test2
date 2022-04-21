VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI 
   Caption         =   "Auto_QA report"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   OleObjectBlob   =   "UI.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Click()
    Start = Timer
    Call AutoSorting.AutoSortingReport
    Finish = Timer
    TotalTime = Finish - Start    ' Calculate total time.
    MsgBox "Total for " & TotalTime & " seconds"
    ActiveWorkbook.Save
End Sub

Private Sub CommandButton2_Click()

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

Private Sub CommandButton3_Click()
    FRRFORM.Show
    Unload FRRFORM
End Sub

Private Sub CommandButton4_Click()
    
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

Private Sub CommandButton5_Click()
    
Dim I As Variant
Dim SN As String
Dim K As Integer
Dim QQ As String
Dim HWBINRow, HWBINCol As Integer
Dim SWBINRow, SWBINCol As Integer
Dim TSRow, TSCol As Integer
Dim UIDRow, UIDCOL As Integer

    PictureForm.Show
    If wtf <> 1 Then
        End
    End If
    
    BookName = ActiveWorkbook.name

    Workbooks(BookName).Activate
    Worksheets(2).Activate
    
    SN = ActiveSheet.name
    I = ActiveWorkbook.path
    
    Application.ScreenUpdating = False
    
    QQ = "HW_BIN"
    Set WW = Worksheets(SN)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    HWBINRow = SS
    HWBINCol = ZZ
    WW.Activate
    
    QQ = " SW_BIN"
    Set WW = Worksheets(SN)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    SWBINRow = SS
    SWBINCol = ZZ
    WW.Activate
    
    QQ = "Test Sequence"
    Set WW = Worksheets(SN)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    TSRow = SS
    TSCol = ZZ
    WW.Activate

    QQ = "UID"
    Set WW = Worksheets(SN)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
    UIDRow = SS
    UIDCOL = ZZ
    WW.Activate
    
    Dim RNG As Range        '自動篩選結果範圍
    Dim theRow As Range     '各區域的資料列
    Dim theArea As Range        '各區域範圍

    Workbooks(BookName).Activate
    Worksheets(SN).Select

    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    For J = 1 To LASTB
    Worksheets(SN).Select
        If J > 1 Then
            Sheets.Add before:=ActiveSheet  '增加Sheet
            Worksheets(J).name = "BIN " & B(J)
        Else
            Sheets(1).name = "BIN " & B(J)
        End If
        Worksheets(SN).Select
        With Worksheets(SN)
            Rows(BINRow + 3 & ":" & BINRow + 3).Select
            Selection.AutoFilter
            ActiveWindow.SmallScroll ToRight:=28
            ActiveSheet.Range("$A$22:$BW" & Last).AutoFilter Field:=BINCol, Criteria1:=B(J)
            LASTA = Cells(65536, 1).End(xlUp).Row
            
            Range(Cells(BINRow - 1, TSCol), Cells(LASTA, TSCol)).Copy _
                Destination:=Worksheets(J).Range("A1")
            Range(Cells(BINRow - 1, UIDCOL), Cells(LASTA, UIDCOL)).Copy _
                Destination:=Worksheets(J).Range("B1")
            Range(Cells(BINRow - 1, BINCol), Cells(LASTA, BINCol)).Copy _
                Destination:=Worksheets(J).Range("C1")
            Range(Cells(BINRow - 1, HWBINCol), Cells(LASTA, HWBINCol)).Copy _
                Destination:=Worksheets(J).Range("D1")
            Range(Cells(BINRow - 1, SWBINCol), Cells(LASTA, SWBINCol)).Copy _
                Destination:=Worksheets(J).Range("E1")
            ActiveSheet.UsedRange.AutoFilter
            Range("A1").Activate
        End With
            Worksheets(J).Select
         Call PICKIMG(I)
    Next J
    
    ActiveWorkbook.Save
End Sub

Private Sub CommandButton6_Click()
    
Dim Last As Long
Dim TextBox1
Dim QQ, WW, G, SS, ZZ, first
Dim UIDCOL
Dim RightEND

    TextBox1 = Application.GetOpenFilename("(*.csv),*.csv", , Title:="請瀏覽並選取欲載入的檔案")
    If TextBox1 = False Then
        End
    End If
    Workbooks.Open Filename:=TextBox1
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.path & "\" & "UID_checked_" & Left(ActiveWorkbook.name, Len(ActiveWorkbook.name) - 3) & "csv", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False
        
    'Application.ScreenUpdating = False
    
'********************************UID 定位******************************************************************************
    QQ = "UID"
    Set WW = Worksheets(1)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first) 'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row
        If Range("i1") = "Nothing" Then
            QQ = " Sensor UID"
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
    UIDCOL = ZZ
    
    '從第10萬行往回找最後一筆資料
    Last = Cells(100000, 5).End(xlUp).Row

    
    '刪掉 UID為 0x000000000000 的所有資
    With ActiveSheet.Range(Cells(first - 1, ZZ), Cells(Last, ZZ))
        .AutoFilter Field:=1, Criteria1:="0x000000000000"
        If Cells(first - 1, ZZ).End(xlDown).Row <> Rows.Count Then
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        .AutoFilter
    End With

    '刪掉 UID為 0xFFFFFFFFFFFF 的所有資料
        With ActiveSheet.Range(Cells(first - 1, ZZ), Cells(Last, ZZ))
        .AutoFilter Field:=1, Criteria1:="0xFFFFFFFFFFFF"
        If Cells(first - 1, ZZ).End(xlDown).Row <> Rows.Count Then
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        .AutoFilter
    End With
    
'********************************Test Sequence 定位******************************************************************************
    QQ = "Test Sequence"
    Set WW = Worksheets(1)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first) 'QQ=篩選目標的名稱M,G=位置,  SS = G.Row,ZZ = G.Column,LAST=最後一筆資料的Row


    '反向排序，測試順序最後面的，排到第一筆
    first = first - 1
    Rows(first).Select
    Selection.AutoFilter
    ActiveSheet.AutoFilter.sort.SortFields.Clear
    ActiveSheet.AutoFilter.sort.SortFields.Add Key:= _
        Cells(first, ZZ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        '.AutoFilter
    End With
    '.AutoFilter
    
    RightEND = Range("ZZ31").End(xlToLeft).Column
    
    '刪掉重複資料
    first = first + 1
    Range(Cells(first, 1), Cells(Last, RightEND)).Select
    
    ActiveSheet.Range(Cells(first, 1), Cells(Last, RightEND)).RemoveDuplicates Columns:=Array(UIDCOL), Header:=xlNo
    
    ActiveSheet.AutoFilter.sort.SortFields.Clear
    ActiveSheet.AutoFilter.sort.SortFields.Add Key:= _
        Cells(first, ZZ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveWorkbook.Save

    Cells(first, UIDCOL).Select
End Sub


Private Sub CommandButton7_Click()
Dim MainSheet As String
Dim Filename As String

    UserForm1.Show
End Sub

Private Sub CommandButton8_Click()
Dim MainSheet As String
Dim Filename As String

    UserForm2.Show

End Sub

Private Sub CommandButton9_Click()
Dim MainSheet As String
Dim Filename As String

    UserForm3.Show


End Sub
