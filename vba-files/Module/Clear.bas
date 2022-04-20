Attribute VB_Name = "Clear"
Option Explicit

Sub Delete_UID()

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

