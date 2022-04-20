Attribute VB_Name = "Clear"
Option Explicit

Sub Delete_UID()

Dim Last As Long
Dim TextBox1
Dim QQ, WW, G, SS, ZZ, first
Dim UIDCOL
Dim RightEND

    TextBox1 = Application.GetOpenFilename("(*.csv),*.csv", , Title:="���s���ÿ�������J���ɮ�")
    If TextBox1 = False Then
        End
    End If
    Workbooks.Open Filename:=TextBox1
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.path & "\" & "UID_checked_" & Left(ActiveWorkbook.name, Len(ActiveWorkbook.name) - 3) & "csv", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False
        
    'Application.ScreenUpdating = False
    
'********************************UID �w��******************************************************************************
    QQ = "UID"
    Set WW = Worksheets(1)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first) 'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row
        If Range("i1") = "Nothing" Then
            QQ = " Sensor UID"
            Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        End If
    UIDCOL = ZZ
    
    '�q��10�U�橹�^��̫�@�����
    Last = Cells(100000, 5).End(xlUp).Row

    
    '�R�� UID�� 0x000000000000 ���Ҧ���
    With ActiveSheet.Range(Cells(first - 1, ZZ), Cells(Last, ZZ))
        .AutoFilter Field:=1, Criteria1:="0x000000000000"
        If Cells(first - 1, ZZ).End(xlDown).Row <> Rows.Count Then
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        .AutoFilter
    End With

    '�R�� UID�� 0xFFFFFFFFFFFF ���Ҧ����
        With ActiveSheet.Range(Cells(first - 1, ZZ), Cells(Last, ZZ))
        .AutoFilter Field:=1, Criteria1:="0xFFFFFFFFFFFF"
        If Cells(first - 1, ZZ).End(xlDown).Row <> Rows.Count Then
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        .AutoFilter
    End With
    
'********************************Test Sequence �w��******************************************************************************
    QQ = "Test Sequence"
    Set WW = Worksheets(1)
    Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first) 'QQ=�z��ؼЪ��W��M,G=��m,  SS = G.Row,ZZ = G.Column,LAST=�̫�@����ƪ�Row


    '�ϦV�ƧǡA���ն��ǳ̫᭱���A�ƨ�Ĥ@��
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
    
    '�R�����Ƹ��
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

