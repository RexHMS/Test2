Attribute VB_Name = "Check"
Sub CheckUID(WW, SS, ZZ, HH, Last)
'********************************尋找 UID 位置******************************************************************************

 'A欄資料複製到B欄後，排序B欄
    WW.Range(Cells(SS, ZZ), Cells(HH, ZZ)).Copy Destination:=Worksheets("all_log").Range("A2")    '複製 UID整排到 ALL LOG 的 A1
    Worksheets("all_log").Columns(1).sort key1:=Worksheets("all_log").Range("A3")   '排序 ALL LOG 的 A
    Sheets("all_log").Activate
    Range("A1").Insert Shift:=xlDown
    '設定A1為現在的儲存格位置
   
    Set currentCell = Range("A1")
    With ActiveSheet.Range("A1 :A" & Last + 10 & "")
        .AutoFilter Field:=1, Criteria1:="0x000000000000"
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        .AutoFilter
    End With

    With ActiveSheet.Range("A1 :A" & Last + 10 & "")
        .AutoFilter Field:=1, Criteria1:="0xFFFFFFFFFFFF"
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        .AutoFilter
    End With

    Range("e1") = Cells(65536, 1).End(xlUp).Row
    ActiveSheet.Range("A1 :A" & Last + 10 & "").RemoveDuplicates Columns:=1, Header:=xlNo
    Range("f1") = Cells(65536, 1).End(xlUp).Row
    If Range("e1") = Range("F1") Then
        MsgBox "UID沒有重複!!"
    Else
        MsgBox "UID有重複!!"
    End If

End Sub

Sub CheckCurrent(cs, bin1log)
Dim B As String

    B = "Imaging Current Test(3.3V)"
    Set sr = ActiveSheet.Cells.Find(B)
    With Worksheets("all_log").Range("A2:BU2")
        Set sr = .Cells.Find(What:=B, after:=bin1log.Range("A2"), LookIn:=xlFormulas, _
            LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
            MatchCase:=False, MatchByte:=False, SearchFormat:=True)
    End With
    If sr Is Nothing Then
        B = "Imaging Current Test(VCC)"
        Set sr = ActiveSheet.Cells.Find(B)
        With Worksheets("all_log").Range("A2:BU2")
            Set sr = .Cells.Find(What:=B, after:=bin1log.Range("A2"), LookIn:=xlFormulas, _
                LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
                MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        End With
    End If
    If sr Is Nothing Then
        cs.Range("A1").Value = "無Imaging Current Test"
    Else
        cs.Range("A1").Value = "有Imaging Current Test"
        Call Count.Imaging_Current_Test(B)
    End If

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
    B = "FOD Current Test(3.3V)"
    Set sr = ActiveSheet.Cells.Find(B)
    With Worksheets("all_log").Range("A2:BU2")
        Set sr = .Cells.Find(What:=B, after:=bin1log.Range("A2"), LookIn:=xlFormulas, _
            LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
            MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If sr Is Nothing Then
            B = "FOD Current Test(VCC)"
            Set sr = ActiveSheet.Cells.Find(B)
            With Worksheets("all_log").Range("A2:BU2")
                Set sr = .Cells.Find(What:=B, after:=bin1log.Range("A2"), LookIn:=xlFormulas, _
                    LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
                    MatchCase:=False, MatchByte:=False, SearchFormat:=True)
            End With
        End If
        If sr Is Nothing Then
            cs.Range("A2").Value = "無FOD Current Test"
        Else
            cs.Range("A2").Value = "有FOD Current Test"
            Call Count.FOD_Current_Test(B)
        End If
    End With
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
    B = "PowerDown Current Test(3.3V)"
    Set sr = ActiveSheet.Cells.Find(B)
    With Worksheets("all_log").Range("A2:BU2")
        Set sr = .Cells.Find(What:=B, after:=bin1log.Range("A2"), LookIn:=xlFormulas, _
            LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
            MatchCase:=False, MatchByte:=False, SearchFormat:=True)
        If sr Is Nothing Then
            B = "PowerDown Current Test(VCC)"
            Set sr = ActiveSheet.Cells.Find(B)
            With Worksheets("all_log").Range("A2:BU2")
                Set sr = .Cells.Find(What:=B, after:=bin1log.Range("A2"), LookIn:=xlFormulas, _
                    LookAt:=1, SearchOrder:=2, SearchDirection:=xlNext, _
                    MatchCase:=False, MatchByte:=False, SearchFormat:=True)
            End With
        End If
        If sr Is Nothing Then
            cs.Range("A3").Value = "無PowerDown Current Test"
        Else
            cs.Range("A3").Value = "有PowerDown Current Test"
            Call Count.PowerDown_Current_Test(B)
        End If
    End With

End Sub

