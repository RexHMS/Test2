Attribute VB_Name = "PickupPic"
Public MyIMGArray(20) As String
Public IMGF As Integer
Public LASTB As Integer
Public J As Integer
Public BINRow As Integer
Public BINCol As Integer
Public RightEND As Integer
Public B(100) As Integer ': Dim J As Integer
Public wtf As Integer
      

Sub PickupPic()

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
    
    Dim RNG As Range        '�۰ʿz�ﵲ�G�d��
    Dim theRow As Range     '�U�ϰ쪺��ƦC
    Dim theArea As Range        '�U�ϰ�d��

    Workbooks(BookName).Activate
    Worksheets(SN).Select

    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    For J = 1 To LASTB
    Worksheets(SN).Select
        If J > 1 Then
            Sheets.Add before:=ActiveSheet  '�W�[Sheet
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

Sub PICKIMG(I)

Dim A, B As String
Dim F As Integer

Dim IMGN As Integer
    F = 6
    IMGN = 5 + IMGF
    L = 1
    For IMGN = 6 To IMGN
        Cells(2, IMGN) = MyIMGArray(L)
        L = L + 1
    Next IMGN
    
    Range(Columns(6), Columns(IMGN - 1)).Select
    Selection.ColumnWidth = 17.25
    K = 6
    
    Do While Cells(K, 1) <> 0
        Rows(K).Select
        Selection.RowHeight = 54.75
        Cells(K, 1).NumberFormat = "@"
        Cells(K, 1).Value = Format(Cells(K, 1), "000000")
        A = Cells(K, 1)
        B = Cells(K, 4).Value
        L = 1
        For IMGN = 6 To IMGN - 1
            Set R = Cells(K, IMGN)
            If Dir(I & "\image\BIN" & B & "\" & A & "_" & MyIMGArray(L) & ".bmp") <> "" Then
                ActiveSheet.Shapes.AddPicture _
                    (I & "\image\BIN" & B & "\" & A & "_" & MyIMGArray(L) & ".bmp", True, True, R.Left, R.Top, -1, -1).Select
                Selection.ShapeRange.Height = 51.874015748
                Selection.ShapeRange.IncrementTop 0.75
                Selection.ShapeRange.IncrementLeft 0.75
                'With Selection.ShapeRange.Line
                '    .Visible = msoTrue
                '    .ForeColor.RGB = RGB(0, 0, 0)
                '    .Transparency = 0
                'End With
            Else
                Cells(K, IMGN).Value = "N/A"
            End If
            Selection.Interior.Color = RGB(221, 235, 247)
            L = L + 1
        Next IMGN
        Cells(K, 1).NumberFormat = "@"
        Cells(K, 1).Value = Format(Cells(K, 1), "")
        

        K = K + 1
    Loop
    
    Columns("A:E").EntireColumn.AutoFit
    Rows(F).Select
    ActiveWindow.FreezePanes = True
    Range("A1").Activate

End Sub

Sub IMGCount(IMGF)

Dim I1, I2, I3, I4, I5, I6, I7, I8, I9, I10 As Integer
Dim I11, I12, I13, I14, I15, I16, I17, I18, I19, I20 As Integer

    I = 1
    If PictureForm.CheckBox1.Value = True Then      '�ϥΥ��e�w���s��
        I1 = 1
        MyIMGArray(I) = "0"
        I = I + 1
        'Me.ComboBox1.Visible = True     '��ܲ{���s�զW�ٲզX��
        'Me.Label4.Visible = True        '��ܬ���������r����
    Else            '�_�h
        I1 = 0
        'Me.ComboBox1.Visible = False        '���ò{���s�զW�ٲզX��
        'Me.Label4.Visible = False       '���ì���������r����

    End If
    
    If PictureForm.CheckBox2.Value = True Then
        I2 = 1
        MyIMGArray(I) = "1"
        I = I + 1
    Else
        I2 = 0
    End If
    
    If PictureForm.CheckBox3.Value = True Then
        I3 = 1
        MyIMGArray(I) = "2"
        I = I + 1
    Else
        I3 = 0
    End If

    If PictureForm.CheckBox4.Value = True Then
        I4 = 1
        MyIMGArray(I) = "3"
        I = I + 1
    Else
        I4 = 0
    End If

    If PictureForm.CheckBox5.Value = True Then
        I5 = 1
        MyIMGArray(I) = "4"
        I = I + 1
    Else
        I5 = 0
    End If

    If PictureForm.CheckBox6.Value = True Then
        I6 = 1
        MyIMGArray(I) = "cds"
        I = I + 1
    Else
        I6 = 0
    End If

    If PictureForm.CheckBox7.Value = True Then
        I7 = 1
        MyIMGArray(I) = "checkbk"
        I = I + 1
    Else
        I7 = 0
    End If

    If PictureForm.CheckBox8.Value = True Then
        I8 = 1
        MyIMGArray(I) = "fod_bg"
        I = I + 1
    Else
        I8 = 0
    End If

    If PictureForm.CheckBox9.Value = True Then
        I9 = 1
        MyIMGArray(I) = "fod_on"
        I = I + 1
    Else
        I9 = 0
    End If
    
    If PictureForm.CheckBox10.Value = True Then
        I10 = 1
        MyIMGArray(I) = "ori_bk"
        I = I + 1
    Else
        I10 = 0
    End If
    
    If PictureForm.CheckBox11.Value = True Then
        I11 = 1
        MyIMGArray(I) = "raw"
        I = I + 1
    Else
        I11 = 0
    End If

    IMGF = I1 + I2 + I3 + I4 + I5 + I6 + I7 + I8 + I9 + I10 + I11
    


End Sub

Sub OptionBin(LASTB)
    
    SN = ActiveSheet.name
    I = ActiveWorkbook.path
    Sheets.Add before:=ActiveSheet  '�W�[Sheet
    
    If PictureForm.OptionButton1.Value = True Then
        QQ = " BIN"
        Set WW = Worksheets(SN)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        BINRow = SS
        BINCol = ZZ
        WW.Activate
        
        RightEND = Cells(BINRow, 100).End(xlToLeft).Column  '��X��Ƴ̥k�䪺�d��
    
        Range(Cells(BINRow, BINCol), Cells(Last, BINCol)).Copy _
            Destination:=Worksheets(1).Range("A1")  '�ƻs���檺Bin ��Sheet 1 A1

        Worksheets(1).Activate
        Columns(1).sort key1:=Range("A1")
        '�]�wA1���{�b���x�s���m
        Set currentCell = Range("A1")
    
        '�ϥ�do..loop�j���˴��{�b���x�s���m�O�_���ŭ�
        '�ŭȴN����A���O�ŭȴN��U���x�s����
        '�Y�O�ۦP�ȡA�h�R���{�b�����s��
        '�̫�A�N�{�b�x�s��]�w���U���x�s��A�H�K�~����
        Do While Not IsEmpty(currentCell)
            Set nextCell = currentCell.Offset(1, 0)
            If nextCell.Value = currentCell.Value Then
                currentCell.Delete xlShiftUp
            End If
            Set currentCell = nextCell
        Loop

        LASTA = Cells(65536, 1).End(xlUp).Row
        LASTB = LASTA - 1                           '��X�Ҧ������ƪ�BIN���ƶq
        
        For J = 1 To LASTB
            B(J) = Cells(J, 1)
        Next J                                      '�ΰ}�C b �O���Ҧ���BIN
    
        Columns("A").Clear
    
    ElseIf PictureForm.OptionButton2.Value = True Then
        QQ = " BIN"
        Set WW = Worksheets(SN)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        BINRow = SS
        BINCol = ZZ
        WW.Activate
        Worksheets(SN).Select
        With Worksheets(SN)
            Rows(BINRow + 3 & ":" & BINRow + 3).Select
            Selection.AutoFilter
            ActiveWindow.SmallScroll ToRight:=28
            ActiveSheet.Range("$A$22:$BW" & Last).AutoFilter Field:=BINCol, Criteria1:=201
            LASTA = Cells(65536, 1).End(xlUp).Row
            B(1) = 201
            LASTB = 1
        End With

    ElseIf PictureForm.OptionButton3.Value = True Then
        QQ = " BIN"
        Set WW = Worksheets(SN)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        BINRow = SS
        BINCol = ZZ
        WW.Activate
        Worksheets(SN).Select
        With Worksheets(SN)
            Rows(BINRow + 3 & ":" & BINRow + 3).Select
            Selection.AutoFilter
            ActiveWindow.SmallScroll ToRight:=28
            ActiveSheet.Range("$A$22:$BW" & Last).AutoFilter Field:=BINCol, Criteria1:="<>201"
            LASTA = Cells(65536, 1).End(xlUp).Row
            RightEND = Cells(BINRow, 100).End(xlToLeft).Column  '��X��Ƴ̥k�䪺�d��
            Range(Cells(BINRow, BINCol), Cells(Last, BINCol)).Copy _
            Destination:=Worksheets(1).Range("A1")  '�ƻs���檺Bin ��Sheet 1 A1
            Worksheets(1).Activate
            Columns(1).sort key1:=Range("A1")
            '�]�wA1���{�b���x�s���m
            Set currentCell = Range("A1")
    
            '�ϥ�do..loop�j���˴��{�b���x�s���m�O�_���ŭ�
            '�ŭȴN����A���O�ŭȴN��U���x�s����
            '�Y�O�ۦP�ȡA�h�R���{�b�����s��
            '�̫�A�N�{�b�x�s��]�w���U���x�s��A�H�K�~����
            Do While Not IsEmpty(currentCell)
                Set nextCell = currentCell.Offset(1, 0)
                If nextCell.Value = currentCell.Value Then
                    currentCell.Delete xlShiftUp
                End If
                Set currentCell = nextCell
            Loop
            LASTA = Cells(65536, 1).End(xlUp).Row
            LASTB = LASTA - 1                           '��X�Ҧ������ƪ�BIN���ƶq
            For J = 1 To LASTB
                B(J) = Cells(J, 1)
            Next J                                      '�ΰ}�C b �O���Ҧ���BIN
            Columns("A").Clear
        End With


    ElseIf PictureForm.OptionButton4.Value = True Then
        FB = PictureForm.TextBox2.Value
        QQ = " BIN"
        Set WW = Worksheets(SN)
        Call Positioning.Position(WW, QQ, G, SS, ZZ, Last, first)
        BINRow = SS
        BINCol = ZZ
        WW.Activate
        
        Worksheets(SN).Select
        With Worksheets(SN)
            Rows(BINRow + 3 & ":" & BINRow + 3).Select
            Selection.AutoFilter
            ActiveWindow.SmallScroll ToRight:=28
            ActiveSheet.Range("$A$22:$BW" & Last).AutoFilter Field:=BINCol, Criteria1:=FB
            LASTA = Cells(65536, 1).End(xlUp).Row
            B(1) = FB
            LASTB = 1
        End With
    End If
End Sub

