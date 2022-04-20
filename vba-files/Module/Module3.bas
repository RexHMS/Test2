Attribute VB_Name = "Module3"
Option Explicit

Sub Mod_three_Click()

Dim MainSheet As String
Dim Filename As String

    UserForm3.Show

End Sub
Sub CountPeople()

    Dim strNowPath As String   '�x�s�ثe�ɮץؿ�
    Dim strFileName As String   'Ū���쪺�ɮצW��
    Dim strFileExt As String    '�ɮװ��ɦW
    Dim strFileNames() As String
    Dim strFFilename() As Variant  '�����ɮצW��
    Dim strFileNameTime, WBN, FT As String  '�ɮװ��ɦW
    Dim N, M, Z, O, FN, I, FC As Integer
    Dim MyArray(10) As Variant
    Dim filetype As Integer
    Dim Sheet1 As Worksheet
    On Error Resume Next
    
    If UserForm3.OptionButton1.Value = True Then '���]��Ʀ�������
        filetype = 1
    Else                                           '���]��Ƨ��S��������
        filetype = 0
    End If
        
    If UserForm3.CheckBox11.Value = True Then
        FC = 1
    ElseIf UserForm3.CheckBox11.Value = True Then
        FC = 0
    End If
    
    If UserForm3.OptionButton15.Value = True Then
        FT = "png"
    ElseIf UserForm3.OptionButton16.Value = True Then
        FT = "bin"
    End If
    '*********�ɮ׸��|*********
    strNowPath = UserForm3.TextBox1.Text '�p�G���]�w�H�]�w���D
 
    
    If Trim(strNowPath) = "" Then        '���]�S�����|
       strNowPath = Excel.ActiveWorkbook.path
    End If
    
    'Unload UserForm3
    N = 0
    
    '�ù���s
    Application.ScreenUpdating = False
    
    If Right(strNowPath, 1) = "\" Then
        strFileName = Dir(strNowPath & strFileExt, vbDirectory)
        strFileNameTime = strNowPath
    Else
        strFileName = Dir(strNowPath & "\" & strFileExt, vbDirectory)
        strFileNameTime = strNowPath & "\"
    End If
         
    While strFileName <> ""
        If strFileName <> ActiveWorkbook.name Then '�o���ɮפ��n���
            If strFileName <> "." And strFileName <> ".." Then
                M = N
                N = N + 1  '�p��H��

                Worksheets("SDK").Cells(N, 18) = strFileName
                strFileNames = Split(strFileName, ".")

                Worksheets("SDK").Cells(N, 19).Value = strFileNames(0)
                           End If
        End If
        strFileName = Dir() 'Ū���U�@���ɮ�
    Wend
    
    ReDim strFFilename(M) As Variant '���s�w�q�B�T�{�H�W�}�C�ƶq
    Z = 0
    O = 1
    Do Until Z > M
        strFFilename(Z) = Worksheets("SDK").Cells(O, 19)
        Z = Z + 1
        O = O + 1
    Loop
    
    Z = strFFilename(M)
    
    Worksheets("SDK").Columns("R:T").Clear  '�N���e�����G�M��
  
    'N = N - 1
    Call FRRFORM(strFFilename(), MyArray(), N, FN, I, WBN) '�ǻ��H�W�}�C�B��������ƶq�}�C�B�H��
    Unload UserForm3
    Call FRRtimes(strFFilename(), MyArray(), N, strNowPath, FN, I, WBN, filetype, FT, FC)
    
    ActiveWorkbook.Save
    Range("A2").Select
    N = Nothing
    FN = Nothing
    Range("A2").Select
End Sub


Sub FRRFORM(ByRef Filename(), ByRef MyArray(), N, FN, I, WBN)

Dim Peo As Integer
Dim EC As Variant
Dim SNa As Variant
Dim MNa As Variant
Dim K, L, M, J, O, P, Q


    EC = UserForm3.TextBox2.Text 'Enroll Count

    Call Module3.Finger_count(FN, MyArray(), I)  '�p��ä�������ƶq

    Workbooks.Add
    
    '*********�e�X���Title*********
    Range("A1") = "Note"
    Range("B1") = "Name"
    Range("C1") = "Finger"
    Range("D1") = "Finger Humidity %"
    
    Range("D1:E1").Select '�X���x�s��
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    Range("F1") = "Enroll count"
    
    If UserForm3.OptionButton1.Value = True Then '���]��Ʀ�������
        Range("G1") = "0'fail count"
        Range("H1") = "45'fail count"
        Range("I1") = "90'fail count"
        Range("J1") = "Avg"
    
        Range("A1:J1").Interior.ColorIndex = 49
        Range("A1:J1").Font.ColorIndex = 2
        Range("A1:J1").HorizontalAlignment = xlCenter
    
        Columns("F").ColumnWidth = 10.63
        Columns("G").ColumnWidth = 9.63
        Columns("H").ColumnWidth = 9.63
        Columns("I").ColumnWidth = 9.63
    Else                                           '���]��Ƨ��S��������
        Range("G1") = "Fail count"
        Range("H1") = "Verify����"
        Range("I1") = "Avg"
        Range("A1:I1").Interior.ColorIndex = 49
        Range("A1:I1").Font.ColorIndex = 2
        Range("A1:I1").HorizontalAlignment = xlCenter
    End If
    '*********�e�X���Title*********
    
    J = 2
    K = 0
    O = 0
    Q = 1
    Do Until Q > N '�H�H�Ƭ��j��A�� 1 �}�l�A��H�� N
        P = J
        'P = Filename(O)
        Range("B" & P).Value = Filename(O) '�x�s�� B-P �L�X�H�W�}�C�� O �Ӧ�m���H
        Range("A" & P).Value = O + 1
    
        Do Until K = I  '�H����������j��A�� 0 �}�l�AI ������Ӽ�
            '*********����������B��׳]�w�BEnroll Count�]�w*********
            Range("c" & J).Value = MyArray(K)
            Range("e" & J).Formula = "=IF($D" & J & ">42%, ""Wet"", IF($D" & J & "<38%,""Dry"",""Normal""))"
            Range("D" & J).Select
            Selection.Style = "Percent"
            Selection.NumberFormatLocal = "0%"
            Range("F" & J).Value = EC
            '*********����������B��׳]�w�BEnroll Count�]�w*********
            
            If UserForm3.OptionButton1.Value = True Then  '���]��Ʀ�������
                Range("j" & J).Formula = "=AVERAGE(G" & J & ":I" & J & ")"

                'Range("G" & J & ":I" & J).Select
                'Selection.Style = "Percent"
                'Selection.NumberFormatLocal = "0%"
    
                Range("J" & J).Select
                Selection.Style = "Percent"
                Selection.NumberFormatLocal = "0.00%"
            Else                                           '���]��Ƨ��S��������
                'Range("F" & J).Value = EC
                Range("I" & J).Select
                Selection.Style = "Percent"
                Selection.NumberFormatLocal = "0%"
                Selection.Formula = "=$G" & J & "/$H" & J
            End If
            J = J + 1
            K = K + 1
        Loop
        
        '*********�椺�A�@�ӤH���Ҧ�����d��X���x�s��*********
        Range("A" & P & ":A" & J - 1).Select '�X���x�s��
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
        Range("B" & P & ":B" & J - 1).Select '�X���x�s��
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
        '*********�椺�A�@�ӤH���Ҧ�����d��X���x�s��*********
        
        Q = Q + 1 '�U�@��
        O = O + 1 '����U�@�쪺�H�W�}�C
        K = 0     '��������j���k�s
    Loop
     
    L = 1
    M = J
    J = J - 1
    
    '*********������*********
    Range("A" & M) = "Avg"
    Range("A" & M & ":F" & M).Select   '�X���x�s��
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    If UserForm3.OptionButton1.Value = True Then  '���]��Ʀ�������
        Range("G" & M).Formula = "=AVERAGE(G2:G" & M - 1 & ")"
        Range("H" & M).Formula = "=AVERAGE(H2:H" & M - 1 & ")"
        Range("I" & M).Formula = "=AVERAGE(I2:I" & M - 1 & ")"
        Range("J" & M).Formula = "=AVERAGE(G" & M & ":I" & M & ")"
        Range("G" & M & ":J" & M).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
        Range("G2:J" & M - 1).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.02999"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = -16383844
        .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    Else                                           '���]��Ƨ��S��������
        Range("G" & M).Formula = "=SUM(G2:G" & M - 1 & ")"
        Range("H" & M).Formula = "=SUM(H2:H" & M - 1 & ")"
        Range("I" & M).Formula = "=AVERAGE(I2:I" & M - 1 & ")"
        Range("I2:I" & M - 1).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
    End If
    '*********������*********
    
    '*********�e�u*********
    If UserForm3.OptionButton1.Value = True Then  '���]��Ʀ�������
        Range("A1:J" & M).Select     '�N�n�ƻs���d�������_��
        Selection.Borders.LineStyle = xlContinuous   '�b������x�s��d��ø�s��u
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
            '.LineStyle = xlContinuous '��u
            '.Weight = xlThick  '�ʽu
            '.Borders(xlEdgeRight).ColorIndex = 3 '����
            'xlContinuous = �ӽu
            'xlThick = �ʽu
        End With
    Else
        Range("A1:I" & M).Select     '�N�n�ƻs���d�������_��
        Selection.Borders.LineStyle = xlContinuous   '�b������x�s��d��ø�s��u
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
            '.LineStyle = xlContinuous '��u
            '.Weight = xlThick  '�ʽu
            '.Borders(xlEdgeRight).ColorIndex = 3 '����
            'xlContinuous = �ӽu
            'xlThick = �ʽu
        End With
    End If
    '*********�e�u*********
    
    Application.Dialogs(xlDialogSaveAs).Show
    
    WBN = ActiveWorkbook.name

    Range("A2").Select
    'Selection.Value = Filename(1)

End Sub


Sub Finger_count(FN, ByRef MyArray(), I)
Dim L1, L2, L3, L4, L5
Dim R1, R2, R3, R4, R5
    I = 0
    If UserForm3.CheckBox1.Value = True Then      '�ϥΥ��e�w���s��
        L1 = 1
        MyArray(I) = "L1"
        I = I + 1
        'Me.ComboBox1.Visible = True     '��ܲ{���s�զW�ٲզX��
        'Me.Label4.Visible = True        '��ܬ���������r����
    Else            '�_�h
        L1 = 0
        'Me.ComboBox1.Visible = False        '���ò{���s�զW�ٲզX��
        'Me.Label4.Visible = False       '���ì���������r����

    End If
    
    If UserForm3.CheckBox2.Value = True Then
        L2 = 1
        MyArray(I) = "L2"
        I = I + 1
    Else
        L2 = 0
    End If
    
    If UserForm3.CheckBox3.Value = True Then
        L3 = 1
        MyArray(I) = "L3"
        I = I + 1
    Else
        L3 = 0
    End If

    If UserForm3.CheckBox4.Value = True Then
        L4 = 1
        MyArray(I) = "L4"
        I = I + 1
    Else
        L4 = 0
    End If

    If UserForm3.CheckBox5.Value = True Then
        L5 = 1
        MyArray(I) = "L5"
        I = I + 1
    Else
        L5 = 0
    End If

    If UserForm3.CheckBox6.Value = True Then
        R1 = 1
        MyArray(I) = "R1"
        I = I + 1
    Else
        R1 = 0
    End If

    If UserForm3.CheckBox7.Value = True Then
        R2 = 1
        MyArray(I) = "R2"
        I = I + 1
    Else
        R2 = 0
    End If

    If UserForm3.CheckBox8.Value = True Then
        R3 = 1
        MyArray(I) = "R3"
        I = I + 1
    Else
        R3 = 0
    End If

    If UserForm3.CheckBox9.Value = True Then
        R4 = 1
        MyArray(I) = "R4"
        I = I + 1
    Else
        R4 = 0
    End If

    If UserForm3.CheckBox10.Value = True Then
        R5 = 1
        MyArray(I) = "R5"
        I = I + 1
    Else
        R5 = 0
    End If

    FN = L1 + L2 + L3 + L4 + L5 + R1 + R2 + R3 + R4 + R5
    
End Sub

Sub FRRtimes(strFFilename(), MyArray(), N, strNowPath, FN, I, WBN, filetype, FT, FC)
Dim DB
Dim myfile, myfrfile, myffrfile, mydrg, path
Dim K, T, L, Z, S, U, V, W, Y, A, B, X, M
Dim fingerpath
Dim fingerArr(1 To 10) As Variant 'Device�h���}�C
Dim fingerName
Dim fingerNameTime
Dim strFileExt As String    '�ɮװ��ɦW
Dim F As Integer


    Workbooks(WBN).Activate
    Z = strNowPath
    strNowPath = Z
    mydrg = "st\"

    L = 0
    K = 1
    S = 1
    B = strNowPath
    W = 2
 
    '�H
    Do Until L > N
        strNowPath = B
        strNowPath = strNowPath & "\" & strFFilename(L)
        A = strNowPath
    
        '���
        Do Until K > I
            F = 1
            If filetype = 1 Then
                strNowPath = A
                If Right(strNowPath, 1) = "\" Then
                    fingerName = Dir(strNowPath & strFileExt, vbDirectory)
                    fingerNameTime = strNowPath
                Else
                    fingerName = Dir(strNowPath & "\" & strFileExt, vbDirectory)
                    fingerNameTime = strNowPath & "\"
                End If
         
                While fingerName <> ""
                    If fingerName <> ActiveWorkbook.name Then '�o���ɮפ��n���
                        If fingerName <> "." And fingerName <> ".." Then
                            fingerArr(F) = fingerName
                            F = F + 1
                        End If
                    End If
                    fingerName = Dir() 'Ū���U�@���ɮ�
                Wend
                
                
                
                strNowPath = strNowPath & "\" & fingerArr(K)
                strNowPath = strNowPath & "\verify\"
                S = 1
        
           
                '����
                Do Until S > 3
                    If S = 1 Then
                        mydrg = "st\"
                        Y = 7
                    ElseIf S = 2 Then
                        mydrg = "45d\"
                        Y = 8
                    Else
                        mydrg = "90d\"
                        Y = 9
                    End If
                    Z = strNowPath
                    strNowPath = strNowPath & mydrg
                    'Z = strNowPath
         
                    myfile = "*." & FT
                    myfrfile = "*_F." & FT
                    myffrfile = "*_F_*." & FT

                    '�έp�`�ƶq
                    strNowPath = strNowPath & myfile
                    'MsgBox strNowPath
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        T = T + 1
                        DB = Dir
                    Loop
                    
                    '�έp _F.png �ƶq
                    strNowPath = Z & mydrg & myfrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        U = U + 1
                        DB = Dir
                    Loop
                    
                    '�έp *_F_*.png �ƶq
                    strNowPath = Z & mydrg & myffrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        X = X + 1
                        DB = Dir
                    Loop
                    
                    '�έp *_fail_*.png �ƶq
                    myffrfile = "*_fail_*." & FT
                    strNowPath = Z & mydrg & myffrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        X = X + 1
                        DB = Dir
                    Loop
                    
                    '�έp *_fail_*.png �ƶq
                    myffrfile = "*_fail." & FT
                    strNowPath = Z & mydrg & myffrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        X = X + 1
                        DB = Dir
                    Loop
                    '��ʤ���
                    V = (U + X) / T
            
                    Workbooks(WBN).Activate
                    Cells(W, Y).Select
                    Cells(W, Y) = V
                    Selection.Style = "Percent"
                    Selection.NumberFormatLocal = "0.00%"
                    
                    If FC = 1 Then
                        
                        path = Z & mydrg
                        
                        Call sort(path)
            
                    End If
                    strNowPath = Z
                    S = S + 1
                    T = 0
                    U = 0
                    X = 0
                    
                Loop
            Else
                strNowPath = A
                If Right(strNowPath, 1) = "\" Then
                    fingerName = Dir(strNowPath & strFileExt, vbDirectory)
                    fingerNameTime = strNowPath
                Else
                    fingerName = Dir(strNowPath & "\" & strFileExt, vbDirectory)
                    fingerNameTime = strNowPath & "\"
                End If
         
                While fingerName <> ""
                    If fingerName <> ActiveWorkbook.name Then '�o���ɮפ��n���
                        If fingerName <> "." And fingerName <> ".." Then
                            fingerArr(F) = fingerName
                            F = F + 1
                        End If
                    End If
                    fingerName = Dir() 'Ū���U�@���ɮ�
                Wend
                
                
                
                strNowPath = strNowPath & "\" & fingerArr(K)
                strNowPath = strNowPath & "\verify\"
                S = 1
                mydrg = "st\"
                Y = 7
                Z = strNowPath
                strNowPath = strNowPath & mydrg
                'Z = strNowPath
         
                myfile = "*.png"
                myfrfile = "*_F_*." & FT
                
                '�έp�`�ƶq
                strNowPath = strNowPath & myfile
                'MsgBox strNowPath
                DB = Dir(strNowPath)
                Do While DB <> ""
                    T = T + 1
                    DB = Dir
                Loop
                
                '�έp *_F_*.png �ƶq
                strNowPath = Z & mydrg & myfrfile
                'MsgBox strNowPath
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
                
                '�έp _F.png �ƶq
                myfrfile = "*_F.*" & FT
                strNowPath = Z & mydrg & myfrfile
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
                
                '�έp *_fail_*.png �ƶq
                myfrfile = "*_fail_*." & FT
                strNowPath = Z & mydrg & myfrfile
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
                
                       '�έp *_fail_*.png �ƶq
                myfrfile = "*_fail.*" & FT
                strNowPath = Z & mydrg & myfrfile
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
            
                Workbooks(WBN).Activate
                Cells(W, Y).Select
                Cells(W, Y) = U
                Cells(W, Y + 1) = T
                
                'Selection.Style = "Percent"
                'Selection.NumberFormatLocal = "0.00%"
                    
                If FC = 1 Then
                        
                    path = Z & mydrg
                        
                    Call sort(path)
            
                End If
                
                strNowPath = Z
                S = S + 1
                T = 0
                U = 0
            End If
            W = W + 1
            K = K + 1
    
        Loop
    
        L = L + 1
        K = 1
    
    Loop

End Sub

Sub sort(path)

'path ���פU�����|
'myffrfile �����ɮצW��
'name ��Ƨ��W��

Dim name As String
Dim xlsApp As New Excel.Application
    xlsApp.Workbooks.Add
    xlsApp.Workbooks(1).SaveAs (path & "list.csv")
    xlsApp.Quit
    Set xlsApp = Nothing
Dim mydir
Dim I
Dim file, first, A
Dim Str
Dim ppath


    '���}list
    Workbooks.Open Filename:=path & "list.csv"
    ActiveWorkbook.Save
    
    '�ù���s
    Application.ScreenUpdating = False
    
    '���|�U�Ҧ��ɮ�
    mydir = Dir(path & "*.*")
    
    Do While mydir <> ""
        Range("A1").Offset(I, 0) = mydir
        mydir = Dir()
        I = I + 1
    Loop
    
    ActiveWorkbook.Save
    'ActiveWorkbook.Close
    '****************************************************************
    
    file = Dir(path & "*", vbDirectory)
    Do While file <> ""
        
        If file <> "." And file <> ".." Then
        
            With ActiveWorkbook  '�Q��FSO�ާ@���}�C����
            
                Sheets(1).Select
                Str = Range("A1").Value

                If Trim(Str) <> "" Then
                    first = file
                    file = Str
                    Do
                        If file Like "*08000000_fail*" Then
                            ppath = path & "partial"
                            If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                MkDir (ppath)                       '�إߥؿ�
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*80000000_fail*" Then
                            ppath = path & "bad"
                            If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                MkDir (ppath)                       '�إߥؿ�
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*88000000_fail*" Then
                            ppath = path & "too_partial"
                            If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                MkDir (ppath)                       '�إߥؿ�
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*80000002_fail*" Then
                            ppath = path & "fast"
                            If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                MkDir (ppath)                       '�إߥؿ�
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*80002000_fail*" Then
                            ppath = path & "water"
                            If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                MkDir (ppath)                       '�إߥؿ�
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*00000000_fail*" Then
                            ppath = path & "original_fail"
                            If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                MkDir (ppath)                       '�إߥؿ�
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                        Else
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                        End If
                         
                    Loop While file <> "" And first <> file
                    Sheets(1).Select
                    Rows(1).Delete Shift:=xlUp
                        
                End If

            End With
        End If
        If file <> "" Then
        file = Dir
        End If
    Loop
    
    ActiveWorkbook.Save
    
    ActiveWorkbook.Close

    Kill path & "list.csv"  '�M���C����
                        
End Sub



