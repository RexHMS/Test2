Attribute VB_Name = "Module3"
Option Explicit

Sub Mod_three_Click()

Dim MainSheet As String
Dim Filename As String

    UserForm3.Show

End Sub
Sub CountPeople()

    Dim strNowPath As String   '儲存目前檔案目錄
    Dim strFileName As String   '讀取到的檔案名稱
    Dim strFileExt As String    '檔案副檔名
    Dim strFileNames() As String
    Dim strFFilename() As Variant  '紀錄檔案名稱
    Dim strFileNameTime, WBN, FT As String  '檔案副檔名
    Dim N, M, Z, O, FN, I, FC As Integer
    Dim MyArray(10) As Variant
    Dim filetype As Integer
    Dim Sheet1 As Worksheet
    On Error Resume Next
    
    If UserForm3.OptionButton1.Value = True Then '假設資料有分角度
        filetype = 1
    Else                                           '假設資料夾沒有分角度
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
    '*********檔案路徑*********
    strNowPath = UserForm3.TextBox1.Text '如果有設定以設定為主
 
    
    If Trim(strNowPath) = "" Then        '假設沒有路徑
       strNowPath = Excel.ActiveWorkbook.path
    End If
    
    'Unload UserForm3
    N = 0
    
    '螢幕更新
    Application.ScreenUpdating = False
    
    If Right(strNowPath, 1) = "\" Then
        strFileName = Dir(strNowPath & strFileExt, vbDirectory)
        strFileNameTime = strNowPath
    Else
        strFileName = Dir(strNowPath & "\" & strFileExt, vbDirectory)
        strFileNameTime = strNowPath & "\"
    End If
         
    While strFileName <> ""
        If strFileName <> ActiveWorkbook.name Then '這個檔案不要顯示
            If strFileName <> "." And strFileName <> ".." Then
                M = N
                N = N + 1  '計算人數

                Worksheets("SDK").Cells(N, 18) = strFileName
                strFileNames = Split(strFileName, ".")

                Worksheets("SDK").Cells(N, 19).Value = strFileNames(0)
                           End If
        End If
        strFileName = Dir() '讀取下一個檔案
    Wend
    
    ReDim strFFilename(M) As Variant '重新定義且確認人名陣列數量
    Z = 0
    O = 1
    Do Until Z > M
        strFFilename(Z) = Worksheets("SDK").Cells(O, 19)
        Z = Z + 1
        O = O + 1
    Loop
    
    Z = strFFilename(M)
    
    Worksheets("SDK").Columns("R:T").Clear  '將之前的結果清除
  
    'N = N - 1
    Call FRRFORM(strFFilename(), MyArray(), N, FN, I, WBN) '傳遞人名陣列、手指種類數量陣列、人數
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

    Call Module3.Finger_count(FN, MyArray(), I)  '計算並分類手指數量

    Workbooks.Add
    
    '*********畫出表格Title*********
    Range("A1") = "Note"
    Range("B1") = "Name"
    Range("C1") = "Finger"
    Range("D1") = "Finger Humidity %"
    
    Range("D1:E1").Select '合併儲存格
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    Range("F1") = "Enroll count"
    
    If UserForm3.OptionButton1.Value = True Then '假設資料有分角度
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
    Else                                           '假設資料夾沒有分角度
        Range("G1") = "Fail count"
        Range("H1") = "Verify次數"
        Range("I1") = "Avg"
        Range("A1:I1").Interior.ColorIndex = 49
        Range("A1:I1").Font.ColorIndex = 2
        Range("A1:I1").HorizontalAlignment = xlCenter
    End If
    '*********畫出表格Title*********
    
    J = 2
    K = 0
    O = 0
    Q = 1
    Do Until Q > N '以人數為迴圈，由 1 開始，到人數 N
        P = J
        'P = Filename(O)
        Range("B" & P).Value = Filename(O) '儲存格 B-P 印出人名陣列第 O 個位置的人
        Range("A" & P).Value = O + 1
    
        Do Until K = I  '以手指種類為迴圈，由 0 開始，I 為手指個數
            '*********表格手指種類、濕度設定、Enroll Count設定*********
            Range("c" & J).Value = MyArray(K)
            Range("e" & J).Formula = "=IF($D" & J & ">42%, ""Wet"", IF($D" & J & "<38%,""Dry"",""Normal""))"
            Range("D" & J).Select
            Selection.Style = "Percent"
            Selection.NumberFormatLocal = "0%"
            Range("F" & J).Value = EC
            '*********表格手指種類、濕度設定、Enroll Count設定*********
            
            If UserForm3.OptionButton1.Value = True Then  '假設資料有分角度
                Range("j" & J).Formula = "=AVERAGE(G" & J & ":I" & J & ")"

                'Range("G" & J & ":I" & J).Select
                'Selection.Style = "Percent"
                'Selection.NumberFormatLocal = "0%"
    
                Range("J" & J).Select
                Selection.Style = "Percent"
                Selection.NumberFormatLocal = "0.00%"
            Else                                           '假設資料夾沒有分角度
                'Range("F" & J).Value = EC
                Range("I" & J).Select
                Selection.Style = "Percent"
                Selection.NumberFormatLocal = "0%"
                Selection.Formula = "=$G" & J & "/$H" & J
            End If
            J = J + 1
            K = K + 1
        Loop
        
        '*********格內，一個人的所有手指範圍合併儲存格*********
        Range("A" & P & ":A" & J - 1).Select '合併儲存格
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
        Range("B" & P & ":B" & J - 1).Select '合併儲存格
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
        '*********格內，一個人的所有手指範圍合併儲存格*********
        
        Q = Q + 1 '下一位
        O = O + 1 '接續下一位的人名陣列
        K = 0     '手指種類迴圈歸零
    Loop
     
    L = 1
    M = J
    J = J - 1
    
    '*********表格尾部*********
    Range("A" & M) = "Avg"
    Range("A" & M & ":F" & M).Select   '合併儲存格
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    
    If UserForm3.OptionButton1.Value = True Then  '假設資料有分角度
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
    Else                                           '假設資料夾沒有分角度
        Range("G" & M).Formula = "=SUM(G2:G" & M - 1 & ")"
        Range("H" & M).Formula = "=SUM(H2:H" & M - 1 & ")"
        Range("I" & M).Formula = "=AVERAGE(I2:I" & M - 1 & ")"
        Range("I2:I" & M - 1).Select
        Selection.Style = "Percent"
        Selection.NumberFormatLocal = "0.00%"
    End If
    '*********表格尾部*********
    
    '*********畫線*********
    If UserForm3.OptionButton1.Value = True Then  '假設資料有分角度
        Range("A1:J" & M).Select     '將要複製的範圍先選取起來
        Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
            '.LineStyle = xlContinuous '實線
            '.Weight = xlThick  '粗線
            '.Borders(xlEdgeRight).ColorIndex = 3 '紅色
            'xlContinuous = 細線
            'xlThick = 粗線
        End With
    Else
        Range("A1:I" & M).Select     '將要複製的範圍先選取起來
        Selection.Borders.LineStyle = xlContinuous   '在選取的儲存格範圍中繪製格線
        With Selection
            .Borders(xlEdgeTop).Weight = xlThick
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).Weight = xlThick
            '.LineStyle = xlContinuous '實線
            '.Weight = xlThick  '粗線
            '.Borders(xlEdgeRight).ColorIndex = 3 '紅色
            'xlContinuous = 細線
            'xlThick = 粗線
        End With
    End If
    '*********畫線*********
    
    Application.Dialogs(xlDialogSaveAs).Show
    
    WBN = ActiveWorkbook.name

    Range("A2").Select
    'Selection.Value = Filename(1)

End Sub


Sub Finger_count(FN, ByRef MyArray(), I)
Dim L1, L2, L3, L4, L5
Dim R1, R2, R3, R4, R5
    I = 0
    If UserForm3.CheckBox1.Value = True Then      '使用先前已有群組
        L1 = 1
        MyArray(I) = "L1"
        I = I + 1
        'Me.ComboBox1.Visible = True     '顯示現有群組名稱組合框
        'Me.Label4.Visible = True        '顯示相應說明文字標籤
    Else            '否則
        L1 = 0
        'Me.ComboBox1.Visible = False        '隱藏現有群組名稱組合框
        'Me.Label4.Visible = False       '隱藏相應說明文字標籤

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
Dim fingerArr(1 To 10) As Variant 'Device多維陣列
Dim fingerName
Dim fingerNameTime
Dim strFileExt As String    '檔案副檔名
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
 
    '人
    Do Until L > N
        strNowPath = B
        strNowPath = strNowPath & "\" & strFFilename(L)
        A = strNowPath
    
        '手指
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
                    If fingerName <> ActiveWorkbook.name Then '這個檔案不要顯示
                        If fingerName <> "." And fingerName <> ".." Then
                            fingerArr(F) = fingerName
                            F = F + 1
                        End If
                    End If
                    fingerName = Dir() '讀取下一個檔案
                Wend
                
                
                
                strNowPath = strNowPath & "\" & fingerArr(K)
                strNowPath = strNowPath & "\verify\"
                S = 1
        
           
                '角度
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

                    '統計總數量
                    strNowPath = strNowPath & myfile
                    'MsgBox strNowPath
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        T = T + 1
                        DB = Dir
                    Loop
                    
                    '統計 _F.png 數量
                    strNowPath = Z & mydrg & myfrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        U = U + 1
                        DB = Dir
                    Loop
                    
                    '統計 *_F_*.png 數量
                    strNowPath = Z & mydrg & myffrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        X = X + 1
                        DB = Dir
                    Loop
                    
                    '統計 *_fail_*.png 數量
                    myffrfile = "*_fail_*." & FT
                    strNowPath = Z & mydrg & myffrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        X = X + 1
                        DB = Dir
                    Loop
                    
                    '統計 *_fail_*.png 數量
                    myffrfile = "*_fail." & FT
                    strNowPath = Z & mydrg & myffrfile
                    DB = Dir(strNowPath)
                    Do While DB <> ""
                        X = X + 1
                        DB = Dir
                    Loop
                    '算百分比
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
                    If fingerName <> ActiveWorkbook.name Then '這個檔案不要顯示
                        If fingerName <> "." And fingerName <> ".." Then
                            fingerArr(F) = fingerName
                            F = F + 1
                        End If
                    End If
                    fingerName = Dir() '讀取下一個檔案
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
                
                '統計總數量
                strNowPath = strNowPath & myfile
                'MsgBox strNowPath
                DB = Dir(strNowPath)
                Do While DB <> ""
                    T = T + 1
                    DB = Dir
                Loop
                
                '統計 *_F_*.png 數量
                strNowPath = Z & mydrg & myfrfile
                'MsgBox strNowPath
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
                
                '統計 _F.png 數量
                myfrfile = "*_F.*" & FT
                strNowPath = Z & mydrg & myfrfile
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
                
                '統計 *_fail_*.png 數量
                myfrfile = "*_fail_*." & FT
                strNowPath = Z & mydrg & myfrfile
                DB = Dir(strNowPath)
                Do While DB <> ""
                    U = U + 1
                    DB = Dir
                Loop
                
                       '統計 *_fail_*.png 數量
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

'path 角度下的路徑
'myffrfile 關鍵檔案名稱
'name 資料夾名稱

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


    '打開list
    Workbooks.Open Filename:=path & "list.csv"
    ActiveWorkbook.Save
    
    '螢幕更新
    Application.ScreenUpdating = False
    
    '路徑下所有檔案
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
        
            With ActiveWorkbook  '利用FSO操作打開列表文件
            
                Sheets(1).Select
                Str = Range("A1").Value

                If Trim(Str) <> "" Then
                    first = file
                    file = Str
                    Do
                        If file Like "*08000000_fail*" Then
                            ppath = path & "partial"
                            If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                MkDir (ppath)                       '建立目錄
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*80000000_fail*" Then
                            ppath = path & "bad"
                            If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                MkDir (ppath)                       '建立目錄
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*88000000_fail*" Then
                            ppath = path & "too_partial"
                            If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                MkDir (ppath)                       '建立目錄
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*80000002_fail*" Then
                            ppath = path & "fast"
                            If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                MkDir (ppath)                       '建立目錄
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*80002000_fail*" Then
                            ppath = path & "water"
                            If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                MkDir (ppath)                       '建立目錄
                            End If
                                
                            Name path & file As ppath & "\" & file
                                    
                            Sheets(1).Select
                            Rows(1).Delete Shift:=xlUp
                            Str = Range("A1").Value
                            file = Str
                                    
                        ElseIf file Like "*00000000_fail*" Then
                            ppath = path & "original_fail"
                            If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                MkDir (ppath)                       '建立目錄
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

    Kill path & "list.csv"  '清除列表文件
                        
End Sub



