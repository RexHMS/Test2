VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRRForm 
   Caption         =   "FRR 表格生成器"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "FRRForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "FRRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Dim Peo As Integer
Dim EC As Variant
Dim SNa As Variant
Dim MNa As Variant
Dim FN, K, L, M
Dim MyArray(10) As Variant

    Peo = TextBox1.Text
    EC = TextBox2.Text
    SNa = TextBox3.Text
    MNa = TextBox4.Text
    Call Count.Finger_count(FN, MyArray(), I)
        
    'MsgBox Peo & "人," & FN & "手指," & i & "," & MyArray(0)
    
    Unload Me

    Workbooks.Add
    
    Sheets(1).name = MNa & "_RV" & SNa
    
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
    
    J = 2
    K = 0
    Do Until K = I
    Range("c" & J).Value = MyArray(K)
    Range("e" & J).Formula = "=IF($D" & J & ">42%, ""Wet"", IF($D" & J & "<38%,""Dry"",""Normal""))"
    Range("j" & J).Formula = "=AVERAGE(G" & J & ":I" & J & ")"
    Range("D" & J).Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0%"
    
    Range("F" & J).Value = EC
    
    Range("G" & J & ":I" & J).Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0%"
    
    Range("J" & J).Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.00%"
    J = J + 1
    K = K + 1
    Loop
    
    Range("A2:A" & J - 1).Select '合併儲存格
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
    Range("B2:B" & J - 1).Select '合併儲存格
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
        
    L = 1
    M = J
    J = J - 1
    Do Until L = Peo
    Range("A2:J" & J).Copy _
    Destination:=Range("A" & M)
    L = L + 1
    M = M + FN
    Loop
    
    Range("A" & M) = "Avg"
    Range("A" & M & ":F" & M).Select   '合併儲存格
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
        End With
    
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
    
    
   ' Application.Dialogs(xlDialogSaveAs).Show
   
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
'-----------------------------------------------------------------------------------------------
        End With

    Application.Dialogs(xlDialogSaveAs).Show
    
    Range("A2").Select
    

End Sub

Private Sub OptionButton11_Click()

        Label1.Visible = True
        TextBox2.Visible = True
        TextBox2.Text = ""
End Sub
Private Sub OptionButton12_Click()

        Label1.Visible = False
        TextBox2.Visible = True
        TextBox2.Text = "Swipe"
        TextBox2.Visible = False
End Sub

Private Sub OptionButton13_Click()

        Label1.Visible = False
        TextBox2.Visible = True
        TextBox2.Text = "Swipe2"
        TextBox2.Visible = False
        

End Sub



Private Sub UserForm_Click()



End Sub
