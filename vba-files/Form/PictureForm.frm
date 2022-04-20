VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PictureForm 
   Caption         =   "PickupPicture"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14640
   OleObjectBlob   =   "PictureForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "PictureForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim fileNameObj As Variant
Dim aFile As Variant                            '提取文件名fileName使用
  
     '打開文件對話框返回的文件名，是一個全路徑文件名，其值也可能是False，因此?型?Variant
Dim fullName As String
Dim Filename As String                         '?FileName中提取的路?名
Dim I As Integer
  
    fileNameObj = Application.GetOpenFilename("Excel 文件 (*.csv),*.csv")
    '?用Windows打?文件??框
    If fileNameObj <> False Then                   '如果未按“取消”?
        aFile = Split(fileNameObj, "")
  
        Filename = aFile(UBound(aFile))            '??的最后一?元素?文件名
        fullName = aFile(0)
        TextBox1.Text = fullName
        For I = 1 To UBound(aFile)                 '循?合成全路?
            
        Next
    Else
        MsgBox "請選擇Log"
    End If
    '得到Excel全路?
    allExcelFullPath = fullName
    '得到Excel文件名
    workbookName = Filename
End Sub

Sub CommandButton2_Click()
Dim fileNameObj As Variant
Dim aFile As Variant                            '??，提取文件名fileName?使用
  
     '打開文件對話框返回的文件名，是一個全路徑文件名，其值也可能是False，因此?型?Variant
Dim fullName As String
Dim Filename As String                         '?FileName中提取的路?名
Dim I As Integer
    PickupPic.wtf = 1
    If TextBox1.Value = "" Then
        MsgBox "未載入檔案!!" & vbCrLf & "請重新執行!!", 0 + 64
        End
    ElseIf OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False And OptionButton4.Value = False Then
        MsgBox "未選擇BIN!!" & vbCrLf & "請重新執行!!", 0 + 64
        End
    ElseIf OptionButton4.Value = True And TextBox2.Value > "300" Then
        MsgBox "未填入BIN!!" & vbCrLf & "請重新執行!!", 0 + 64
        End
    ElseIf CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And _
           CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False And CheckBox9.Value = False And CheckBox10.Value = False And _
           CheckBox11.Value = False Then
        MsgBox "未選擇Image!!" & vbCrLf & "請重新執行!!", 0 + 64
        End
    End If
    
    Call PickupPic.IMGCount(IMGF)
    
    Workbooks.Open (TextBox1.Text)
    MainSheet = ActiveWorkbook.name
    
    FolderPath = ThisWorkbook.path '取得當前路徑
    'Application.DisplayAlerts = False '關閉警告視窗
    'ThisWorkbook.SaveAs '(FolderPath & "\檔名)
    'Application.Quit '關閉當前Excel檔案
    
    'Application.ActiveWorkbook.Path
    
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.path & "\" & Left(ActiveWorkbook.name, Len(ActiveWorkbook.name) - 3) & "xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False
    
    Call PickupPic.OptionBin(LASTB)
    
    Unload Me

End Sub

Private Sub OptionButton1_Click()
TextBox2.Visible = False
End Sub
Private Sub OptionButton2_Click()
TextBox2.Visible = False
End Sub
Private Sub OptionButton3_Click()
TextBox2.Visible = False
End Sub
Private Sub OptionButton4_Click()
TextBox2.Visible = True
End Sub
