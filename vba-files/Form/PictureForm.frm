VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PictureForm 
   Caption         =   "PickupPicture"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14640
   OleObjectBlob   =   "PictureForm.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "PictureForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim fileNameObj As Variant
Dim aFile As Variant                            '�������WfileName�ϥ�
  
     '���}����ܮت�^�����W�A�O�@�ӥ����|���W�A��Ȥ]�i��OFalse�A�]��?��?Variant
Dim fullName As String
Dim Filename As String                         '?FileName����������?�W
Dim I As Integer
  
    fileNameObj = Application.GetOpenFilename("Excel ��� (*.csv),*.csv")
    '?��Windows��?���??��
    If fileNameObj <> False Then                   '�p�G������������?
        aFile = Split(fileNameObj, "")
  
        Filename = aFile(UBound(aFile))            '??���̦Z�@?����?���W
        fullName = aFile(0)
        TextBox1.Text = fullName
        For I = 1 To UBound(aFile)                 '�`?�X������?
            
        Next
    Else
        MsgBox "�п��Log"
    End If
    '�o��Excel����?
    allExcelFullPath = fullName
    '�o��Excel���W
    workbookName = Filename
End Sub

Sub CommandButton2_Click()
Dim fileNameObj As Variant
Dim aFile As Variant                            '??�A�������WfileName?�ϥ�
  
     '���}����ܮت�^�����W�A�O�@�ӥ����|���W�A��Ȥ]�i��OFalse�A�]��?��?Variant
Dim fullName As String
Dim Filename As String                         '?FileName����������?�W
Dim I As Integer
    PickupPic.wtf = 1
    If TextBox1.Value = "" Then
        MsgBox "�����J�ɮ�!!" & vbCrLf & "�Э��s����!!", 0 + 64
        End
    ElseIf OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False And OptionButton4.Value = False Then
        MsgBox "�����BIN!!" & vbCrLf & "�Э��s����!!", 0 + 64
        End
    ElseIf OptionButton4.Value = True And TextBox2.Value > "300" Then
        MsgBox "����JBIN!!" & vbCrLf & "�Э��s����!!", 0 + 64
        End
    ElseIf CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And _
           CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False And CheckBox9.Value = False And CheckBox10.Value = False And _
           CheckBox11.Value = False Then
        MsgBox "�����Image!!" & vbCrLf & "�Э��s����!!", 0 + 64
        End
    End If
    
    Call PickupPic.IMGCount(IMGF)
    
    Workbooks.Open (TextBox1.Text)
    MainSheet = ActiveWorkbook.name
    
    FolderPath = ThisWorkbook.path '���o��e���|
    'Application.DisplayAlerts = False '����ĵ�i����
    'ThisWorkbook.SaveAs '(FolderPath & "\�ɦW)
    'Application.Quit '������eExcel�ɮ�
    
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
