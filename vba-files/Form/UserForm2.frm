VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "���w�ɮצ�m"
   ClientHeight    =   1380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub CommandButton2_Click()

Dim Coll, Roww As Integer
Dim Tmp, N
Dim path
Dim R As Range
Dim myfile
Dim myPath
Set WSH = CreateObject("wscript.shell")  '�Ы�WSH ���إΩ�ާ@�R�O��
Set fso = CreateObject("scripting.filesystemobject")  '�Ы�FSO ���إΤ_�ާ@���

    path = TextBox1.Value
    Unload Me
    
    WSH.Run Environ("comspec") & " /c dir """ & path & "\" & Tmp & "\*.png"" /s/b/a-d>""" & path & "\" & Tmp & "\list.csv""", 0, 1
    
    Workbooks.Open Filename:=path & "\list.csv"
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.path & "\" & Left(ActiveWorkbook.name, Len(ActiveWorkbook.name) - 3) & "xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False
        
    Application.ScreenUpdating = False
    
    myPath = Range("A1").Value
    myfile = Dir(myPath)
    
    Range("B1").Select
    Coll = Range("A1").Column
    Roww = Range("A1").Row
    
    Columns(2).ColumnWidth = 16
    path = path & "\"
    Do While myPath <> ""
        myfile = Dir(myPath)
    
        '3.�j�鴡�J�Ҧ��Y��

        '���J�Ϥ���
        
        'Range("J1") = path & myFile
        ActiveSheet.Shapes.AddPicture path & "\" & myfile, True, True, Selection.Left, Selection.Top, -1, -1
        'ActiveSheet.Shapes.AddPicture _
                    (I & "\image\BIN" & b & "\" & a & "_" & MyIMGArray(L) & ".bmp", True, True, R.Left, R.Top, -1, -1).Select
        Rows(Roww).RowHeight = 24
        
        '�j�M�U�@���ɮ�
        Roww = Roww + 1
        myPath = Cells(Roww, 1).Value
     
        Cells(Roww, 2).Select
    Loop
    
    filePath = path & "\list.csv"

    ' �R���ɮ�
    Kill (filePath)
    
    Range("A1").Select
    
    ActiveWorkbook.Save
    
End Sub

