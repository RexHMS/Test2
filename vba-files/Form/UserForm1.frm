VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���wBin �ɸ��|"
   ClientHeight    =   1485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CommandButton1_Click()

Dim FCol
Dim fso As Object
Dim FN$, WSH As Object
Dim FEqualCol, FEqualRow, LEqualCol, LEqualRow
Dim TotalEndRow, PaperEndRow
Dim FingerNum
Dim Col, Row
Dim Tmp, N
Dim path
Dim myPath
Dim myfile
Dim Str
Dim B
Dim A, P
Set WSH = CreateObject("wscript.shell")  '�Ы�WSH ���إΩ�ާ@�R�O��
Set fso = CreateObject("scripting.filesystemobject")  '�Ы�FSO ���إΤ_�ާ@���


    Unload Me
    
    N = TextBox1.Value

    Set WSH = CreateObject("wscript.shell")  '�Ы�WSH ���إΩ�ާ@�R�O��
    path = N
    
    '�R�� 0x8
    'WSH.Run Environ("comspec") & " /c del """ & path & "\*0x8*.bin """, 0, 1
    
Dim xlsApp   As New Excel.Application
    xlsApp.Workbooks.Add
    xlsApp.Workbooks(1).SaveAs (path & "\list.csv")
    xlsApp.Quit
    Set xlsApp = Nothing
    
    Workbooks.Open Filename:=path & "\list.csv"
    ActiveWorkbook.Save
        
    Application.ScreenUpdating = False
    
    mydir = Dir(ActiveWorkbook.path & "\*.*")
    
    Do While mydir <> ""
        Range("A1").Offset(I, 0) = mydir
        mydir = Dir()
        I = I + 1
    Loop
        
    Application.ScreenUpdating = False
    Worksheets(1).Copy after:=Worksheets(1)
    Worksheets(2).Select

    '�Ĥ@��  \
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="_", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1)), TrailingMinusNumbers:=True
    
    '�T�{�d��
    Range("A1").Select
    Selection.End(xlToRight).Select
    FCol = Selection.Column
   
   
    '��c01
    Range("A1").Select
    Cells.Find(What:="c01", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
    
    '����c01��m
    FEqualCol = ActiveCell.Column
    FEqualRow = ActiveCell.Row
    
    '������̫�@��C
    Selection.End(xlDown).Select
    TotalEndRow = ActiveCell.Row
    
    '��������A��������
    Cells(FEqualRow, FEqualCol).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    
    '��i�Ƴ̫�@��C
    Cells(FEqualRow, FEqualCol).Select
    Selection.End(xlDown).Select
    PaperEndRow = ActiveCell.Row
    
    FingerNum = TotalEndRow / PaperEndRow
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    '�W�h���|
    P = path
    P = Mid(P, 1, InStrRev(P, "\"))
    Debug.Print P
    
    '����j���l��
    Tmp = 1

    file = Dir(path & "\*", vbDirectory)
    

    Do While file <> ""
 
        If file <> "." And file <> ".." Then
        
            With Workbooks.Open(path & "\list.csv")  '�Q��FSO�ާ@���}�C����
                Sheets(1).Select
                'While Not .atendofstream  '�`��������C����
                    Str = Range("A1")
                    If Trim(Str) <> "" Then
                        
                        Do While Tmp < FingerNum + 1
                        
                            B = 1
                            Do While B < PaperEndRow + 1
                                Str = Range("A1")
                                A = Split(Str, "\")(UBound(Split(Str, "\")))
                                'b = Split(Str, "\")
                                'ReDim Preserve b(UBound(b) - 1)
                                'thePath = Join(b, "\")
                                'If path = thePath Then
                         
                                file = A
                                first = file
                                Do
                                    'If file Like "*" & Tmp & "*.bin" Then
                                    ppath = P & Tmp
                                    If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                        MkDir (ppath)                       '�إߥؿ�
                                    End If
                                
                                    pppath = P & Tmp & "\enroll"
                                    If Dir(pppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                        MkDir (pppath)                       '�إߥؿ�
                                    End If
                                
                                    ppppath = pppath & "\st"
                                    If Dir(ppppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                        MkDir (ppppath)                       '�إߥؿ�
                                    End If
                                
                                    Name path & "\" & file As ppppath & "\" & file
                                
                                    'Else
                                    'End If
                                Loop While file <> "" And first <> file
                                Sheets(1).Select
                                Rows(1).Delete Shift:=xlUp
                        
                                B = B + 1
                            Loop
                        
                            Tmp = Tmp + 1
                        Loop
                        'Else
                           'MsgBox "QQ"
                        'End If
                        
                    End If
                'Wend
                .Save
                .Close  '�����C����
            End With
        End If
        file = Dir
    Loop
    
    
    
    'If Dir(path & "\list.txt") <> "" Then
    Kill path & "\list.csv"  '�M���C����
        
    '�R����lenroll ��Ƨ�
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    With oFso
        .DeleteFolder (P & "\enroll")
    End With
    
    Dim myFolder As String
    Dim myNewFilePath As String
    Dim idstpath As String
    myFolder = P & "identify"    '�n��?�����?
    myNewFilePath = P & "1\"    '�n��?����m
    
    idstpath = myFolder & "\st"
    If Dir(idstpath, vbDirectory) = "" Then    '�ؿ����s�b��
        MkDir (idstpath)                       '�إߥؿ�
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.movefile myFolder & "\*.bin*", idstpath
    Set fso = Nothing  '�M�ųЫت�����
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder myFolder, myNewFilePath
    P = myNewFilePath & "identify"
    
    Name P As myNewFilePath & "verify"
    
    Set A = Nothing
    Set B = Nothing
    Set P = Nothing
    Set fso = Nothing  '�M�ųЫت�����
    Set WSH = Nothing
    'MsgBox "�B�z����"  '���ܴ��ܫH��

End Sub
