VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "指定Bin 檔路徑"
   ClientHeight    =   1485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所屬視窗中央
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
Set WSH = CreateObject("wscript.shell")  '創建WSH 項目用於操作命令行
Set fso = CreateObject("scripting.filesystemobject")  '創建FSO 項目用于操作文件


    Unload Me
    
    N = TextBox1.Value

    Set WSH = CreateObject("wscript.shell")  '創建WSH 項目用於操作命令行
    path = N
    
    '刪除 0x8
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

    '第一切  \
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="_", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1)), TrailingMinusNumbers:=True
    
    '確認範圍
    Range("A1").Select
    Selection.End(xlToRight).Select
    FCol = Selection.Column
   
   
    '找c01
    Range("A1").Select
    Cells.Find(What:="c01", after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
    
    '紀錄c01位置
    FEqualCol = ActiveCell.Column
    FEqualRow = ActiveCell.Row
    
    '找全部最後一顆C
    Selection.End(xlDown).Select
    TotalEndRow = ActiveCell.Row
    
    '選取全部，移除重複
    Cells(FEqualRow, FEqualCol).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    
    '找張數最後一顆C
    Cells(FEqualRow, FEqualCol).Select
    Selection.End(xlDown).Select
    PaperEndRow = ActiveCell.Row
    
    FingerNum = TotalEndRow / PaperEndRow
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    '上層路徑
    P = path
    P = Mid(P, 1, InStrRev(P, "\"))
    Debug.Print P
    
    '手指迴圈初始值
    Tmp = 1

    file = Dir(path & "\*", vbDirectory)
    

    Do While file <> ""
 
        If file <> "." And file <> ".." Then
        
            With Workbooks.Open(path & "\list.csv")  '利用FSO操作打開列表文件
                Sheets(1).Select
                'While Not .atendofstream  '循環取直到列表文件末
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
                                    If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                        MkDir (ppath)                       '建立目錄
                                    End If
                                
                                    pppath = P & Tmp & "\enroll"
                                    If Dir(pppath, vbDirectory) = "" Then    '目錄不存在時
                                        MkDir (pppath)                       '建立目錄
                                    End If
                                
                                    ppppath = pppath & "\st"
                                    If Dir(ppppath, vbDirectory) = "" Then    '目錄不存在時
                                        MkDir (ppppath)                       '建立目錄
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
                .Close  '關閉列表文件
            End With
        End If
        file = Dir
    Loop
    
    
    
    'If Dir(path & "\list.txt") <> "" Then
    Kill path & "\list.csv"  '清除列表文件
        
    '刪除原始enroll 資料夾
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    With oFso
        .DeleteFolder (P & "\enroll")
    End With
    
    Dim myFolder As String
    Dim myNewFilePath As String
    Dim idstpath As String
    myFolder = P & "identify"    '要移?的文件?
    myNewFilePath = P & "1\"    '要移?的位置
    
    idstpath = myFolder & "\st"
    If Dir(idstpath, vbDirectory) = "" Then    '目錄不存在時
        MkDir (idstpath)                       '建立目錄
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.movefile myFolder & "\*.bin*", idstpath
    Set fso = Nothing  '清空創建的項目
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder myFolder, myNewFilePath
    P = myNewFilePath & "identify"
    
    Name P As myNewFilePath & "verify"
    
    Set A = Nothing
    Set B = Nothing
    Set P = Nothing
    Set fso = Nothing  '清空創建的項目
    Set WSH = Nothing
    'MsgBox "處理完成"  '提示提示信息

End Sub
