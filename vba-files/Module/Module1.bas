Attribute VB_Name = "Module1"
Option Explicit

Sub Mod_one_Click()

Dim MainSheet As String
Dim Filename As String

    UserForm1.Show

End Sub

Sub testTT(Tmp, N)

Dim path$, file$, first$, ppath$, Str$
Dim B
Dim A
Dim fso As Object
Dim FN$, WSH As Object
'Option Compare Binary
Set WSH = CreateObject("wscript.shell")  '創建WSH 項目用於操作命令行
Set fso = CreateObject("scripting.filesystemobject")  '創建FSO 項目用于操作文件
    path = N  '設定起始目錄
    
    '利用WSH運行命令行命令將起始目路下的文件生成列表清單輸出于起始目錄下的list.txt列表文件中,並等待命令執行完成後再繼續行代碼
    WSH.Run Environ("comspec") & " /c dir """ & path & "\*.*"" /s/b/a-d>""" & path & "\list.txt""", 0, 1

    Tmp = "Img16bBkg"
    
    file = Dir(path & "\*", vbDirectory)
    

    Do While file <> ""
 
        If file <> "." And file <> ".." Then
        
            With fso.opentextfile(path & "\list.txt")  '利用FSO操作打開列表文件
                While Not .atendofstream  '循環取直到列表文件末
                    Str = .readline
                    If Trim(Str) <> "" Then
                    
                        A = Split(Str, "\")(UBound(Split(Str, "\")))
                        B = Split(Str, "\")
                        ReDim Preserve B(UBound(B) - 1)
                        thePath = Join(B, "\")
                        
                        If path = thePath Then
                         
                            file = A
                            first = file
                            Do
                                If file Like "*" & Tmp & "*.bin" Then
                                    ppath = path & "\" & Tmp
                                    If Dir(ppath, vbDirectory) = "" Then    '目錄不存在時
                                        MkDir (ppath)                       '建立目錄
                                    End If
                                    Name path & "\" & file As ppath & "\" & file
                               
                                Else
                                End If
                            Loop While file <> "" And first <> file
                        
                        End If
                        
                    End If
                Wend
                
                .Close  '關閉列表文件
            End With
        End If
        file = Dir
    Loop
    
    If Dir(path & "\list.txt") <> "" Then
        Kill path & "\list.txt"  '清除列表文件
    End If
    
    Set A = Nothing
    Set B = Nothing
    Set fso = Nothing  '清空創建的項目
    Set WSH = Nothing
    'MsgBox "處理完成"  '提示提示信息
    
End Sub
