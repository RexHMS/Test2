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
Set WSH = CreateObject("wscript.shell")  '�Ы�WSH ���إΩ�ާ@�R�O��
Set fso = CreateObject("scripting.filesystemobject")  '�Ы�FSO ���إΤ_�ާ@���
    path = N  '�]�w�_�l�ؿ�
    
    '�Q��WSH�B��R�O��R�O�N�_�l�ظ��U�����ͦ��C��M���X�_�_�l�ؿ��U��list.txt�C����,�õ��ݩR�O���槹����A�~���N�X
    WSH.Run Environ("comspec") & " /c dir """ & path & "\*.*"" /s/b/a-d>""" & path & "\list.txt""", 0, 1

    Tmp = "Img16bBkg"
    
    file = Dir(path & "\*", vbDirectory)
    

    Do While file <> ""
 
        If file <> "." And file <> ".." Then
        
            With fso.opentextfile(path & "\list.txt")  '�Q��FSO�ާ@���}�C����
                While Not .atendofstream  '�`��������C����
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
                                    If Dir(ppath, vbDirectory) = "" Then    '�ؿ����s�b��
                                        MkDir (ppath)                       '�إߥؿ�
                                    End If
                                    Name path & "\" & file As ppath & "\" & file
                               
                                Else
                                End If
                            Loop While file <> "" And first <> file
                        
                        End If
                        
                    End If
                Wend
                
                .Close  '�����C����
            End With
        End If
        file = Dir
    Loop
    
    If Dir(path & "\list.txt") <> "" Then
        Kill path & "\list.txt"  '�M���C����
    End If
    
    Set A = Nothing
    Set B = Nothing
    Set fso = Nothing  '�M�ųЫت�����
    Set WSH = Nothing
    'MsgBox "�B�z����"  '���ܴ��ܫH��
    
End Sub
