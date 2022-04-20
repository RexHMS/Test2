VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "指定統計FRR的 Image 路徑"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205.001
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CheckBox7_Click()

End Sub

Sub CommandButton1_Click()
    If UserForm3.OptionButton15.Value = False And UserForm3.OptionButton16.Value = False Then
        MsgBox "請重新執行，並選擇 png or bin 檔!!"
        End
    End If

 
    Call Module3.CountPeople
    


End Sub

Private Sub OptionButton11_Click()

        Label2.Visible = True
        TextBox2.Visible = True
        TextBox2.Text = ""
End Sub
Private Sub OptionButton12_Click()

        Label2.Visible = False
        TextBox2.Visible = True
        TextBox2.Text = "Swipe"
        TextBox2.Visible = False
End Sub

Private Sub OptionButton13_Click()

        Label2.Visible = False
        TextBox2.Visible = True
        TextBox2.Text = "Swipe2"
        TextBox2.Visible = False
        

End Sub
Private Sub OptionButton15_Click()
    If OptionButton15.Value = True Then
            
        CheckBox11.Visible = True
    Else
        CheckBox11.Visible = False
        
    End If
End Sub


Private Sub OptionButton16_Click()
    If OptionButton16.Value = True Then
        OptionButton15.Value = False
        CheckBox11.Visible = False
    Else
        OptionButton15.Value = True
    End If
End Sub


