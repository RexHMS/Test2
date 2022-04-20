VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Analytical_options 
   Caption         =   "Analytical options"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   OleObjectBlob   =   "Analytical_options.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Analytical_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CommandButton1_Click()

    If OptionButton2.Value = True Then
        Call Multiple.OpFcount
    
    End If
Unload Me
MainSheet = ActiveWorkbook.name

End Sub

Private Sub OptionButton1_Click()

    Frame1.Visible = False
    CheckBox1.Visible = False
    CheckBox2.Visible = False
    CheckBox3.Visible = False
    CheckBox4.Visible = False
    CheckBox5.Visible = False
    CheckBox6.Visible = False
    CheckBox7.Visible = False

End Sub

Private Sub OptionButton2_Click()

    Frame1.Visible = True
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    CheckBox3.Visible = True
    CheckBox4.Visible = True
    CheckBox5.Visible = True
    CheckBox6.Visible = True
    CheckBox7.Visible = True


End Sub






