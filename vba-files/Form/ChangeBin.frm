VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeBin 
   Caption         =   "ChangeBin"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3990
   OleObjectBlob   =   "ChangeBin.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "ChangeBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()

BINAME = TB1.Text
'GoSub AutoSortingReport_Click
Range("a1") = TB1.Text
Unload Me

End Sub


