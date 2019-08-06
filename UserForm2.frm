VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   1680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Call U2_Module1.CreatDefaulFile(Replace(tempWorkBookName, ".xlsm", ""), tempWorkBookNamePath)
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 180
    Me.Left = Application.Left + Application.Width / 1.5 - Me.Width
End Sub


