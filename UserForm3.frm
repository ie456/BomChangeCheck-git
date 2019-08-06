VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "ErrorMessage"
   ClientHeight    =   2220
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7464
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 180
    'Me.Left = Application.Left + Application.Width / 1.5 - Me.Width
    Me.Left = UserForm1.Left
End Sub

