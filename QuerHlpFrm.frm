VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuerHlpFrm 
   Caption         =   "QuerHlpFrm"
   ClientHeight    =   2568
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6492
   OleObjectBlob   =   "QuerHlpFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QuerHlpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Me.Hide
End Sub
