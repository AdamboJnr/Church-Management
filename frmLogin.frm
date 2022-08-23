VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "LOGIN FORM"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCreatAcc_Click()
frmRegister.Show
End Sub

Private Sub login_Click()
frmMain.Show
frmLogin.Hide
End Sub

