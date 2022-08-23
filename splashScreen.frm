VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form splashScreen 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Splash"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10260
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splashScreen.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   9360
      Top             =   1320
   End
   Begin ComctlLib.ProgressBar StatusBar 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   0
      Max             =   105
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   120
      Picture         =   "splashScreen.frx":CE881
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P.C.E.A KISAUNI CHURCH MANAGEMENT SYSTEM  V1.0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   7455
   End
End
Attribute VB_Name = "splashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
StatusBar.Value = StatusBar.Value + 5
If StatusBar.Value <= 40 Then
Label3.Caption = "Welcome to P.C.E.A KISAUNI MANAGEMENT SYSTEM..Setup is starting up"
ElseIf StatusBar.Value >= 41 And StatusBar.Value <= 59 Then
Label3.Caption = "Getting your workspace set up..Almost there"
ElseIf StatusBar.Value >= 60 And StatusBar.Value <= 84 Then
 Label3.Caption = "UNPACKING APPROPRIATE FILES FOR YOU...!"
ElseIf StatusBar.Value >= 85 Then
Label3.Caption = "PREPARING ACCESS INTO THE SYSTEM...!"
End If
Label4.Caption = StatusBar.Value & "%"
If (StatusBar.Value = StatusBar.Max) Then
Timer1.Enabled = False
Unload Me
MsgBox "Welcome to P.C.E.A Kisauni church management system"
frmLogin.Show
End If
'Private Sub tmrSplashScreen_Timer()
'ProgressBar1.Value = ProgressBar1.Value + 5
'If ProgressBar1.Value <= 40 Then
 '   lbstatusbar.Caption = "WELCOME TO THE MWAMBA HOUSING MANAGEMENT SYSTEM....KINDLY WAIT AS THE SYSTEM LOADS"
'ElseIf ProgressBar1.Value >= 41 And ProgressBar1.Value <= 59 Then
 '   lbstatusbar.Caption = "STARTING THE DATABASE SERVER....ALMOST THERE!"
'ElseIf ProgressBar1.Value >= 60 And ProgressBar1.Value <= 84 Then
 '   lbstatusbar.Caption = "UNPACKING APPROPRIATE FILES FOR YOU...!"
'ElseIf ProgressBar1.Value >= 85 Then
 '   lbstatusbar.Caption = "PREPARING ACCESS INTO THE SYSTEM...!"
'End If
'lblstatus.Caption = ProgressBar1.Value & "%"
'If ProgressBar1.Value = ProgressBar1.Max Then
'Unload Me
'frmLoginForm.Show
'End If
'End Sub

End Sub
