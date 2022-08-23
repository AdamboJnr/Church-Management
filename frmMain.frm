VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000B&
   Caption         =   "LOGIN"
   ClientHeight    =   7185
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin MSAdodcLib.Adodc dtaLogin 
         Height          =   495
         Left            =   2280
         Top             =   6240
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CHURCH MANAGEMENT DATABASE.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CHURCH MANAGEMENT DATABASE.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT *  FROM LOGIN"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdForgotPass 
         Caption         =   "FORGOT PASSWORD?"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         DataSource      =   "dtaLogin"
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         DataSource      =   "dtaLogin"
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   4920
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "USER LOGIN"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label lblUserName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   2640
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreatAcc_Click()
Create = MsgBox("You are about to create a new user do you wish to proceed?", vbExclamation + vbYesNo, "Add confirm")
If Create = vbYes Then
frmRegister.Show
frmMain.Hide
End If

End Sub
Private Sub cmdForgotPass_Click()
answer = MsgBox("Kindly contact admin for assistance", vbExclamation + vbYesNo)
End Sub
Private Sub cmdLogin_Click()
If txtPassword.Text = "" Or txtUsername.Text = "" Then
MsgBox "Kindly enter a username and password", vbInformation
Else
   dtaLogin.RecordSource = " select * from LOGIN where USERNAME = '" + txtUsername.Text + "' and PASSWORD = '" + txtPassword.Text + "' "
   dtaLogin.Refresh
    If dtaLogin.Recordset.EOF = True Then
        MsgBox "Wrong Username Or Password!", vbCritical
        txtUsername.SetFocus
    Else
        typeOfUser = dtaLogin.Recordset.Fields(3).Value
        Unload Me
        frmHomePage.Show
    End If
End If
End Sub

Private Sub Form_Load()
 Top = (Screen.Height - Height) / 2
  Left = (Screen.Width - Width) / 2
txtPassword.Text = ""
txtUsername.Text = ""
End Sub
