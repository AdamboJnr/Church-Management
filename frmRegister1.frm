VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "CREDENTIALS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.CheckBox UserType 
         BackColor       =   &H00400040&
         Caption         =   "Admin Priviledge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtUserName 
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Height          =   495
         Left            =   4200
         TabIndex        =   3
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtConfirm 
         DataSource      =   "dtaNewUser"
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "CREATE"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   1
         Top             =   4080
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc dtaNewUser 
         Height          =   495
         Left            =   5280
         Top             =   4440
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "select * from LOGIN"
         Caption         =   ""
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
      Begin VB.Label lblUserName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblPassword 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblConfirmPass 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RE-ENTER PASSWORD"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   2520
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCreate_Click()
    Dim strTypeOfUser As String
    If txtPassword.Text = "" Or txtConfirm.Text = "" Or txtUsername.Text = "" Then
        MsgBox "Kindly fill in all details", vbCritical
    ElseIf Len(txtPassword.Text) < 5 Then
        MsgBox "Password should be greater than 5 characters", vbCritical
    ElseIf txtConfirm.Text <> txtPassword.Text Then
        MsgBox "PASSWORDS DON'T MATCH", vbCritical
    ElseIf UserType.Value = Checked Then
        strTypeOfUser = "ADMIN"
        dtaNewUser.Recordset.AddNew
        dtaNewUser.Recordset.Fields(0).Value = txtUsername.Text
        dtaNewUser.Recordset.Fields(1).Value = txtPassword.Text
        dtaNewUser.Recordset.Fields(3).Value = strTypeOfUser
        dtaNewUser.Recordset.Update
        MsgBox "New user created succesfully"
        txtConfirm.Text = ""
        txtPassword.Text = ""
        txtUsername.Text = ""
    Else 'If UserType.Value <> Checked Then
        strTypeOfUser = "USER"
        dtaNewUser.Recordset.AddNew
        dtaNewUser.Recordset.Fields(0).Value = txtUsername.Text
        dtaNewUser.Recordset.Fields(1).Value = txtPassword.Text
        dtaNewUser.Recordset.Fields(3).Value = strTypeOfUser
        dtaNewUser.Recordset.Update
        MsgBox "New user created succesfully"
    txtConfirm.Text = ""
    txtPassword.Text = ""
    txtUsername.Text = ""
    
    End If
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)
Top = (Screen.Height - Height) / 2
  Left = (Screen.Width - Width) / 2
End Sub
