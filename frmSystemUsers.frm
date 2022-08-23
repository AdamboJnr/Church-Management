VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSystemUsers 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtanewdetails 
      Height          =   375
      Left            =   5760
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "LOGIN"
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
   Begin MSAdodcLib.Adodc dtaLoginDetails 
      Height          =   375
      Left            =   7920
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "SELECT * FROM LOGIN"
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
   Begin VB.TextBox txtDateOfCreation 
      Height          =   405
      Left            =   7800
      TabIndex        =   15
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox CboUSERID 
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtUserPriviledge 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtPasssword 
      DataSource      =   "dtaLoginDetails"
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtxUsername 
      DataSource      =   "dtaLoginDetails"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   8280
      Picture         =   "frmSystemUsers.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "DATE CREATED"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "SYSTEM USERS PAGE"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "SELECT USER ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "USER PRIVILEDGE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "P.C.E.A KISAUNI CHURCH MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmSystemUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combo_populate()
While dtaLoginDetails.Recordset.EOF = False
    CboUSERID.AddItem dtaLoginDetails.Recordset.Fields(4).Value
    dtaLoginDetails.Recordset.MoveNext
Wend
End Sub

Private Sub CboUSERID_Click()
Dim UserId As Integer
UserId = CboUSERID.Text
dtaLoginDetails.Recordset.MoveFirst
dtaLoginDetails.Recordset.Find " [UserID]= " & UserId, 0, adSearchForward
If dtaLoginDetails.Recordset.EOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
ElseIf dtaLoginDetails.Recordset.Fields(4).Value = UserId Then
txtPasssword.Text = dtaLoginDetails.Recordset.Fields(1).Value
txtxUsername.Text = dtaLoginDetails.Recordset.Fields(0).Value
txtUserPriviledge.Text = dtaLoginDetails.Recordset.Fields(3).Value
End If


End Sub

Private Sub cmdPrevious_Click()
Dim UserId As Integer
UserId = CboUSERID.Text
dtaLoginDetails.Recordset.MoveFirst
dtaLoginDetails.Recordset.Find "UserID= " & UserId, 0, adSearchForward

If dtaLoginDetails.Recordset.Fields(4).Value = UserId Then
    dtaLoginDetails.Recordset.MovePrevious
    If dtaLoginDetails.Recordset.BOF = False Then
        'dtaLoginDetails.Recordset.MoveNext
        txtPasssword.Text = dtaLoginDetails.Recordset.Fields(1).Value
        CboUSERID.Text = dtaLoginDetails.Recordset.Fields(4).Value
        txtxUsername.Text = dtaLoginDetails.Recordset.Fields(0).Value
        txtUserPriviledge.Text = dtaLoginDetails.Recordset.Fields(3).Value
    Else
        dtaLoginDetails.Recordset.MoveLast
        CboUSERID.Text = dtaLoginDetails.Recordset.Fields(4).Value
        txtPasssword.Text = dtaLoginDetails.Recordset.Fields(1).Value
        txtxUsername.Text = dtaLoginDetails.Recordset.Fields(0).Value
        txtUserPriviledge.Text = dtaLoginDetails.Recordset.Fields(3).Value
    End If
End If


End Sub

Private Sub Command2_Click()
Dim UserId As Integer
UserId = CboUSERID.Text
dtaLoginDetails.Recordset.MoveFirst
dtaLoginDetails.Recordset.Find "UserID= " & UserId, 0, adSearchForward

If dtaLoginDetails.Recordset.Fields(4).Value = UserId Then
    dtaLoginDetails.Recordset.MoveNext
    If dtaLoginDetails.Recordset.EOF = False Then
        'dtaLoginDetails.Recordset.MoveNext
        txtPasssword.Text = dtaLoginDetails.Recordset.Fields(1).Value
        CboUSERID.Text = dtaLoginDetails.Recordset.Fields(4).Value
        txtxUsername.Text = dtaLoginDetails.Recordset.Fields(0).Value
        txtUserPriviledge.Text = dtaLoginDetails.Recordset.Fields(3).Value
    Else
        dtaLoginDetails.Recordset.MoveFirst
        CboUSERID.Text = dtaLoginDetails.Recordset.Fields(4).Value
        txtPasssword.Text = dtaLoginDetails.Recordset.Fields(1).Value
        txtxUsername.Text = dtaLoginDetails.Recordset.Fields(0).Value
        txtUserPriviledge.Text = dtaLoginDetails.Recordset.Fields(3).Value
    End If
End If

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
dtanewdetails.Recordset.Update
End Sub

Private Sub Form_Load()
Call combo_populate

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
