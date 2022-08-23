VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCreateGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "GROUPS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtGroupNumber 
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cboGroups 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblGroupNumber 
         Alignment       =   2  'Center
         Caption         =   "TOTAL NUMBER OF GROUPS"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblGroupList 
         Alignment       =   2  'Center
         Caption         =   "LIST OF GROUPS"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame frmGroups 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CREATE GROUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   10695
      Begin MSAdodcLib.Adodc GROUPS 
         Height          =   375
         Left            =   7560
         Top             =   4200
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
         RecordSource    =   "GROUPS"
         Caption         =   "GROUPS"
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
      Begin VB.CommandButton Cmdcreategroup 
         Caption         =   "CREATE GROUP"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         Picture         =   "frmCreateGroups.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox VICESECRETARY 
         Height          =   375
         Left            =   7920
         TabIndex        =   15
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox VICECHAIR 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox SECRETARY 
         Height          =   375
         Left            =   7920
         TabIndex        =   11
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox TREASURER 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox CONGREGATION 
         Height          =   315
         Left            =   7920
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox CHAIRPERSON 
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox PATRON 
         DataSource      =   "GROUPS"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox GroupName 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblGroupName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GROUP NAME"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V.SECRETARY"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V.CHAIRPERSON"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SECRETARY"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRESURER"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHAIRPERSON"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONGREGATION"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PATRON"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   120
      Picture         =   "frmCreateGroups.frx":1082
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmCreateGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmdcreategroup_Click()
If GroupName.Text = "" Or CHAIRPERSON.Text = "" Or PATRON.Text = "" Or VICECHAIR.Text = "" Or SECRETARY.Text = "" Or VICESECRETARY.Text = "" Or TREASURER.Text = "" Then
MsgBox "Kindly fill in all details to complete group registration", vbInformation
Else
GROUPS.Recordset.AddNew
GROUPS.Recordset.Fields(0).Value = GroupName.Text
GROUPS.Recordset.Fields(1).Value = CHAIRPERSON.Text
GROUPS.Recordset.Fields(2).Value = PATRON.Text
GROUPS.Recordset.Fields(3).Value = SECRETARY.Text
GROUPS.Recordset.Fields(4).Value = TREASURER.Text
GROUPS.Recordset.Fields(5).Value = VICECHAIR.Text
GROUPS.Recordset.Fields(6).Value = VICESECRETARY.Text
GROUPS.Recordset.Fields(7).Value = CONGREGATION.Text
GROUPS.Recordset.Update
MsgBox "GROUP SUCCESSFULLY CREATED", vbInformation
End If


End Sub

Private Sub Form_Load()
CONGREGATION.AddItem "KISAUNI"
CONGREGATION.AddItem "BETHSAIDA"
CONGREGATION.AddItem "BAMBURI"
CONGREGATION.AddItem "MSHOMORONI"


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
