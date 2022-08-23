VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSundayFellowship 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUNDAY SERVICE"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmScheduleSunday 
      BackColor       =   &H00400040&
      Caption         =   "SUNDAY SERVICE"
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin MSAdodcLib.Adodc SUNDAY_SERVICE 
         Height          =   495
         Left            =   6720
         Top             =   6480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
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
         RecordSource    =   "SUNDAY_FELLOWSHIP"
         Caption         =   "ORDER OF SERVICE"
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
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
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
         Left            =   2280
         TabIndex        =   20
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   19
         Top             =   6480
         Width           =   1935
      End
      Begin VB.ComboBox cboLeadingGroup 
         Height          =   315
         Left            =   6600
         TabIndex        =   18
         Top             =   5640
         Width           =   2055
      End
      Begin VB.ComboBox cboDeacons 
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txtPreacher 
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox txtServiceLeader 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox txt2ndReader 
         DataSource      =   "SUNDAY_SERVICE"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox txt1stReader 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txt1stReading 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txt2ndReading 
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GROUP LEADING "
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
         Left            =   4440
         TabIndex        =   17
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREACHER"
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
         Left            =   0
         TabIndex        =   14
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEACON"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SERVICE LEADER"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2nd Bible Reader"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1st Bible Reader"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.C.E.A KISAUNI PARISH"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUNDAY SERVICE PROGRAM"
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
         Left            =   2760
         TabIndex        =   5
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1st Reading"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2nd Reading"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   240
         Picture         =   "frmSundayFellowship.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSundayFellowship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub cmdSave_Click()
If txt1stReader.Text = "" Or txt1stReading.Text = "" Or txtPreacher.Text = "" Or txtServiceLeader.Text = "" Or cboLeadingGroup.Text = "" Or cboDeacons.Text = "" Then
    MsgBox "Kindly fill in all details", vbCritical
Else
    SUNDAY_SERVICE.Recordset.AddNew
    SUNDAY_SERVICE.Recordset.Fields(0).Value = txtPreacher.Text
    SUNDAY_SERVICE.Recordset.Fields(1).Value = cboDeacons.Text
    SUNDAY_SERVICE.Recordset.Fields(2).Value = cboLeadingGroup.Text
    SUNDAY_SERVICE.Recordset.Fields(3).Value = txt1stReading.Text
    SUNDAY_SERVICE.Recordset.Fields(4).Value = txt2ndReading.Text
    SUNDAY_SERVICE.Recordset.Fields(5).Value = txt1stReader.Text
    SUNDAY_SERVICE.Recordset.Fields(6).Value = txt2ndReader.Text
    SUNDAY_SERVICE.Recordset.Fields(7).Value = txtServiceLeader.Text
    SUNDAY_SERVICE.Recordset.Update
    MsgBox "SCHEDULE SUCCESFULLY CREATED AND SAVED", vbInformation
    txt1stReader.Text = ""
    txt1stReading.Text = ""
    txt2ndReader.Text = ""
    txt2ndReading.Text = ""
    txtPreacher.Text = ""
    txtServiceLeader.Text = ""
    cboDeacons.Text = ""
    cboLeadingGroup.Text = ""
End If
End Sub

Private Sub Form_Load()
cboDeacons.AddItem "Lilian Muli"
cboDeacons.AddItem "Alex Mathnge"
cboDeacons.AddItem "Phillip Mwangi"
cboDeacons.AddItem "Wachira Samuel"
cboDeacons.AddItem "Edgar Obare"

cboLeadingGroup.AddItem "YOUTH"
cboLeadingGroup.AddItem "P.C.M.F"
cboLeadingGroup.AddItem "WOMANS GUILD"
cboLeadingGroup.AddItem "CHRISTIAN EDUCATION"
cboLeadingGroup.AddItem "BRIDGADE"


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
