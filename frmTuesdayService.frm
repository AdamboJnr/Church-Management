VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTuesdayService 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TuesdayFellowship"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "FELLOWSHIP PROGRAM"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8400
         TabIndex        =   17
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   44428
      End
      Begin MSAdodcLib.Adodc dtaTuesday 
         Height          =   495
         Left            =   1200
         Top             =   5760
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   "TUESDAY_FELLOWSHIP"
         Caption         =   "TUESDAY"
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
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         Picture         =   "frmTuesdayService.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton SAVE 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         Picture         =   "frmTuesdayService.frx":4888A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5640
         Width           =   1215
      End
      Begin VB.ComboBox cboGroupLeading 
         Height          =   315
         Left            =   8040
         TabIndex        =   14
         Top             =   4680
         Width           =   2295
      End
      Begin VB.ComboBox cboPreacher 
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   3600
         Width           =   2295
      End
      Begin VB.ComboBox cboDistrict 
         Height          =   315
         Left            =   8040
         TabIndex        =   11
         Top             =   2640
         Width           =   2295
      End
      Begin VB.ComboBox cboCongregation 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         TabIndex        =   9
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtReadings 
         DataSource      =   "dtaTuesday"
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtIntercessor 
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GROUP LEADING SERVICE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   13
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.C.E.A KISAUNI PARISH"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   10
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1815
         Left            =   120
         Picture         =   "frmTuesdayService.frx":4990C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TUESDAY FELLOWSHIP PROGRAM"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CONGREGATION"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INTERCESSOR"
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
         Left            =   120
         TabIndex        =   5
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "READINGS"
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
         Left            =   5640
         TabIndex        =   4
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREACHER"
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
         Left            =   120
         TabIndex        =   2
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblDistrict 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DISTRICT"
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
         Left            =   5640
         TabIndex        =   1
         Top             =   2640
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTuesdayService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cboCongregation.AddItem "BAMBURI"
cboCongregation.AddItem "MSHOMORONI"
cboCongregation.AddItem "BETHSAIDA"
cboCongregation.AddItem "KISAUNI"
 cboGroupLeading.AddItem "YOUTH"
 cboGroupLeading.AddItem "NORMAL SERVICE"
cboGroupLeading.AddItem "WOMANS GUILD"
 cboDistrict.AddItem "CONCORDIA"
 cboDistrict.AddItem "VIPINGO"
 cboDistrict.AddItem "BETHSAIDA"
 cboPreacher.AddItem "FRANCIS MUIRU"
 cboPreacher.AddItem "REV NGUGI"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub

Private Sub SAVE_Click()
If txtIntercessor.Text = "" Or txtReadings.Text = "" Or cboCongregation.Text = "" Or cboDistrict.Text = "" Then
MsgBox "Kindly fill in all details", vbCritical
Else
dtaTuesday.Recordset.AddNew
dtaTuesday.Recordset.Fields(0).Value = cboPreacher.Text
dtaTuesday.Recordset.Fields(1).Value = txtReadings.Text
dtaTuesday.Recordset.Fields(2).Value = txtIntercessor.Text
dtaTuesday.Recordset.Fields(3).Value = cboCongregation.Text
dtaTuesday.Recordset.Fields(4).Value = cboGroupLeading.Text
dtaTuesday.Recordset.Fields(5).Value = cboDistrict.Text
dtaTuesday.Recordset.Fields(6).Value = DTPicker1.Value
dtaTuesday.Recordset.Update
MsgBox "SERVICE SUCCESSFULLY SCHEDULED", vbInformation
cboCongregation.Text = ""
cboDistrict.Text = ""
cboGroupLeading.Text = ""
txtIntercessor.Text = ""
txtReadings.Text = ""
End If
End Sub
