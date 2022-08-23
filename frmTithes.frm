VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTithes 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmTithe 
      BackColor       =   &H00400040&
      Caption         =   "TITHES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   2760
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc dtaTithes 
         Height          =   330
         Left            =   6600
         Top             =   5160
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
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
         RecordSource    =   "TITHES"
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   44425
      End
      Begin VB.CommandButton Command1 
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
         Height          =   615
         Left            =   3960
         Picture         =   "frmTithes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5040
         Width           =   1695
      End
      Begin VB.ComboBox cboDistrict 
         DataSource      =   "dtaTithes"
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtAmount 
         DataSource      =   "dtaTithes"
         Height          =   375
         Left            =   7680
         TabIndex        =   5
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cboPaymentMethods 
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lblFullName 
         Caption         =   "FULL  NAME"
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
         TabIndex        =   11
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.C.E.A KISAUNI PARISH"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label4 
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
         Left            =   5400
         TabIndex        =   6
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AMOUNT"
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
         Left            =   5400
         TabIndex        =   4
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblPaymentMethod 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Payment Method"
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
         Left            =   120
         TabIndex        =   2
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   240
         Picture         =   "frmTithes.frx":1082
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TITHE RECORDING PAGE"
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
         Left            =   3240
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmTithes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cboDistrict.Text = "" Or cboPaymentMethods.Text = "" Then
MsgBox "Selct the District and payment method", vbExclamation
Else
dtaTithes.Recordset.AddNew
dtaTithes.Recordset.Fields(0).Value = txtName.Text
dtaTithes.Recordset.Fields(1).Value = cboDistrict.Text
dtaTithes.Recordset.Fields(2).Value = cboPaymentMethods.Text
dtaTithes.Recordset.Fields(3).Value = DTPicker1.Value
dtaTithes.Recordset.Fields(4).Value = txtAmount.Text
dtaTithes.Recordset.Update
MsgBox "TITHE SUCCESFULLY SAVED"
cboDistrict.Text = ""
cboPaymentMethods.Text = ""
txtAmount.Text = ""
txtName.Text = ""

End If

End Sub

Private Sub Form_Load()
cboDistrict.AddItem "mtopanga"
cboDistrict.AddItem "mwembeni"
cboDistrict.AddItem "concordia."
cboDistrict.AddItem "kizimani"
cboPaymentMethods.AddItem "cash"
cboPaymentMethods.AddItem "mpesa"
cboPaymentMethods.AddItem "equity bank"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
