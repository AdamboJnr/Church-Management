VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddNew 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   11400
      Picture         =   "frmRegister.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlImages 
      Left            =   12600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dtaYouthRecords 
      Height          =   375
      Left            =   480
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   "SELECT * FROM YOUTH_MEMBERS"
      Caption         =   "Records"
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
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "SAVE"
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
      Left            =   5160
      Picture         =   "frmRegister.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Frame frmChurchInfo 
      Caption         =   "GROUP  INFORMATION"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   11055
      Begin VB.ComboBox cboDistrictElder 
         Height          =   315
         ItemData        =   "frmRegister.frx":2104
         Left            =   8880
         List            =   "frmRegister.frx":2106
         TabIndex        =   30
         Top             =   1440
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DOCommisioning 
         Height          =   375
         Left            =   8880
         TabIndex        =   27
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   44426
      End
      Begin MSComCtl2.DTPicker DOBaptism 
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   44426
      End
      Begin VB.ComboBox cboDistrict 
         Height          =   315
         ItemData        =   "frmRegister.frx":2108
         Left            =   3000
         List            =   "frmRegister.frx":210A
         TabIndex        =   22
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtMemberID 
         Height          =   495
         Left            =   8880
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboCongregation 
         Height          =   315
         ItemData        =   "frmRegister.frx":210C
         Left            =   3240
         List            =   "frmRegister.frx":210E
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblDOC 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE OF COMMISSIONING"
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
         Left            =   6120
         TabIndex        =   21
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblElder 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DISTRICT  SERVED"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblDistrictElder 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DISTRICT ELDER"
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
         Left            =   6000
         TabIndex        =   18
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblBaptism 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE OF BAPTISM"
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
         Left            =   480
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblMemberID 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREVIOUS CHURCH SERVED"
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
         Left            =   6000
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERSONAL INFORMATION"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10935
      Begin VB.TextBox txtStatus 
         Height          =   375
         Left            =   8520
         TabIndex        =   32
         Top             =   3120
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DOB 
         Height          =   255
         Left            =   8280
         TabIndex        =   25
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   44426
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2880
         TabIndex        =   24
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtFirstName 
         DataSource      =   "dtaYouthRecords"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtResidence 
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmRegister.frx":2110
         Left            =   8400
         List            =   "frmRegister.frx":211D
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtIDNum 
         Height          =   525
         Left            =   2880
         TabIndex        =   2
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtPhoneNum 
         DataSource      =   "dtaYouthRecords"
         Height          =   495
         Left            =   8400
         TabIndex        =   1
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label STATUS 
         Caption         =   "MEMBER R/SHIP STATUS"
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
         Left            =   5880
         TabIndex        =   31
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATION"
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
         TabIndex        =   23
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblFirstame 
         Alignment       =   2  'Center
         Caption         =   "FULL NAME"
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
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblDOB 
         Alignment       =   2  'Center
         Caption         =   "DATE OF BIRTH"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblAge 
         Alignment       =   2  'Center
         Caption         =   " RESIDENCE"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblGender 
         Alignment       =   2  'Center
         Caption         =   "GENDER"
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
         Left            =   5880
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblIDNum 
         Alignment       =   2  'Center
         Caption         =   "ID NUMBER"
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
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblPhoneNum 
         Alignment       =   2  'Center
         Caption         =   "PHONE NUMBER"
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
         Left            =   5760
         TabIndex        =   6
         Top             =   2400
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "P.C.E.A KISAUNI MEMBER REGISTRATION FORM"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   29
      Top             =   240
      Width           =   6375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   11160
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCongregation_Click()
If cboCongregation.Text = "BAMBURI" Then
cboDistrict.Clear
cboDistrictElder.Clear
    cboDistrict.AddItem "nsbsje"
    cboDistrict.AddItem "ngvghg"
    cboDistrict.AddItem "ghffcg"
    cboDistrict.AddItem "ghhdff"
    cboDistrict.AddItem "hghghf"
    cboDistrictElder.AddItem "Elder AAWE"
    cboDistrictElder.AddItem "Elder ADSE"
    cboDistrictElder.AddItem "Elder ADDR"
    cboDistrictElder.AddItem "Elder ADNR"
 ElseIf cboCongregation.Text = "KISAUNI" Then
 cboDistrict.Clear
 cboDistrictElder.Clear
    cboDistrict.AddItem "nsbsje"
    cboDistrict.AddItem "nfvgghg"
    cboDistrict.AddItem "gxgcg"
    cboDistrict.AddItem "dff"
    cboDistrict.AddItem "xfccf"
    cboDistrictElder.AddItem "Elder Maguta"
    cboDistrictElder.AddItem "Elder JANE"
    cboDistrictElder.AddItem "Elder SUSAN"
    cboDistrictElder.AddItem "Elder PHILL"
 ElseIf cboCongregation.Text = "BETHSAIDA" Then
 cboDistrict.Clear
 cboDistrictElder.Clear
    cboDistrict.AddItem "CONCORDIA"
    cboDistrict.AddItem "VIKWATANI"
    cboDistrictElder.AddItem "Elder ANN"
    cboDistrictElder.AddItem "Elder SAMUEL WACHIRA II"
  ElseIf cboCongregation.Text = "MSHOMORONI" Then
    cboDistrict.Clear
    cboDistrictElder.Clear
  
  cboDistrict.AddItem "nsbsje"
    cboDistrict.AddItem "Masasuchets"
    cboDistrict.AddItem "kizingoni"
    cboDistrict.AddItem "Likoni"
    cboDistrict.AddItem "Qwerty"
    cboDistrictElder.AddItem "Elder Reginah"
    cboDistrictElder.AddItem "Elder JANE"
    cboDistrictElder.AddItem "Elder Jnae"
    cboDistrictElder.AddItem "Elder Sam"
End If
End Sub

Private Sub cmdSubmit_Click()
Dim lngPhoneNumber As Long
If txtFirstName.Text = "" Or txtResidence.Text = "" Or txtPhoneNum.Text = "" Or cboGender.Text = "" Or cboCongregation.Text = "" Or txtIDNum.Text = "" Or cboDistrict.Text = "" Then
MsgBox "KINDLY FILL IN ALL THE  DETAILS", vbInformation
ElseIf txtPhoneNum.Text < 9 Then
MsgBox "phone number can not be less than 10 digits", vbExclamatio
Else
    lngPhoneNumber = txtPhoneNum.Text
    dtaYouthRecords.Recordset.AddNew
    dtaYouthRecords.Recordset.Fields(1).Value = txtFirstName.Text
    dtaYouthRecords.Recordset.Fields(2).Value = txtResidence.Text
    dtaYouthRecords.Recordset.Fields(3).Value = txtIDNum.Text
    dtaYouthRecords.Recordset.Fields(4).Value = DOB
    dtaYouthRecords.Recordset.Fields(5).Value = cboGender.Text
    dtaYouthRecords.Recordset.Fields(6).Value = lngPhoneNumber
    dtaYouthRecords.Recordset.Fields(7).Value = cboCongregation.Text
    dtaYouthRecords.Recordset.Fields(8).Value = DOBaptism
    dtaYouthRecords.Recordset.Fields(9).Value = cboDistrict.Text
    dtaYouthRecords.Recordset.Fields(10).Value = cboDistrictElder.Text
    dtaYouthRecords.Recordset.Fields(11).Value = DOCommisioning.Value
    dtaYouthRecords.Recordset.Update
    MsgBox "MEMBER SUCCCESSFULLY SAVED", vbInformation
    txtDOB.Text = ""
    txtFirstName.Text = ""
    txtIDNum = ""
    txtMemberID = ""
    txtPhoneNum.Text = ""
End If
End Sub

Private Sub Command1_Click()
cdlImages.Filter = "picture file | *.JPG"
cdlImages.ShowOpen
If cdlImages.FileName <> "" Then
Image1.Picture = LoadPicture(cdlImages.FileName)
End If
End Sub

Private Sub Form_Load()
    cboCongregation.AddItem "BAMBURI"
    cboCongregation.AddItem "KISAUNI"
    cboCongregation.AddItem "BETHSAIDA"
    cboCongregation.AddItem "MSHOMORONI"

  
  
 
 
 
 
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
