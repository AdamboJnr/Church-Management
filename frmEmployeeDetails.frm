VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmployeeDetails 
   BackColor       =   &H80000011&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMPLOYEE DETAILS"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE DETAILS"
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
         Left            =   3960
         Picture         =   "frmEmployeeDetails.frx":0000
         TabIndex        =   23
         Top             =   7440
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc dtaEmp 
         Height          =   375
         Left            =   360
         Top             =   7320
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         RecordSource    =   "EMPLOYEE_RECORDS"
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
      Begin VB.ComboBox cboGender 
         Height          =   315
         Left            =   7080
         TabIndex        =   22
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtBANKNumber 
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         DataSource      =   "dtaEmp"
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox txtPhoneNumber 
         DataSource      =   "dta"
         Height          =   375
         Left            =   7080
         TabIndex        =   15
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox txtHomeResidence 
         DataSource      =   "dtaEmployeeRecords"
         Height          =   375
         Left            =   7080
         TabIndex        =   13
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtKRAPIN 
         DataSource      =   "dta"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox txtReferreContacts 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox txtCurResidence 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   6480
         Picture         =   "frmEmployeeDetails.frx":4888A
         Top             =   1320
         Width           =   720
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lblBanner 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.C.E.A KISAUNI KISAUNI MANAGEMENT STSTEM"
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
         Left            =   2400
         TabIndex        =   21
         Top             =   480
         Width           =   5775
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   120
         Picture         =   "frmEmployeeDetails.frx":91114
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "EMPLOYEE RECORDS"
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
         Left            =   3000
         TabIndex        =   20
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "BANK ACCOUNT"
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
         Left            =   5040
         TabIndex        =   18
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "ADDRESS"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "PHONE NUMBER"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "HOME RESIDENCE"
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
         Left            =   5160
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Left            =   5160
         TabIndex        =   11
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "KRA PIN"
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
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "REFEREE CONTACTS"
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
         TabIndex        =   6
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label3 
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
         TabIndex        =   5
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "CURENT RESIDENCE"
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
         TabIndex        =   3
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEmployeeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSave_Click()

End Sub

Private Sub Command1_Click()
Dim BankNumber As Long


Dim refereecont As Long
If txtName.Text = "" Or txtAddress.Text = "" Or txtCurResidence.Text = "" Or txtHomeResidence.Text = "" Or txtIDNumber.Text = "" Or txtKRAPIN.Text = "" Or txtBANKNumber.Text = "" Or txtReferreContacts.Text = "" Or cboGender.Text = "" Then
MsgBox "KINDLY FILL IN ALL THE DETAILS", vbExclamation
ElseIf txtIDNumber < 8 Then
MsgBox "ID Number can not be less that 8 digits!", vbExclamation
ElseIf txtPhone2 And txtPhoneNumber < 10 Then
MsgBox "Check Phone Number", vbInformation
ElseIf txtBANKNumber.Text < 10 Then
MsgBox "Bank Account cant be less than 13 digits", vbExclamation
Else
refereecont = txtReferreContacts.Text
BankNumber = txtBANKNumber.Text
dtaEmp.Recordset.AddNew
dtaEmp.Recordset.Fields(1).Value = txtName.Text
dtaEmp.Recordset.Fields(2).Value = txtCurResidence.Text
dtaEmp.Recordset.Fields(3).Value = txtIDNumber.Text
dtaEmp.Recordset.Fields(4).Value = refereecont
dtaEmp.Recordset.Fields(5).Value = txtKRAPIN
dtaEmp.Recordset.Fields(6).Value = cboGender.Text
dtaEmp.Recordset.Fields(7).Value = txtHomeResidence.Text
dtaEmp.Recordset.Fields(8).Value = txtPhoneNumber.Text
dtaEmp.Recordset.Fields(9).Value = txtAddress.Text
dtaEmp.Recordset.Fields(10).Value = BankNumber
 dtaEmp.Recordset.Update
MsgBox "EMPLOYEE RECORD SUCCESSFULLY SAVED", vbInformation
txtAddress.Text = ""
txtBANKNumber.Text = ""
txtCurResidence.Text = ""
txtHomeResidence = ""
txtIDNumber.Text = ""
txtKRAPIN.Text = ""
txtName.Text = ""
cboGender.Text = ""
End If
End Sub

Private Sub Form_Load()
cboGender.AddItem "MALE"
cboGender.AddItem "FEMALE"


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
