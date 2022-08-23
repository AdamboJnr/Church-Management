VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form details 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin MSAdodcLib.Adodc dtaEmployeeID 
         Height          =   330
         Left            =   240
         Top             =   8160
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         RecordSource    =   "SELECT * FROM EMPLOYEE_RECORDS"
         Caption         =   "empDetails"
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
         TabIndex        =   15
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox txtBANKNumber 
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   7440
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   7080
         TabIndex        =   13
         Top             =   6480
         Width           =   2295
      End
      Begin VB.TextBox txtPhoneNumber 
         DataSource      =   "dtaEmployeeID"
         Height          =   375
         Left            =   7080
         TabIndex        =   12
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox txtHomeResidence 
         DataSource      =   "dtaEmployeeID"
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox txtKRAPIN 
         DataSource      =   "dta"
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   7440
         Width           =   2295
      End
      Begin VB.TextBox txtReferreContacts 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   6480
         Width           =   2295
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox txtCurResidence 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "PREVIOUS"
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "NEXT"
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   8520
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   5640
         TabIndex        =   3
         Top             =   8520
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   7440
         TabIndex        =   2
         Top             =   8520
         Width           =   1335
      End
      Begin VB.ComboBox cboEmployeeId 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   6480
         Picture         =   "details.frx":0000
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
         TabIndex        =   28
         Top             =   480
         Width           =   5775
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   120
         Picture         =   "details.frx":4888A
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   7440
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
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   6480
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
         TabIndex        =   24
         Top             =   5520
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
         Left            =   5040
         TabIndex        =   23
         Top             =   4680
         Width           =   1935
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
         Left            =   5040
         TabIndex        =   22
         Top             =   3720
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
         TabIndex        =   21
         Top             =   7440
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
         TabIndex        =   20
         Top             =   6480
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
         TabIndex        =   19
         Top             =   5520
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
         TabIndex        =   18
         Top             =   4680
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
         Left            =   240
         TabIndex        =   17
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblSELECT 
         Caption         =   "SELECT EMPLOYEE ID"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   2055
      End
   End
End
Attribute VB_Name = "details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub populate_combo()
    While dtaEmployeeID.Recordset.EOF = False
        cboEmployeeId.AddItem dtaEmployeeID.Recordset.Fields(0).Value
        dtaEmployeeID.Recordset.MoveNext
    Wend
End Sub

Private Sub cboEmployeeId_Click()
Dim employeeId As Integer
employeeId = cboEmployeeId.Text
dtaEmployeeID.Recordset.MoveFirst
dtaEmployeeID.Recordset.Find " [EMPLOYEE ID]= " & employeeId, 0, adSearchForward
If dtaEmployeeID.Recordset.EOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
ElseIf dtaEmployeeID.Recordset.Fields(0).Value = employeeId Then
txtName.Text = dtaEmployeeID.Recordset.Fields(1).Value
txtCurResidence.Text = dtaEmployeeID.Recordset.Fields(2).Value
txtIDNumber.Text = dtaEmployeeID.Recordset.Fields(3).Value
txtReferreContacts.Text = dtaEmployeeID.Recordset.Fields(4).Value
txtKRAPIN.Text = dtaEmployeeID.Recordset.Fields(5).Value
cboGender.Text = dtaEmployeeID.Recordset.Fields(6).Value
txtHomeResidence.Text = dtaEmployeeID.Recordset.Fields(7).Value
txtPhoneNumber.Text = dtaEmployeeID.Recordset.Fields(8).Value
txtAddress.Text = dtaEmployeeID.Recordset.Fields(9).Value
txtBANKNumber.Text = dtaEmployeeID.Recordset.Fields(10).Value
End If
End Sub

Private Sub cmdPrevious_Click()
Dim employeeId As Integer
employeeId = cboEmployeeId.Text
dtaEmployeeID.Recordset.MovePrevious
dtaEmployeeID.Recordset.Find " [EMPLOYEE ID]= " & employeeId, 0, adSearchForward
If dtaEmployeeID.Recordset.BOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
If dtaEmployeeID.Recordset.Fields(0).Value = employeeId Then
    dtaEmployeeID.Recordset.MovePrevious
    If dtaEmployeeID.Recordset.BOF = False Then
        'dtaEmployeeID.Recordset.MoveNext
        cboEmployeeId.Text = dtaEmployeeID.Recordset.Fields(0).Value
        txtName.Text = dtaEmployeeID.Recordset.Fields(1).Value
        txtCurResidence.Text = dtaEmployeeID.Recordset.Fields(2).Value
        txtIDNumber.Text = dtaEmployeeID.Recordset.Fields(3).Value
        txtReferreContacts.Text = dtaEmployeeID.Recordset.Fields(4).Value
        txtKRAPIN.Text = dtaEmployeeID.Recordset.Fields(5).Value
        cboGender.Text = dtaEmployeeID.Recordset.Fields(6).Value
        txtHomeResidence.Text = dtaEmployeeID.Recordset.Fields(7).Value
        txtPhoneNumber.Text = dtaEmployeeID.Recordset.Fields(8).Value
        txtAddress.Text = dtaEmployeeID.Recordset.Fields(9).Value
        txtBANKNumber.Text = dtaEmployeeID.Recordset.Fields(10).Value
    Else
        dtaEmployeeID.Recordset.MoveLast
        
         cboEmployeeId.Text = dtaEmployeeID.Recordset.Fields(0).Value
        txtName.Text = dtaEmployeeID.Recordset.Fields(1).Value
        txtCurResidence.Text = dtaEmployeeID.Recordset.Fields(2).Value
        txtIDNumber.Text = dtaEmployeeID.Recordset.Fields(3).Value
        txtReferreContacts.Text = dtaEmployeeID.Recordset.Fields(4).Value
        txtKRAPIN.Text = dtaEmployeeID.Recordset.Fields(5).Value
        cboGender.Text = dtaEmployeeID.Recordset.Fields(6).Value
        txtHomeResidence.Text = dtaEmployeeID.Recordset.Fields(7).Value
        txtPhoneNumber.Text = dtaEmployeeID.Recordset.Fields(8).Value
        txtAddress.Text = dtaEmployeeID.Recordset.Fields(9).Value
        txtBANKNumber.Text = dtaEmployeeID.Recordset.Fields(10).Value
    End If
End If
End If

End Sub

Private Sub Command2_Click()
Dim employeeId As Integer
employeeId = cboEmployeeId.Text
dtaEmployeeID.Recordset.MoveFirst
dtaEmployeeID.Recordset.Find " [EMPLOYEE ID]= " & employeeId, 0, adSearchForward
If dtaEmployeeID.Recordset.EOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
ElseIf dtaEmployeeID.Recordset.Fields(0).Value = employeeId Then
    dtaEmployeeID.Recordset.MoveNext
    If dtaEmployeeID.Recordset.EOF = False Then
        'dtaEmployeeID.Recordset.MoveNext
        cboEmployeeId.Text = dtaEmployeeID.Recordset.Fields(0).Value
        txtName.Text = dtaEmployeeID.Recordset.Fields(1).Value
        txtCurResidence.Text = dtaEmployeeID.Recordset.Fields(2).Value
        txtIDNumber.Text = dtaEmployeeID.Recordset.Fields(3).Value
        txtReferreContacts.Text = dtaEmployeeID.Recordset.Fields(4).Value
        txtKRAPIN.Text = dtaEmployeeID.Recordset.Fields(5).Value
        cboGender.Text = dtaEmployeeID.Recordset.Fields(6).Value
        txtHomeResidence.Text = dtaEmployeeID.Recordset.Fields(7).Value
        txtPhoneNumber.Text = dtaEmployeeID.Recordset.Fields(8).Value
        txtAddress.Text = dtaEmployeeID.Recordset.Fields(9).Value
        txtBANKNumber.Text = dtaEmployeeID.Recordset.Fields(10).Value
    Else
        dtaEmployeeID.Recordset.MoveFirst
        
         cboEmployeeId.Text = dtaEmployeeID.Recordset.Fields(0).Value
        txtName.Text = dtaEmployeeID.Recordset.Fields(1).Value
        txtCurResidence.Text = dtaEmployeeID.Recordset.Fields(2).Value
        txtIDNumber.Text = dtaEmployeeID.Recordset.Fields(3).Value
        txtReferreContacts.Text = dtaEmployeeID.Recordset.Fields(4).Value
        txtKRAPIN.Text = dtaEmployeeID.Recordset.Fields(5).Value
        cboGender.Text = dtaEmployeeID.Recordset.Fields(6).Value
        txtHomeResidence.Text = dtaEmployeeID.Recordset.Fields(7).Value
        txtPhoneNumber.Text = dtaEmployeeID.Recordset.Fields(8).Value
        txtAddress.Text = dtaEmployeeID.Recordset.Fields(9).Value
        txtBANKNumber.Text = dtaEmployeeID.Recordset.Fields(10).Value
    End If
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
 Call populate_combo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub

