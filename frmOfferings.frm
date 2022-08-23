VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOfferings 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7395
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "OFFERINGS RECORD"
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin MSAdodcLib.Adodc dtaOfferings 
         Height          =   375
         Left            =   840
         Top             =   6720
         Width           =   2175
         _ExtentX        =   3836
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
         RecordSource    =   "OFFERINGS"
         Caption         =   "OFFERINGS"
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
      Begin VB.CommandButton Command1 
         Caption         =   "UPLOAD"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   16
         Top             =   6720
         Width           =   2055
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   15
         Top             =   5880
         Width           =   2055
      End
      Begin VB.TextBox txtChurchSchool 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7320
         TabIndex        =   13
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox txtThanksGiving 
         DataSource      =   "dtaOfferings"
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
         Left            =   7320
         TabIndex        =   11
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox txtYouthProject 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtCashPayment 
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
         Left            =   2640
         TabIndex        =   7
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox txtMpesaPayment 
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
         Left            =   2640
         TabIndex        =   5
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtBankPayment 
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
         Left            =   2640
         TabIndex        =   3
         Top             =   2760
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50069505
         CurrentDate     =   44425
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "OFFERING RECORD"
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
         TabIndex        =   18
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C000C0&
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
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "GRAND TOTAL"
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
         Left            =   2760
         TabIndex        =   14
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "CHURCH SCHOOL"
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
         Left            =   5400
         TabIndex        =   12
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "THANKS GIVING"
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
         Left            =   5280
         TabIndex        =   10
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "YOUTH PROJECT"
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
         TabIndex        =   8
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "CASH PAYMENT"
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
         Left            =   240
         TabIndex        =   6
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label lblMpesaPayment 
         Alignment       =   2  'Center
         Caption         =   "MPESA PAYMENT"
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
         TabIndex        =   4
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "BANK PAYMENT"
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
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   7440
         Picture         =   "frmOfferings.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOfferings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 If txtBankPayment.Text = "" Or txtCashPayment.Text = "" Or txtChurchSchool.Text = "" Or txtMpesaPayment.Text = "" Or txtYouthProject.Text = "" Then
 MsgBox "KINDLY KEY IN ALL COLLECTIONS", vbInformation
 Else
 Dim curTotalContribution As Currency
 Dim curChurchSchool As Currency
 Dim curBankPayment As Currency
 Dim curCashPayment As Currency
 Dim curThanksGiving As Currency
 Dim curYouthProject As Currency
 Dim curMpesaPayment As Currency
     
    curBankPayment = txtBankPayment.Text
    curCashPayment = txtCashPayment.Text
    curChurchSchool = txtChurchSchool.Text
    curMpesaPayment = txtMpesaPayment.Text
    curYouthProject = txtYouthProject.Text
    curThanksGiving = txtThanksGiving.Text
     curTotalContribution = (curBankPayment + curCashPayment + curChurchSchool + curMpesaPayment + curYouthProject + curThanksGiving)
 dtaOfferings.Recordset.AddNew
 dtaOfferings.Recordset.Fields(0).Value = curBankPayment
 dtaOfferings.Recordset.Fields(1).Value = curMpesaPayment
 dtaOfferings.Recordset.Fields(2).Value = curCashPayment
 dtaOfferings.Recordset.Fields(3).Value = curYouthProject
 dtaOfferings.Recordset.Fields(4).Value = curThanksGiving
 dtaOfferings.Recordset.Fields(5).Value = DTPicker1.Value
 dtaOfferings.Recordset.Fields(6).Value = curChurchSchool
 dtaOfferings.Recordset.Fields(7).Value = curTotalContribution
 dtaOfferings.Recordset.Update
 MsgBox "OFFERINGS SUCCESFULLY UPLOADED", vbInformation
 txtBankPayment.Text = ""
 txtCashPayment.Text = ""
 txtChurchSchool.Text = ""
 txtMpesaPayment.Text = ""
 txtThanksGiving.Text = ""
 txtYouthProject.Text = ""
 txtTotal.Text = ""
 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub

Private Sub txtChurchSchool_LostFocus()
Dim curMpesaPayment As Currency
 Dim curTotalContribution As Currency
 Dim curChurchSchool As Currency
 Dim curBankPayment As Currency
 Dim curCashPayment As Currency
 Dim curThanksGiving As Currency
 Dim curYouthProject As Currency
    curBankPayment = txtBankPayment.Text
    curCashPayment = txtCashPayment.Text
    curChurchSchool = txtChurchSchool.Text
    curMpesaPayment = txtMpesaPayment.Text
    curYouthProject = txtYouthProject.Text
    curThanksGiving = txtThanksGiving.Text
 curTotalContribution = (curBankPayment + curCashPayment + curChurchSchool + curMpesaPayment + curYouthProject + curThanksGiving)
 txtTotal.Text = curTotalContribution
End Sub
