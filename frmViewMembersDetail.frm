VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmViewMembersDetail 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12690
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaMemberDet 
      Height          =   375
      Left            =   10320
      Top             =   7080
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
      RecordSource    =   "YOUTH_MEMBERS"
      Caption         =   "MEMBER DETAIL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10080
      TabIndex        =   27
      Top             =   1800
      Width           =   2535
      Begin VB.ComboBox cboMemberId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "SELECT MEMBER ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton PREVIOUS 
      Caption         =   "PREVIOUS"
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
      Left            =   6720
      TabIndex        =   25
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton NEXT 
      Caption         =   "NEXT"
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
      Left            =   8400
      TabIndex        =   24
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
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
      Left            =   4680
      TabIndex        =   23
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
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
      TabIndex        =   22
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton ADD_NEW 
      Caption         =   "ADD NEW"
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
      Index           =   0
      Left            =   840
      TabIndex        =   21
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "CHURCH DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5280
      TabIndex        =   1
      Top             =   1800
      Width           =   4695
      Begin VB.TextBox txtDOBaptism 
         DataSource      =   "dtaMemberDet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtDOC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtDistrictServed 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtDistrictElder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtCongregation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblDOC 
         Caption         =   "DATE OF COMMISSION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label lblDistrictServed 
         Caption         =   "DISTRICT SERVED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "DISTRICT ELDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "DATE OF BAPTISM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label CONGREGATION 
         Caption         =   "MEMBERS CONGREGATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERSONAL INFORMATION"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4815
      Begin VB.TextBox txtIDNmuber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox txtOccupation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtContacts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtResidence 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtMemberName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "ID NUMBER"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label OCCUPATION 
         Caption         =   "OCCUPATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label CONTACTS 
         Caption         =   "CONTACTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "RESIDENCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label NAME 
         Caption         =   "MEMBER NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frmViewMembersDetail.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "MEMBERS DETAIL PAGE"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   31
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "P.C.E.A KISAUNI MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   30
      Top             =   120
      Width           =   7695
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   9255
   End
End
Attribute VB_Name = "frmViewMembersDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub populate_combo()
While dtaMemberDet.Recordset.EOF = False
        cboMemberId.AddItem dtaMemberDet.Recordset.Fields(0).Value
        dtaMemberDet.Recordset.MoveNext
    Wend
End Sub

Private Sub ADD_NEW_Click(Index As Integer)
frmAddNew.Show
Unload Me
End Sub

Private Sub cboMemberId_Click()
Dim MEMBERID As Integer
MEMBERID = cboMemberId.Text
dtaMemberDet.Recordset.MoveFirst
dtaMemberDet.Recordset.Find " [MEMBER ID]= " & MEMBERID, 0, adSearchForward
If dtaMemberDet.Recordset.EOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
ElseIf dtaMemberDet.Recordset.Fields(0).Value = MEMBERID Then
 txtCongregation.Text = dtaMemberDet.Recordset.Fields(7).Value
 txtContacts.Text = dtaMemberDet.Recordset.Fields(6).Value
 txtDistrictElder.Text = dtaMemberDet.Recordset.Fields(10).Value
 txtDistrictServed.Text = dtaMemberDet.Recordset.Fields(9).Value
 txtDOBaptism.Text = dtaMemberDet.Recordset.Fields(8).Value
 txtDOC.Text = dtaMemberDet.Recordset.Fields(11).Value
 txtMemberName.Text = dtaMemberDet.Recordset.Fields(1).Value
 txtOccupation.Text = dtaMemberDet.Recordset.Fields(12).Value
 txtResidence.Text = dtaMemberDet.Recordset.Fields(2).Value
 txtIDNmuber.Text = dtaMemberDet.Recordset.Fields(3).Value
End If
End Sub

Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub Command2_Click()

Dim searchvalue As Long
    
    searchvalue = cboMemberId.Text
    dtaMemberDet.Recordset.MoveFirst
    dtaMemberDet.Recordset.Find "[MEMBER ID]= " & searchvalue, 0, adSearchForward
    If dtaMemberDet.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaMemberDet.Recordset.MoveFirst
    ElseIf dtaMemberDet.Recordset.Fields(0).Value = searchvalue Then
        If MsgBox("Are you sure you want to delete the member details?", vbOKCancel + vbQuestion) = vbOK Then
            dtaMemberDet.Recordset.Delete
            dtaMemberDet.Recordset.Update
            MsgBox "MEMBER DETAILS RECORDED SUCCESSFULLY", vbInformation
        End If
        cboMemberId.Clear
        txtCongregation = ""
        txtContacts.Text = ""
        txtDistrictElder.Text = ""
        txtDistrictServed.Text = ""
        txtDOBaptism.Text = ""
        txtDOC.Text = ""
        txtIDNmuber.Text = ""
        txtCongregation = ""
        txtMemberName.Text = ""
        txtOccupation.Text = ""
        txtResidence.Text = ""
        dtaMemberDet.Refresh
        Call populate_combo
    End If
        
   
End Sub

Private Sub Command3_Click()
frmUpdateMembersList.Show
frmViewMembersDetail.Hide
End Sub

Private Sub Form_Load()
Call populate_combo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub

Private Sub NEXT_Click()
Dim MEMBER_ID As Long
MEMBER_ID = cboMemberId.Text
dtaMemberDet.Recordset.MoveFirst
dtaMemberDet.Recordset.Find " [MEMBER ID]= " & MEMBER_ID, 0, adSearchForward
If dtaMemberDet.Recordset.EOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
ElseIf dtaMemberDet.Recordset.Fields(0).Value = MEMBER_ID Then
    dtaMemberDet.Recordset.MoveNext
    If dtaMemberDet.Recordset.EOF = False Then
       
        cboMemberId.Text = dtaMemberDet.Recordset.Fields(0).Value
        txtContacts.Text = dtaMemberDet.Recordset.Fields(6).Value
        txtDistrictElder.Text = dtaMemberDet.Recordset.Fields(10).Value
        txtDistrictServed.Text = dtaMemberDet.Recordset.Fields(9).Value
        txtDOC.Text = dtaMemberDet.Recordset.Fields(11).Value
        txtIDNmuber.Text = dtaMemberDet.Recordset.Fields(3).Value
        txtMemberName.Text = dtaMemberDet.Recordset.Fields(1).Value
        txtOccupation.Text = dtaMemberDet.Recordset.Fields(12).Value
        txtResidence.Text = dtaMemberDet.Recordset.Fields(2)
        
    Else
    dtaMemberDet.Recordset.MoveFirst
    

        cboMemberId.Text = dtaMemberDet.Recordset.Fields(0).Value
        txtContacts.Text = dtaMemberDet.Recordset.Fields(6).Value
        txtDistrictElder.Text = dtaMemberDet.Recordset.Fields(10).Value
        txtDistrictServed.Text = dtaMemberDet.Recordset.Fields(9).Value
        txtDOC.Text = dtaMemberDet.Recordset.Fields(11).Value
        txtIDNmuber.Text = dtaMemberDet.Recordset.Fields(3).Value
        txtMemberName.Text = dtaMemberDet.Recordset.Fields(1).Value
        txtOccupation.Text = dtaMemberDet.Recordset.Fields(12).Value
        txtResidence.Text = dtaMemberDet.Recordset.Fields(2)
        
    End If
End If
    
End Sub

Private Sub PREVIOUS_Click()
Dim MEMBER_ID As Integer
MEMBER_ID = cboMemberId.Text
dtaMemberDet.Recordset.MovePrevious
dtaMemberDet.Recordset.Find " [MEMBER ID]= " & MEMBER_ID, 0, adSearchForward
If dtaMemberDet.Recordset.EOF = True Then
    'dtaEmployeeID.Recordset.MoveFirst
ElseIf dtaMemberDet.Recordset.Fields(0).Value = MEMBER_ID Then
    dtaMemberDet.Recordset.MovePrevious
    If dtaMemberDet.Recordset.EOF = False Then
       
        cboMemberId.Text = dtaMemberDet.Recordset.Fields(0).Value
        txtContacts.Text = dtaMemberDet.Recordset.Fields(6).Value
        txtDistrictElder.Text = dtaMemberDet.Recordset.Fields(10).Value
        txtDistrictServed.Text = dtaMemberDet.Recordset.Fields(9).Value
        txtDOC.Text = dtaMemberDet.Recordset.Fields(11).Value
        txtIDNmuber.Text = dtaMemberDet.Recordset.Fields(3).Value
        txtMemberName.Text = dtaMemberDet.Recordset.Fields(1).Value
        txtOccupation.Text = dtaMemberDet.Recordset.Fields(12).Value
        txtResidence.Text = dtaMemberDet.Recordset.Fields(2)
        
    Else
    dtaMemberDet.Recordset.MoveLast
        cboMemberId.Text = dtaMemberDet.Recordset.Fields(0).Value
        txtContacts.Text = dtaMemberDet.Recordset.Fields(6).Value
        txtDistrictElder.Text = dtaMemberDet.Recordset.Fields(10).Value
        txtDistrictServed.Text = dtaMemberDet.Recordset.Fields(9).Value
        txtDOC.Text = dtaMemberDet.Recordset.Fields(11).Value
        txtIDNmuber.Text = dtaMemberDet.Recordset.Fields(3).Value
        txtMemberName.Text = dtaMemberDet.Recordset.Fields(1).Value
        txtOccupation.Text = dtaMemberDet.Recordset.Fields(12).Value
        txtResidence.Text = dtaMemberDet.Recordset.Fields(2)
        
    End If
End If
End Sub
