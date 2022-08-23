VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditGroupInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EDIT GROUP INFO"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboGroups 
      Height          =   315
      Left            =   720
      TabIndex        =   18
      Top             =   1080
      Width           =   2655
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
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   10695
      Begin MSAdodcLib.Adodc dtaGroups 
         Height          =   375
         Left            =   600
         Top             =   4200
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
         Left            =   4080
         TabIndex        =   19
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7920
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   7920
         TabIndex        =   1
         Top             =   3240
         Width           =   2295
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
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
         TabIndex        =   15
         Top             =   480
         Width           =   2055
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
         TabIndex        =   14
         Top             =   1440
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
         TabIndex        =   13
         Top             =   2280
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
         TabIndex        =   12
         Top             =   2280
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
         TabIndex        =   11
         Top             =   3240
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
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GROUP LEADER"
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
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label lblGroupSelect 
      Alignment       =   2  'Center
      Caption         =   "SELECT GROUP TO EDIT INFO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   480
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      Height          =   1575
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmEditGroupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comboPopulate()
While dtaGroups.Recordset.EOF = False
    cboGroups.AddItem dtaGroups.Recordset.Fields(0).Value
    dtaGroups.Recordset.MoveNext
Wend
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmHomepage.Show
End Sub
