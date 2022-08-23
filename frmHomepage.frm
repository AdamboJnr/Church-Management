VERSION 5.00
Begin VB.Form frmHomepage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KISAUNI CHURCH MANAGEMENT SYSTEM"
   ClientHeight    =   6075
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   0
      Picture         =   "frmHomepage.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "&USERS"
      Begin VB.Menu mnuCreateUser 
         Caption         =   "&Create User"
      End
      Begin VB.Menu mnuSystemUsers 
         Caption         =   "&View system users"
      End
   End
   Begin VB.Menu mnuMEMBERS 
      Caption         =   "&MEMBERS"
      Begin VB.Menu mnuAddNewmember 
         Caption         =   "&Add new members"
      End
      Begin VB.Menu mnuViewMembers 
         Caption         =   "&view members"
      End
   End
   Begin VB.Menu mnuScheduleServices 
      Caption         =   "&SCHEDULE SERVICES"
      Begin VB.Menu mnuSundayService 
         Caption         =   "&Sunday Service Schedule"
      End
      Begin VB.Menu mnuWednesdayschedule 
         Caption         =   "&Wednesday service schedule"
      End
      Begin VB.Menu mnuTuesdayService 
         Caption         =   "&Tuesday Schedule"
      End
   End
   Begin VB.Menu mnuFinancials 
      Caption         =   "&FINANCIALS"
      Begin VB.Menu mnuofferings 
         Caption         =   "&Offerings collection"
      End
      Begin VB.Menu mnuTithes 
         Caption         =   "&Tithes"
      End
   End
   Begin VB.Menu mnuStaff 
      Caption         =   "&STAFF"
      Begin VB.Menu mnuEmployee 
         Caption         =   "&Add Employee Record"
      End
      Begin VB.Menu mnuEmployeerecord 
         Caption         =   "&View Employees Records"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&REPORTS"
      Begin VB.Menu mnuofferingsReport 
         Caption         =   "&Generate Offerings Report"
      End
      Begin VB.Menu mnuTitheReport 
         Caption         =   "&Generate Tithe Report"
      End
      Begin VB.Menu mnuTuesdayReport 
         Caption         =   "&Generate Tuesday Service Schedule"
      End
      Begin VB.Menu mnuWednesdayReport 
         Caption         =   "&Generate Wednesday Service Schedule"
      End
      Begin VB.Menu mnuSundayReport 
         Caption         =   "&Generate Sunday Service Schedule"
      End
   End
End
Attribute VB_Name = "frmHomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmLogin.Show
End Sub

Private Sub mnuAddNewmember_Click()
frmAddNew.Show
frmHomepage.Hide

End Sub

Private Sub mnuCreateUser_Click()
frmRegister.Show
frmHomepage.Hide
End Sub

Private Sub mnuEmployee_Click()
frmEmployeeDetails.Show
End Sub

Private Sub mnuEmployeerecord_Click()
details.Show
End Sub

Private Sub mnuofferings_Click()
frmOfferings.Show

End Sub

Private Sub mnuofferingsReport_Click()
rptOfferings.Show
End Sub

Private Sub mnuSundayReport_Click()
rptSundaySchedule.Show
End Sub

Private Sub mnuSundayService_Click()
frmSundayFellowship.Show
End Sub

Private Sub mnuSystemUsers_Click()
If typeOfUser = "ADMIN" Then
frmSystemUsers.Show
frmHomepage.Hide
Else
MsgBox "UNAUTHORIZED ACCESS", vbCritical
End If

End Sub

Private Sub mnuThanksgivings_Click()

End Sub

Private Sub mnuTitheReport_Click()
rptTithes.Show
End Sub

Private Sub mnuTithes_Click()
frmTithes.Show
End Sub

Private Sub mnuTuesdayReport_Click()
rptTuesdayFellowship.Show
End Sub

Private Sub mnuTuesdayService_Click()
frmTuesdayService.Show
End Sub

Private Sub mnuViewMembers_Click()
frmViewMembersDetail.Show

End Sub
