VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_PayeOnlineStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Status"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "View Errors"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmd 
      Caption         =   "View File"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   3
      Top             =   3480
      Width           =   1740
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel Submission"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Resume Submission"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSComctlLib.ListView lvSubmissions 
      Height          =   3300
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   5821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2222
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employer"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Message"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Last Updated"
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "F_PayeOnlineStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is used to display information extracted from the e-filer submission database
Implements IFrmGeneral

Private m_InvalidVT As Control
Public OK As Boolean
Public loaded As Boolean


Private Sub cmd_Click(Index As Integer)

  Dim i As Integer
  Dim li As ListItem
  Dim SubIDs As String
  
  Select Case Index
    Case 0 ' Cancel - Exit
        Unload Me
    Case 1 ' Cancel - Submission
        For i = 1 To F_PayeOnlineStatus.lvSubmissions.listitems.Count
            Set li = F_PayeOnlineStatus.lvSubmissions.listitems(i)
            If li.Checked Then Call p11d32.PAYEonline.CancelSubmission(li.Tag, li.Key)
        Next
        'Refresh status by reloading form
        Call GetStatus(False)
    Case 2 'Resume Submission
        For i = 1 To F_PayeOnlineStatus.lvSubmissions.listitems.Count
            Set li = F_PayeOnlineStatus.lvSubmissions.listitems(i)
            If li.Checked And Len(li.Key) Then p11d32.PAYEonline.eFiler.ResumeSubmission (li.Key)
        Next
        'Refresh status by reloading form
        Call GetStatus(False)
    Case 4 'View Errors
        SubIDs = "("
        For i = 1 To F_PayeOnlineStatus.lvSubmissions.listitems.Count
            Set li = F_PayeOnlineStatus.lvSubmissions.listitems(i)
            If li.Checked Then SubIDs = SubIDs & "'" & li.Key & "', "
        Next
        SubIDs = SubIDs & "'x')"
        Call p11d32.PAYEonline.Errors(PREPARE_REPORT, VET_PAYEONLINE_SUBMISSION, SubIDs)
        

  End Select
End Sub

Private Sub Form_Load()
  'Call SetButtons(False)
  Set lvSubmissions.SmallIcons = MDIMain.imlTree
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 1 To 4
        cmd(i).Visible = False
        cmd(i).Enabled = False
    Next
End Sub

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  
End Property
Private Function HighLightRecord(rt As RichTextBox, ByVal sRecordID As String, Optional lCol As Long = vbBlue) As Long
  Dim l As Long, m As Long
  Dim lLens As String, lLenlf As Long

  On Error GoTo HighLightRecord_ERR
  
  Call xSet("HighLightRecord")
  

  If Len(sRecordID) = 0 Then GoTo HighLightRecord_END
  sRecordID = vbLf & sRecordID
  lLens = Len(sRecordID)
  lLenlf = Len(vbLf)
  
  l = rt.Find(sRecordID, 0)
  Do While l <> -1
    m = rt.Find(vbLf, l + lLens)
    If m <> -1 Then
      HighLightRecord = HighLightRecord + Abs(RTSelText(rt, l + lLenlf, (m - 1) - l, lCol))
    End If
    l = l + lLens
    l = rt.Find(sRecordID, l)
  Loop
  
HighLightRecord_END:
  Call xSet("HighLightRecord")
  Exit Function
HighLightRecord_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HighLightRecord", "High Light Record", "Error highlighting a record with ID " & sRecordID & ".")
  Resume HighLightRecord_END
  
End Function

Private Sub lvSubmissions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call EditEmployerDetails(Button, Shift, x, y)
End Sub

Private Sub EditEmployerDetails(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ibf As IBenefitForm2
  Dim li As ListItem, liMe As ListItem, liPrevious As ListItem
  
  On Error GoTo EditEmployerDetails_ERR
   
  Call xSet("EditEmployerDetails")
  
  If Not ((Button And vbRightButton) = vbRightButton) Then GoTo EditEmployerDetails_END
    
  Set ibf = CurrentForm
    
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, Err, "The benfit for is nothing") 'JN
  If Not ibf.benclass = BC_EMPLOYER Then Call Err.Raise(ERR_IS_NOTHING, Err, "Edit Employer Details") 'JN
  
  Set liMe = lvSubmissions.HitTest(x, y)
  
  If liMe Is Nothing Then GoTo EditEmployerDetails_END
  
  For Each li In ibf.lv.listitems
    If li.Tag = liMe.Tag Then
      Set liPrevious = ibf.lv.SelectedItem
      Set ibf.lv.SelectedItem = li
      Call p11d32.EditEmployer(li.Tag)
      Call p11d32.PAYEonline.UpdateListViewItem(liMe, p11d32.Employers(li.Tag))
      Set ibf.lv.SelectedItem = liPrevious
      Exit For
    End If
  Next
  
EditEmployerDetails_END: 'JN
  Call xReturn("EditEmployerDetails")
  Exit Sub
EditEmployerDetails_ERR: 'JN
  Call ErrorMessage(ERR_ERROR, Err, "EditEmployerDetails", "Edit Employer Details", "Error editing Employer Details from the magnetic media screen.")
  Resume EditEmployerDetails_END
  Resume
End Sub

Private Sub GetStatus(AllEmployers As Boolean)
    
  Dim ey As Employer
  Dim ben As IBenefitClass
  Dim i As Long
  Dim li As ListItem

  Call xSet("GetStatus")

  On Error GoTo GetStatus_ERR


  'Emptying list view - need to do this when we refresh
  For i = 1 To F_PayeOnlineStatus.lvSubmissions.listitems.Count
        F_PayeOnlineStatus.lvSubmissions.listitems.Remove (F_PayeOnlineStatus.lvSubmissions.listitems.Count)
  Next

  Call p11d32.PAYEonline.StatusToListView(AllEmployers)

'  If InvalidFields(Me) Then Call Err.Raise(ERR_INVALID_FIELDS, "GetStatus", "Some of the data entry fields are invalid, please amend.")
'
'
'  Select Case AllEmployers
'
'    Case False
'        For i = 1 To F_PayeOnline.lvMMEmployers.listitems.Count
'          Set li = F_PayeOnline.lvMMEmployers.listitems(i)
'          If li.Checked Then
'            Set ey = p11d32.Employers(li.Tag)
'            Call p11d32.PAYEOnline.StatusToListView(ey)
'          End If
'        Next
'    Case True
'        For i = 1 To p11d32.Employers.Count
'            Set ey = p11d32.Employers(i)
'            Call p11d32.LoadEmployer(ey, False)
'            Call p11d32.PAYEOnline.StatusToListView(ey)
'         Next
'
'    End Select
  
GetStatus_END:
  Call xSet("GetStatus")
  Exit Sub
GetStatus_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetStatus", "Check PayeOnline Status", "Error checking status for PayeOnline Submissions.")
  Resume GetStatus_END
  Resume
End Sub

Public Sub LoadForm(Index As Long)
    Select Case Index
        Case PO_Status
            cmd(1).Visible = True
            cmd(1).Enabled = True
            cmd(2).Visible = True
            cmd(2).Enabled = True
            Call GetStatus(False)
            
        Case PO_Errors
            cmd(4).Visible = True
            cmd(4).Enabled = True
            cmd(4).Top = cmd(2).Top
            cmd(4).Left = cmd(2).Left
            Call GetStatus(True)
            
    End Select

   Me.Show 1
End Sub



