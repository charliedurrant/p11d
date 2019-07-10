VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Object = "{A7CE771F-05B2-43CF-9650-ED841A9049FA}#1.0#0"; "atc3FolderBrowser.ocx"
Begin VB.Form F_PayeOnline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PayeOnline"
   ClientHeight    =   5190
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin atc3FolderBrowser.FolderBrowser fb 
      Height          =   510
      Left            =   45
      TabIndex        =   13
      Top             =   4230
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   900
   End
   Begin VB.Frame frmPAYEOnlineButtons 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox pctFrame 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   90
         ScaleHeight     =   375
         ScaleWidth      =   5010
         TabIndex        =   10
         Top             =   3825
         Width           =   5010
         Begin VB.OptionButton optOnlineForm 
            Caption         =   "P11D + P11D(b)"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optOnlineForm 
            Caption         =   "P46(Car)"
            Height          =   375
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Height          =   375
         Left            =   7995
         TabIndex        =   8
         Top             =   120
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   7995
         TabIndex        =   7
         Top             =   600
         Width           =   1140
      End
      Begin VB.CommandButton cmdViewErrors 
         Caption         =   "V&iew Errors"
         Height          =   375
         Left            =   7995
         TabIndex        =   6
         Top             =   2520
         Width           =   1140
      End
      Begin VB.CommandButton cmdPrintErrors 
         Caption         =   "&Print Errors"
         Height          =   375
         Left            =   7995
         TabIndex        =   5
         Top             =   3000
         Width           =   1140
      End
      Begin VB.CommandButton cmdViewLastFile 
         Caption         =   "&View Last File"
         Height          =   495
         Left            =   7995
         TabIndex        =   4
         Top             =   1200
         Width           =   1140
      End
      Begin VB.CommandButton cmdViewLastResponse 
         Caption         =   "View Last &Response"
         Height          =   495
         Left            =   7995
         TabIndex        =   3
         Top             =   1800
         Width           =   1140
      End
      Begin MSComctlLib.ListView lvPAYEEmployers 
         Height          =   3780
         Left            =   45
         TabIndex        =   9
         Top             =   0
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   6668
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Employer name"
            Object.Width           =   2222
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PAYE ref"
            Object.Width           =   1453
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valid PAYE"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "No. of employees"
            Object.Width           =   2485
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Submitter ID"
            Object.Width           =   1854
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Password"
            Object.Width           =   1552
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Username"
            Object.Width           =   1589
         EndProperty
      End
   End
   Begin TCSPROG.TCSProgressBar prgEmployee 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   4770
      Visible         =   0   'False
      Width           =   9210
      _cx             =   16245
      _cy             =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   50
      BarBackColor    =   12632256
      BarForeColor    =   8388608
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   1
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   0   'False
      Increment       =   1
      TextAlignment   =   0
   End
   Begin VB.Label lblWritingFile 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   3780
      Width           =   5805
   End
   Begin VB.Menu mnuExtraSubmissionProperties 
      Caption         =   "&Debug"
      Begin VB.Menu mnuTestSubmission 
         Caption         =   "Test submission (submits to IR but tells IR it is a test)"
      End
      Begin VB.Menu mnuProceedSubmission 
         Caption         =   "Proceed with submission (submit to IR or if mock then create db entry therefore 'Test submission' is irrelevant)"
      End
      Begin VB.Menu mnuMockSubmission 
         Caption         =   "Mock submission (true don't submit to gateway)"
      End
   End
End
Attribute VB_Name = "F_PayeOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFrmGeneral
Implements IViewFile

Private m_InvalidVT As Control
Public OK As Boolean
Public loaded As Boolean

Private Function SetButtons(bVisible As Boolean) As Boolean
  cmdPrintErrors.Visible = bVisible
  cmdViewErrors.Visible = bVisible
  SetButtons = bVisible
End Function

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdViewOptions_Click()
'   F_PayeOnlineP46Options.Show vbModal
  Call p11d32.Help.ShowForm(F_PayeOnlineP46Options, vbModal)
End Sub

Private Sub cmdPrintErrors_Click()
  'View validation errors
  Call p11d32.PAYEonline.Errors(PRINT_REPORT, VET_PAYEONLINE_VALIDATION)
End Sub

Private Sub PAYERun()
  On Error GoTo PAYERun_ERR
  Dim li As ListItem
  Dim i As Long
  Dim benEmployer As IBenefitClass
  
  Call xSet("PAYERun")
  Call SetCursor
  
  
  If Not FileExists(fb.Directory, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "ValidateMMData", "The directory " & fb.Directory & " does not exist.")
  
  'Validate Employer level required fields
  If ValidatePAYEOnlineEmployerData Then
    'Create files
    
    Call p11d32.PAYEonline.CreatePAYEFiles(prgEmployee, lblWritingFile)
  End If
  
PAYERun_END:
  Call ClearCursor
  Call xReturn("PAYERun")
  Exit Sub
PAYERun_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PAYERun", "MM Run", "Error running the PAYE Online.")
  Resume PAYERun_END
  Resume
End Sub

Private Sub cmdRun_Click()
  On Error GoTo err_Err
  Call EnableFrame(Me, frmPAYEOnlineButtons, False)
    If p11d32.PAYEonline.Type_Selected = POT_P46 Then
      Call p11d32.Help.ShowForm(F_PayeOnlineP46Options, vbModal)
      Call PAYERun
    ElseIf p11d32.PAYEonline.Type_Selected = POT_P11D Then
      Call PAYERun
    End If
  
err_End:
  Call EnableFrame(Me, frmPAYEOnlineButtons, True)
  Call LicenceDependantSettings
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Run", "Run", Err.Description)
  Resume err_End
  Resume
End Sub
Private Sub cmdViewErrors_Click()
  Call p11d32.PAYEonline.Errors(PREPARE_REPORT, VET_PAYEONLINE_VALIDATION)
End Sub

Private Sub cmdViewLastFile_Click()
  On Error GoTo err_Err
   
    
  
  Dim myIE As SHDocVw.InternetExplorer
  If FileExists(p11d32.PAYEonline.LastPathAndFileCreated) Then
    Set myIE = New SHDocVw.InternetExplorer
    Call myIE.Navigate(p11d32.PAYEonline.LastPathAndFileCreated)
    myIE.Visible = True
  Else
    Call F_ViewFile.ViewFile("", , p11d32.PAYEonline.OutputDirectory, Me)
  End If
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ViewLastFile", "ViewLastFile", "Error viewing the last file:'" & p11d32.PAYEonline.LastPathAndFileCreated & "'")
  Resume err_End
End Sub

Private Sub cmdViewLastResponse_Click()
  Dim sLastPathAndFileCreated
  
  On Error GoTo err_Err
  
  sLastPathAndFileCreated = Replace(p11d32.PAYEonline.LastPathAndFileCreated, ".xml", ".txt")
  If Not FileExists(sLastPathAndFileCreated) Then
    sLastPathAndFileCreated = ""
  End If
  
  
  Call F_ViewFile.ViewFile(sLastPathAndFileCreated, , p11d32.PAYEonline.OutputDirectory, Me)
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ViewLastRespose", "ViewLastResponse", "Error viewing the last response")
  Resume err_End
End Sub

Private Sub LicenceDependantSettings()
  'submit P11D
  optOnlineForm(1).Enabled = p11d32.ReportPrint.AllReports
  If Not p11d32.ReportPrint.AllReports Then
    p11d32.PAYEonline.Type_Selected = POT_P46
  End If
  Call optOnlineForm_Click(-1)
  
End Sub
Private Sub Form_Load()
  
  Call SetErrorButtons(True)
  Set lvPAYEEmployers.SmallIcons = MDIMain.imlTree
  If Not FileExists(p11d32.PAYEonline.OutputDirectory, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "F_MM", "The PAYE Online directory does not exist = " & p11d32.PAYEonline.OutputDirectory & ", unable to create export.")
  fb.Directory = p11d32.PAYEonline.OutputDirectory
  cmdRun.Enabled = False 'Disable button until optOnlineForm (POT) selected
  
  'Disable P11D, P11Db options if "short" licence
  
  
  
  Call LicenceDependantSettings
  
  mnuMockSubmission.Checked = p11d32.PAYEonline.Efiler_Mock_Submission
  mnuProceedSubmission.Checked = p11d32.PAYEonline.Efiler_Proceed_Submission
  mnuTestSubmission.Checked = p11d32.PAYEonline.Efiler_Test_Submission
  mnuExtraSubmissionProperties.Visible = p11d32.PAYEonline.ExtraSubmissionPropertiesMenu Or IsRunningInIDE
End Sub
Private Function SetErrorButtons(bVisible As Boolean) As Boolean
  cmdPrintErrors.Visible = bVisible
  cmdViewErrors.Visible = bVisible
  SetErrorButtons = bVisible
End Function

Private Sub Form_Unload(Cancel As Integer)
  If Not p11d32 Is Nothing Then p11d32.MagneticMedia.UserDataSize = 0
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

Private Sub IViewFile_View(rt As RichTextLib.IRichText, sLabelInfo As String)
 On Error GoTo IViewFile_MagneticMedia_ERR
  
  Call xSet("IViewFile_MagneticMedia")
  
  sLabelInfo = HighLightRecord(rt, p11d32.MagneticMedia.RecordType(MM_REC_EMPLOYEE)) & " employees "
  If Len(p11d32.MagneticMedia.RecordViewID) Then
    sLabelInfo = sLabelInfo & ", " & HighLightRecord(rt, p11d32.MagneticMedia.RecordViewID, vbRed) & " type " & p11d32.MagneticMedia.RecordViewID & " records."
  End If
  
IViewFile_MagneticMedia_END:
  Call xReturn("IViewFile_MagneticMedia")
  Exit Sub
IViewFile_MagneticMedia_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IViewFile", "IView File Magnetic Media", "Error setting the text properties when loading a file form the magnetic media dialogue.")
  Resume IViewFile_MagneticMedia_END
End Sub
Private Sub lvPAYEEmployers_BeforeLabelEdit(Cancel As Integer)
  Cancel = True
End Sub

Public Sub lvPAYEEmployers_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  cmdRun.Enabled = ListViewAnyChecked(lvPAYEEmployers)
End Sub

Private Sub lvPAYEEmployers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call EditEmployerDetails(Button, Shift, X, Y)
End Sub

Private Sub EditEmployerDetails(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim ibf As IBenefitForm2
  Dim li As ListItem, liMe As ListItem, liPrevious As ListItem
  
  On Error GoTo EditEmployerDetails_ERR
   
  Call xSet("EditEmployerDetails")
  
  If Not ((Button And vbRightButton) = vbRightButton) Then GoTo EditEmployerDetails_END
    
  Set ibf = CurrentForm
    
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, Err, "The benefit for is nothing") 'JN
  If Not ibf.benclass = BC_EMPLOYER Then Call Err.Raise(ERR_IS_NOTHING, Err, "Edit Employer Details") 'JN
  
  Set liMe = lvPAYEEmployers.HitTest(X, Y)
  
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


Private Sub mnuMockSubmission_Click()
  mnuMockSubmission.Checked = Not mnuMockSubmission.Checked
  
  p11d32.PAYEonline.Efiler_Mock_Submission = mnuMockSubmission.Checked
End Sub

Private Sub mnuProceedSubmission_Click()
  mnuProceedSubmission.Checked = Not mnuProceedSubmission.Checked
  p11d32.PAYEonline.Efiler_Proceed_Submission = mnuProceedSubmission.Checked
End Sub

Private Sub mnuTestSubmission_Click()
  mnuTestSubmission.Checked = Not mnuTestSubmission.Checked
  p11d32.PAYEonline.Efiler_Test_Submission = mnuTestSubmission.Checked
End Sub

Private Sub optOnlineForm_Click(Index As Integer)
    'Set POT type
    Dim i As Integer
    Select Case Index
      Case -1
        optOnlineForm(p11d32.PAYEonline.Type_Selected).value = True
        Exit Sub
      Case 0
        p11d32.PAYEonline.Type_Selected = POT_P46
      Case 1
        p11d32.PAYEonline.Type_Selected = POT_P11D
    End Select
        
    'Enable/disable run button
    
    If optOnlineForm(Index).value Then
      cmdRun.Enabled = True
    Else
      cmdRun.Enabled = False
    End If
    
End Sub


'Public Function SettingsToScreen() As Boolean
'
'  Dim i As Long
'
'  On Error GoTo SettingsToScreen_Err
'  Call xSet("SettingsToScreen")
'
'  Call lvPAYEEmployers_ItemCheck(Nothing)
'
'SettingsToScreen_End:
'  Call xReturn("SettingsToScreen")
'  Exit Function
'
'SettingsToScreen_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error setting dates for P46.")
'  Resume SettingsToScreen_End
'  Resume
'End Function

Private Function CheckStatus() As Boolean
  Dim ey As Employer
  Dim benEmployer As IBenefitClass
  Dim i As Long, j As Long
  Dim li As ListItem

  Call xSet("CheckStatus")

  On Error GoTo CheckStatus_ERR

'  If InvalidFields(Me) Then Call Err.Raise(ERR_INVALID_FIELDS, "CheckStatus", "Some of the data entry fields are invalid, please amend.")
'
  For i = 1 To lvPAYEEmployers.listitems.Count
    Set li = lvPAYEEmployers.listitems(i)
    If Not li.Checked Then GoTo NEXT_ITEM
    Set ey = p11d32.Employers(li.Tag)

    Set benEmployer = ey

    If Not benEmployer.value(employer_PAYEOnlineSelected) Then GoTo NEXT_ITEM
    
    'Call p11d32.PAYEonline.CheckSubmissionStatus(benEmployer)
    

NEXT_ITEM:
  Next

  CheckStatus = True
  
CheckStatus_END:
  Call xSet("CheckStatus")
  Exit Function
CheckStatus_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CheckStatus", "Check PayeOnline Status", "Error checking status for PayeOnline Submissions.")
  Resume CheckStatus_END
  Resume
End Function

Private Function ValidatePAYEOnlineEmployerData() As Boolean
  Dim ey As Employer
  Dim benEmployer As IBenefitClass, benEmployer2 As IBenefitClass
  Dim i As Long, j As Long
  Dim li As ListItem
  Dim bAnyChecked As Boolean
  
  Call xSet("ValidatePAYEOnlineEmployerData")
  
  On Error GoTo ValidatePAYEOnlineEmployerData_ERR
  
  If InvalidFields(Me) Then Call Err.Raise(ERR_INVALID_FIELDS, "ValidatePAYEOnlineEmployerData", "Some of the data entry fields are invalid, please amend.")
  
  For i = 1 To lvPAYEEmployers.listitems.Count
    Set li = lvPAYEEmployers.listitems(i)
    
    'Update Employer objects with user selection
     Set benEmployer = p11d32.Employers(li.Tag)
     benEmployer.value(employer_PAYEOnlineSelected) = li.Checked
    
    If li.Checked Then
      bAnyChecked = True
      Set ey = p11d32.Employers(li.Tag)
    
     'Check Employer level PAYE fields
      Call ey.PAYEOnlineValid(True)
    End If
  Next i
  
  If (Not bAnyChecked) Then Call Err.Raise(ERR_INVALID, "ValidatePAYEOnlineEmployerData", "No employers selected")
  ValidatePAYEOnlineEmployerData = True
  
ValidatePAYEOnlineEmployerData_END:
  Call xSet("ValidatePAYEOnlineEmployerData")
  Exit Function
ValidatePAYEOnlineEmployerData_ERR:
  ValidatePAYEOnlineEmployerData = False
  Call ErrorMessage(ERR_ERROR, Err, "ValidatePAYEOnlineEmployerData", "Validate MM Data", "Error validating the data for Magnetic Media.")
  Resume ValidatePAYEOnlineEmployerData_END
  Resume
End Function

Private Sub vtOutPutDir_Change()

End Sub
Private Sub fb_Ended()
  p11d32.PAYEonline.OutputDirectory = fb.Directory
End Sub

Private Sub fb_Started()
  fb.Directory = p11d32.PAYEonline.OutputDirectory
End Sub

