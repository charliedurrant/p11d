VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Object = "{A7CE771F-05B2-43CF-9650-ED841A9049FA}#1.0#0"; "ATC3FolderBrowser.ocx"
Begin VB.Form F_Intranet 
   Caption         =   "FullPath(OutputDirectory)"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5550
      Top             =   5325
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctBannerBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   4050
      ScaleHeight     =   360
      ScaleWidth      =   660
      TabIndex        =   25
      Top             =   5325
      Width           =   690
   End
   Begin VB.PictureBox pctBannerForeColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1200
      ScaleHeight     =   360
      ScaleWidth      =   660
      TabIndex        =   23
      Top             =   5325
      Width           =   690
   End
   Begin VB.TextBox txtBannerTitle 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1200
      TabIndex        =   15
      Text            =   "txtBannerTitle"
      Top             =   4875
      Width           =   4140
   End
   Begin VB.TextBox txtUserInfoHTML 
      Appearance      =   0  'Flat
      Height          =   690
      Left            =   1200
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Text            =   "txtUserInfoHTML"
      Top             =   5925
      Width           =   7440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output Type"
      Height          =   1650
      Left            =   6390
      TabIndex        =   10
      Top             =   4140
      Width           =   2310
      Begin VB.OptionButton optP11DType 
         Caption         =   "Working Papers"
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1215
         Width           =   1575
      End
      Begin VB.OptionButton optP11DType 
         Caption         =   "Employee Letter"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   855
         Width           =   1575
      End
      Begin VB.OptionButton optP11DType 
         Caption         =   "P11D + Working Papers"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   2070
      End
      Begin VB.OptionButton optP11DType 
         Caption         =   "P11D only"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame frmAuthenticationType 
      Caption         =   "Authentication Type"
      Height          =   1515
      Left            =   6360
      TabIndex        =   6
      Top             =   2565
      Width           =   2310
      Begin VB.CheckBox chkCaseSensitive 
         Caption         =   "Case sensitive"
         Height          =   240
         Left            =   600
         TabIndex        =   22
         Top             =   525
         Width           =   1590
      End
      Begin VB.OptionButton optAuthType 
         Caption         =   "Other"
         Height          =   375
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   1050
         Width           =   1455
      End
      Begin VB.OptionButton optAuthType 
         Caption         =   "Windows authentication"
         Height          =   375
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   750
         Width           =   2055
      End
      Begin VB.OptionButton optAuthType 
         Caption         =   "Full authentication"
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   225
         Value           =   -1  'True
         Width           =   1830
      End
   End
   Begin atc3FolderBrowser.FolderBrowser fb 
      Height          =   555
      Left            =   2175
      TabIndex        =   5
      Top             =   6750
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   979
   End
   Begin VB.Frame fraUsername 
      Caption         =   "User login name source"
      ClipControls    =   0   'False
      Height          =   1710
      Left            =   6345
      TabIndex        =   4
      Top             =   825
      Width           =   2310
      Begin VB.OptionButton optUsername 
         Caption         =   "Full name"
         Height          =   375
         Index           =   3
         Left            =   150
         TabIndex        =   21
         Top             =   1305
         Width           =   1455
      End
      Begin VB.OptionButton optUsername 
         Caption         =   "Email address"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   975
         Width           =   1455
      End
      Begin VB.OptionButton optUsername 
         Caption         =   "Personnel number"
         Height          =   495
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   525
         Width           =   1680
      End
      Begin VB.OptionButton optUsername 
         Caption         =   "Intranet Username"
         Height          =   450
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   225
         Width           =   1905
      End
   End
   Begin TCSPROG.TCSProgressBar prgEmployee 
      Height          =   375
      Left            =   75
      TabIndex        =   3
      Top             =   7425
      Visible         =   0   'False
      Width           =   8550
      _cx             =   15081
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
      BarBackColor    =   -2147483633
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
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6675
      TabIndex        =   2
      Top             =   450
      Width           =   1965
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   6675
      TabIndex        =   1
      Top             =   45
      Width           =   1965
   End
   Begin MSComctlLib.ListView lvIntranetEmployers 
      Height          =   4065
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   7170
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employer Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. of Employees"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblWritingFile 
      Height          =   390
      Left            =   150
      TabIndex        =   27
      Top             =   4275
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Banner background color color"
      Height          =   390
      Left            =   2625
      TabIndex        =   26
      Top             =   5325
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "Banner text color"
      Height          =   390
      Left            =   150
      TabIndex        =   24
      Top             =   5325
      Width           =   915
   End
   Begin VB.Label lblBannerTitle 
      Caption         =   "Banner title"
      Height          =   390
      Left            =   150
      TabIndex        =   16
      Top             =   4875
      Width           =   915
   End
   Begin VB.Label lblUserInfoHTML 
      Caption         =   "User info"
      Height          =   615
      Left            =   150
      TabIndex        =   14
      Top             =   5925
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Location to write files to:"
      Height          =   495
      Left            =   150
      TabIndex        =   9
      Top             =   6750
      Width           =   1815
   End
End
Attribute VB_Name = "F_Intranet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFrmGeneral

Private m_InvalidVT As Control
Public OK As Boolean

Private Sub chkCaseSensitive_Click()
  p11d32.Intranet.CaseSensitiveOnFullAuthentication = ChkBoxToBool(chkCaseSensitive)
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub


Private Sub cmdRun_Click()
 Call IntranetRun
End Sub

Public Sub IntranetRun()
  Dim i As Integer
  Dim benEmployee As IBenefitClass
  
  On Error GoTo IntranetRun_ERR
  
  Call xSet("IntranetRun")
  
  Call SetCursor
  
  
  If ValidateIntranetData Then
      Call p11d32.Intranet.XMLFile(prgEmployee, lblWritingFile)
  End If
  
  
  
IntranetRun_END:
  Call ClearCursor
  Call xReturn("IntranetRun")
  Exit Sub
IntranetRun_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IntranetRun", "Intranet Run", "Error running the intranet.")
  Resume IntranetRun_END
End Sub


Private Function ValidateIntranetData() As Boolean
  Dim ey As Employer
  Dim benEmployer As IBenefitClass
  Dim i As Long
  Dim li As ListItem
  Dim bAnySelected As Boolean
  Call xSet("ValidateIntranetData")
  
  On Error GoTo ValidateIntranetData_ERR
  
  If InvalidFields(Me) Then Call Err.Raise(ERR_INVALID_FIELDS, "ValidateIntranetData", "Some of the data entry fields are invalid, please amend.")
  If Not FileExists(fb.Directory, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "ValidateIntranetData", "The directory " & fb.Directory & " does not exist.")
  For i = 1 To lvIntranetEmployers.listitems.Count
    Set li = lvIntranetEmployers.listitems(i)
    If Not li.Checked Then GoTo NEXT_ITEM
    Set ey = p11d32.Employers(li.Tag)
    If Not ey.IntranetValid(True) Then GoTo ValidateIntranetData_END
    Set benEmployer = ey
    benEmployer.value(employer_IntranetSelected) = li.Checked
    bAnySelected = True
NEXT_ITEM:
  Next
    
  If Not bAnySelected Then Call Err.Raise(ERR_NO_EMPLOYER, ErrorSource(Err, "ValidateIntranetData"), "No employers selected")
    
     
  ValidateIntranetData = True
  
ValidateIntranetData_END:
  Call xSet("ValidateIntranetData")
  Exit Function
ValidateIntranetData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateIntranetData", "Validate Intranet Data", "Error validating the data for Intranet.")
  Resume ValidateIntranetData_END
  Resume
End Function


Private Sub Form_Load()
  Call SetButtons(False)
  Set lvIntranetEmployers.SmallIcons = MDIMain.imlTree
  Call ColumnWidths(lvIntranetEmployers, 50)
  Call SettingsToScreen
End Sub
Private Sub SettingsToScreen()
  optAuthType(p11d32.Intranet.AuthenticationType).value = True
  optP11DType(p11d32.Intranet.OutputType).value = True
  optUsername(p11d32.Intranet.LoginUserNameSource).value = True
  txtBannerTitle = p11d32.Intranet.BannerTitle
  txtUserInfoHTML = p11d32.Intranet.UserInfoHTML
  chkCaseSensitive.value = BoolToChkBox(p11d32.Intranet.CaseSensitiveOnFullAuthentication)
  Call ColorSettings
End Sub
Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  
End Property

Private Sub lvIntranetEmployers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call EditEmployerDetails(Button, Shift, X, Y)
End Sub

Private Function SetButtons(bVisible As Boolean) As Boolean
  SetButtons = bVisible
End Function

Private Sub EditEmployerDetails(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim ibf As IBenefitForm2
  Dim li As ListItem, liMe As ListItem, liPrevious As ListItem
  
  On Error GoTo EditEmployerDetails_ERR
   
  Call xSet("EditEmployerDetails")
  
  If Not ((Button And vbRightButton) = vbRightButton) Then GoTo EditEmployerDetails_END
    
  Set ibf = CurrentForm
    
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, Err, "The benfit for is nothing")
  If Not ibf.benclass = BC_EMPLOYER Then Call Err.Raise(ERR_IS_NOTHING, Err, "Edit Employer Details")
  
  Set liMe = lvIntranetEmployers.HitTest(X, Y)
  
  If liMe Is Nothing Then GoTo EditEmployerDetails_END
  
  For Each li In ibf.lv.listitems
    If li.Tag = liMe.Tag Then
      Set liPrevious = ibf.lv.SelectedItem
      Set ibf.lv.SelectedItem = li
      Call p11d32.EditEmployer(li.Tag)
      Call p11d32.Intranet.UpdateListViewItem(liMe, p11d32.Employers(li.Tag))
      Set ibf.lv.SelectedItem = liPrevious
      Exit For
    End If
  Next
  
EditEmployerDetails_END:
  Call xReturn("EditEmployerDetails")
  Exit Sub
EditEmployerDetails_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EditEmployerDetails", "Edit Employer Details", "Error editing Employer Details from the intranet screen.")
  Resume EditEmployerDetails_END
  Resume
End Sub

Private Sub fb_Ended()
  p11d32.Intranet.OutputDirectory = fb.Directory
End Sub

Private Sub fb_Started()
  fb.Directory = p11d32.Intranet.OutputDirectory
End Sub

Private Sub optAuthType_Click(Index As Integer)
  p11d32.Intranet.AuthenticationType = Index
End Sub

Private Sub optP11DType_Click(Index As Integer)
  p11d32.Intranet.OutputType = Index
End Sub

Private Sub optUsername_Click(Index As Integer)
  p11d32.Intranet.LoginUserNameSource = Index
End Sub
Private Function GetColor(ByVal DEFAULT As Long) As Long
On Error GoTo err_err

  GetColor = DEFAULT
  cdlg.Color = DEFAULT
  cdlg.CancelError = True
  
  cdlg.ShowColor
  GetColor = cdlg.Color
  
err_end:
  Exit Function
err_err:
  Resume err_end
End Function
Private Sub ColorSettings()
  pctBannerBackColor.BackColor = p11d32.Intranet.BannerBackColor
  pctBannerForeColor.BackColor = p11d32.Intranet.BannerForeColor
End Sub
Private Sub pctBannerBackColor_Click()
  p11d32.Intranet.BannerBackColor = GetColor(p11d32.Intranet.BannerBackColor)
  Call ColorSettings
End Sub

Private Sub pctBannerForeColor_Click()
  p11d32.Intranet.BannerForeColor = GetColor(p11d32.Intranet.BannerForeColor)
  Call ColorSettings
End Sub

Private Sub txtBannerTitle_Change()
 p11d32.Intranet.BannerTitle = txtBannerTitle.Text
End Sub

Private Sub txtUserInfoHTML_Change()
  p11d32.Intranet.UserInfoHTML = txtUserInfoHTML.Text
End Sub
