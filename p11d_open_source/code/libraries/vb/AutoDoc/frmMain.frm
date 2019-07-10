VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D7D47D2E-20A1-45D1-B08B-3A509726296E}#1.0#0"; "atc2split.OCX"
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "atc2stat.ocx"

Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Auto documentor"
   ClientHeight    =   6630
   ClientLeft      =   270
   ClientTop       =   840
   ClientWidth     =   9510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox pctSearch 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9510
      TabIndex        =   5
      Top             =   0
      Width           =   9510
      Begin VB.CheckBox chkViewCode 
         Caption         =   "View Code"
         Height          =   240
         Left            =   6840
         TabIndex        =   12
         Top             =   270
         Width           =   1815
      End
      Begin VB.CheckBox chkShowCategories 
         Caption         =   "Show Categories"
         Height          =   240
         Left            =   4905
         TabIndex        =   11
         Top             =   270
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.ComboBox cboSearchWhat 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   225
         Width           =   1950
      End
      Begin VB.TextBox txtSearch 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   2715
      End
      Begin VB.Label lblSearchWhat 
         Caption         =   "What"
         Height          =   195
         Left            =   2835
         TabIndex        =   10
         Top             =   45
         Width           =   960
      End
      Begin VB.Label lblSearch 
         Caption         =   "Search:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   45
         Width           =   825
      End
   End
   Begin VB.PictureBox pctSearchResults 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   0
      ScaleHeight     =   1860
      ScaleWidth      =   9510
      TabIndex        =   4
      Top             =   4440
      Width           =   9510
      Begin MSComctlLib.ListView lvSearchResults 
         Height          =   1770
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   3122
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlTickCross"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList imlTickCross 
      Left            =   5805
      Top             =   2790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pct 
      Align           =   3  'Align Left
      Height          =   3750
      Left            =   0
      ScaleHeight     =   3690
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   690
      Width           =   3735
      Begin MSComctlLib.TreeView tvw 
         Height          =   6165
         Left            =   90
         TabIndex        =   3
         Top             =   0
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   10874
         _Version        =   393217
         Indentation     =   0
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlTickCross"
         Appearance      =   1
      End
      Begin ATC2SPLIT.SPLIT split 
         Height          =   6180
         Left            =   3600
         TabIndex        =   2
         Top             =   0
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   10901
      End
   End
   Begin ATC2Stat.TCSStatus sts 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6300
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExitBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAboutSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F11}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' START --- DO NOT CHANGE CODE BELOW
Private WithEvents clsEvent As TCSEventClass
Attribute clsEvent.VB_VarHelpID = -1
Private LastButton As Button


Private Sub cmdOpenVBP_Click()
  'apf Call Documentor.ChoseFile
  
End Sub

Private Sub cboSearchWhat_Click()
  Call SearchAutoDoc(txtSearch, cboSearchWhat.ItemData(cboSearchWhat.ListIndex))
End Sub

Private Sub chkShowCategories_Click()
  Call ProjectsToScreen(frmMain.chkShowCategories.Value = vbChecked)
End Sub

Private Sub chkViewCode_Click()
  Call NodeClick(tvw.Nodes(gLastFunctionKey), False, False)
End Sub

Private Sub lvSearchResults_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call NodeClick(tvw.Nodes(Item.Key), False)
End Sub

Private Sub MDIForm_Load()
  On Error Resume Next
  gbMDILoaded = True
  Set clsEvent = gEvents
  Call split.Initialise(frmMain.hWnd, True)
  split.MinBorderPixels = 40
End Sub

Private Sub clsEvent_DebugMenuItem(Name As String, Index As Long, Parent As TCSMenuItems)
  On Error GoTo clsEvent_DebugMenuItem_err
  Select Case Parent
    Case MNU_BREAK
      Call ExitApp(True)
    Case Else
      ECASE "clsEvent_DebugMenuItem Unknown menu item: " & Name
  End Select

clsEvent_DebugMenuItem_end:
  Exit Sub
  
clsEvent_DebugMenuItem_err:
  Call ErrorMessage(ERR_ERROR, Err, "DebugMenuItem", "Debug menu", "Error processing the debug menu event " & Name & ".")
  Resume clsEvent_DebugMenuItem_end
End Sub
   
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim doExitApp As Boolean
  
  gbAllowAppExit = True
  If FatalError Or (UnloadMode = vbAppTaskManager) Or (UnloadMode = vbAppWindows) Then
    gbForceExit = True
  End If
  doExitApp = gbForceExit
  If Not doExitApp Then
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    doExitApp = DisplayMessage(Me, "Are you sure you want to exit " & AppName & "?", AppName, "Yes", "No")
  End If
  If doExitApp Then
    doExitApp = UserAppShutDown Or FatalError
    If doExitApp Then Call ExitApp
  End If
  If Not doExitApp Then
    Cancel = True
    gbAllowAppExit = False
  End If
End Sub

Private Sub mnuExit_Click()
  Call ExitApp
End Sub

Private Sub mnuHelpAbout_Click()
  Call AppAbout
End Sub
' END --- DO NOT CHANGE CODE ABOVE

Private Sub mnuRefresh_Click()
  Call Reinitialise
End Sub

Private Sub pct_Resize()
  tvw.Top = L_GAP
  tvw.Height = pct.Height - (tvw.Top + L_GAP)
  tvw.Left = L_GAP
  tvw.Width = pct.Width - (10 * L_GAP)
End Sub

Private Sub pctSearchResults_Resize()
  lvSearchResults.Top = L_GAP
  lvSearchResults.Height = pctSearchResults.Height - (lvSearchResults.Top + L_GAP)
  lvSearchResults.Left = L_GAP
  lvSearchResults.Width = pctSearchResults.Width - (10 * L_GAP)
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
  Call NodeClick(Node, True)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Call SearchAutoDoc(txtSearch.Text, frmMain.cboSearchWhat.ItemData(frmMain.cboSearchWhat.ListIndex))
End Sub
