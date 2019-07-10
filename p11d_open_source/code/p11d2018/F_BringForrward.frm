VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "atc2stat.ocx"
Begin VB.Form F_BringForward 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bring Forward"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConvertToCurrentYear 
      Caption         =   "Convert to current year"
      Height          =   615
      Left            =   5925
      TabIndex        =   15
      Top             =   4275
      Width           =   1365
   End
   Begin VB.PictureBox pctFrame 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5985
      ScaleHeight     =   735
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   2205
      Width           =   1275
      Begin VB.OptionButton optType 
         Caption         =   "Overwrite"
         Height          =   330
         Index           =   0
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   1185
      End
      Begin VB.OptionButton optType 
         Caption         =   "Update"
         Height          =   330
         Index           =   1
         Left            =   45
         TabIndex        =   13
         Top             =   315
         Width           =   1185
      End
   End
   Begin atc2stat.TCSStatus sts 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   5835
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   556
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
   Begin VB.Frame fraOther 
      Caption         =   "Other"
      Height          =   870
      Left            =   5895
      TabIndex        =   10
      Top             =   3015
      Width           =   1410
      Begin VB.CheckBox chkNewFilesForAll 
         Caption         =   "Always create new files"
         Height          =   600
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdSections 
      Caption         =   "&Sections"
      Height          =   420
      Left            =   5895
      TabIndex        =   7
      Top             =   1530
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   420
      Left            =   5895
      TabIndex        =   6
      Top             =   1080
      Width           =   1410
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   420
      Left            =   5895
      TabIndex        =   4
      Top             =   630
      Width           =   1410
   End
   Begin VB.Frame fraType 
      Caption         =   "Type"
      Height          =   960
      Left            =   5895
      TabIndex        =   3
      Top             =   2025
      Width           =   1410
   End
   Begin MSComctlLib.ListView lvCurrentYearFiles 
      Height          =   2580
      Left            =   45
      TabIndex        =   1
      Top             =   3285
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   4551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvPriorYearFiles 
      Height          =   2490
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblPriorYearFiles 
      Caption         =   "Prior year files"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   450
      Width           =   2220
   End
   Begin VB.Label lblCurentYearFiles 
      Caption         =   "Current year files"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   3105
      Width           =   2085
   End
   Begin VB.Label lblEmployersIn 
      Caption         =   "lblEmployersIn"
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   7245
   End
End
Attribute VB_Name = "F_BringForward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IEnumEmployers

Private m_liLastCurrent As ListItem
Public P As TCSPANEL


Private Sub chkNewFilesForAll_Click()
  p11d32.BringForward.NewFilesForAll = ChkBoxToBool(chkNewFilesForAll)
  'Call SetNewFilesForAll
End Sub
Private Sub SetNewFilesForAll()
  Dim li As ListItem
  
  If p11d32.BringForward.NewFilesForAll Then
    For Each li In lvCurrentYearFiles.listitems
      If li.SmallIcon = IMG_SELECTED Then
        Call CurrentYearClick(li)
        Exit For
      End If
    Next
  End If
End Sub
Private Sub cmdCancel_Click()
  Unload Me
  
End Sub

Private Sub cmdOptions_Click()
  
End Sub

Private Sub cmdConvertToCurrentYear_Click()
  p11d32.BringForward.ConvertToCurrentYear
End Sub

Private Sub cmdRun_Click()
  p11d32.BringForward.ProcessFiles
End Sub

Private Sub cmdSections_Click()
  Call p11d32.BringForward.Options
End Sub

Private Sub Form_Load()
  Call SettingsToScreen
  
End Sub

Private Sub Form_Paint()
  Static b As Boolean
  
  If Not b Then
    DoEvents
    Set p11d32.BringForward.OLCurrentEmployers = New ObjectList
    Set p11d32.BringForward.OLPriorEmployers = New ObjectList

    Call Start
    b = True
  End If
End Sub
Public Sub SettingsToScreen()
  On Error GoTo SettingsToScreen_Err
  
  Call xSet("SettingsToScreen")
  
  Call optType_Click(-1)
  chkNewFilesForAll.value = BoolToChkBox(p11d32.BringForward.NewFilesForAll)
  Set P = sts.AddPanel(30)
  lblCurentYearFiles = p11d32.AppYear & "/" & p11d32.AppYear + 1 & " files."    'RH
  lblPriorYearFiles = p11d32.AppYear - 1 & "/" & p11d32.AppYear & " files."     'RH
  lvCurrentYearFiles.SmallIcons = MDIMain.imlTree
  lvPriorYearFiles.SmallIcons = MDIMain.imlTree
  
  Me.Caption = "Bring forward for " & p11d32.AppYear & "/" & p11d32.AppYear + 1 & " files."   'RH
  
  lblEmployersIn = "Employers in " & p11d32.WorkingDirectory
  'add column headers
  Call lvPriorYearFiles.ColumnHeaders.Add(1, , F_Employers.LB.ColumnHeaders(ELVC_EMPLOYER_NAME))
  Call lvPriorYearFiles.ColumnHeaders.Add(2, , F_Employers.LB.ColumnHeaders(ELVC_FILE_NAME))
  Call lvPriorYearFiles.ColumnHeaders.Add(3, , "Brought forward")
  Call lvPriorYearFiles.ColumnHeaders.Add(4, , "Fix level")
  Call ColumnWidths(lvPriorYearFiles, 40, 25, 25, 10)
  
  Call lvCurrentYearFiles.ColumnHeaders.Add(1, , F_Employers.LB.ColumnHeaders(ELVC_EMPLOYER_NAME))
  Call lvCurrentYearFiles.ColumnHeaders.Add(2, , F_Employers.LB.ColumnHeaders(ELVC_FILE_NAME))
  Call lvCurrentYearFiles.ColumnHeaders.Add(3, , "Fix level")
  Call ColumnWidths(lvCurrentYearFiles, 50, 40, 10)
  cmdConvertToCurrentYear.Visible = IsRunningInIDE
  
SettingsToScreen_End:
  Call xReturn("SettingsToScreen")
  Exit Sub
SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the setting to the screen for F_BringForward.")
  Resume SettingsToScreen_End
End Sub


Private Sub lvCurrentYearFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call CurrentYearClick(Item)
End Sub
Private Sub CurrentYearClick(Item As ListItem)
  On Error GoTo CurrentYearClick_ERR
  
  Call xSet("CurrentYearClick")
  
  If Not m_liLastCurrent Is Nothing Then
    If Item Is m_liLastCurrent Then
      If m_liLastCurrent.SmallIcon = IMG_UNSELECTED Then
        m_liLastCurrent.SmallIcon = IMG_SELECTED
      Else
        m_liLastCurrent.SmallIcon = IMG_UNSELECTED
      End If
    Else
      m_liLastCurrent.SmallIcon = IMG_UNSELECTED
      Item.SmallIcon = IMG_SELECTED
    End If
  Else
    Item.SmallIcon = IMG_SELECTED
  End If
  
  Set m_liLastCurrent = Item


CurrentYearClick_END:
  Call xReturn("CurrentYearClick")
  Exit Sub
CurrentYearClick_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CurrentYearClick", "Current Year Click", "Error clicking on a current year files list item.")
  Resume CurrentYearClick_END
End Sub
Private Sub lvPriorYearFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Item.SmallIcon = IMG_SELECTED Then
    Item.SmallIcon = IMG_UNSELECTED
  Else
    Item.SmallIcon = IMG_SELECTED
  End If
End Sub

Private Sub optType_Click(Index As Integer)
    
  Dim i As Long
  
  Select Case Index
    Case -1
      For i = 0 To 1
        If i = p11d32.BringForward.BringForwardType Then
          optType(i).value = True
          Exit For
        End If
      Next
    Case Else
      p11d32.BringForward.BringForwardType = Index
  End Select

End Sub
Private Sub IEnumEmployers_Count(ByVal Count As Long)
  Dim s As String
  
  Select Case p11d32.BringForward.YOE
    Case YOE_CURRENT
      s = "Analysing current year files"
    Case YOE_PRIOR
      s = "Analysing prior year files"
  End Select
  
  Call sts.StartPrg(Count, s, ValueOfMax)
End Sub

Private Sub IEnumEmployers_CurrentFile(ByVal sCurrentFile As String)
  sts.StepCaption (sCurrentFile)
End Sub

Private Sub IEnumEmployers_Employer(ey As Object)
  Dim empr As Employer
  
  Set empr = ey
  Select Case p11d32.BringForward.YOE
    Case YOE_CURRENT
      Call p11d32.BringForward.OLCurrentEmployers.Add(empr)
    Case YOE_PRIOR
      Call p11d32.BringForward.OLPriorEmployers.Add(empr)
  End Select
  
End Sub

Private Sub Start()

  On Error GoTo Start_ERR
  
  Call xSet("Start")
  
  Call ChDir(p11d32.WorkingDirectory)
  
  p11d32.BringForward.YOE = YOE_PRIOR
  Call EnumEmployerFiles(p11d32.Rates.FileExtensionPrior, Me)
  p11d32.BringForward.YOE = YOE_CURRENT
  Call EnumEmployerFiles(p11d32.Rates.FileExtensionCurrent, Me)
  Call p11d32.BringForward.EmployersToListView
Start_END:
  Call sts.StopPrg
  Call xReturn("Start")
  Exit Sub
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error starting the bring forward process.")
  Resume Start_END
  Resume
End Sub

