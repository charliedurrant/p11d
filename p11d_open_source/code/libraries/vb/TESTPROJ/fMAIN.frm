VERSION 5.00
Object = "{4BA5AE86-C9BA-4B77-8E15-D04582204FDD}#1.0#0"; "atc2stat.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6600
   ClientLeft      =   270
   ClientTop       =   840
   ClientWidth     =   9585
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin atc2Stat.TCSStatus Status1 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   5625
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   1720
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
   Begin MSComctlLib.Toolbar TB_Main 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "TB_ImageList_Normal"
      DisabledImageList=   "TB_ImageList_Normal"
      HotImageList    =   "TB_ImageList_Hot"
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList TB_ImageList_Hot 
      Left            =   225
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList TB_ImageList_Normal 
      Left            =   240
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuExportDB 
         Caption         =   "Test Export DB"
      End
      Begin VB.Menu mnuTestGo 
         Caption         =   "Test"
      End
      Begin VB.Menu mnuTCSNotes 
         Caption         =   "Test TCS Notes"
      End
      Begin VB.Menu mnuTestMAPI 
         Caption         =   "Test MAPI"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Test Status"
      End
      Begin VB.Menu mnuAddMenu 
         Caption         =   "Add &Menus"
      End
      Begin VB.Menu mnuTestLDAP 
         Caption         =   "Test LDAP"
      End
      Begin VB.Menu mnuErrorTest 
         Caption         =   "Test Error"
      End
      Begin VB.Menu mnuStringTest 
         Caption         =   "Test Strings"
      End
      Begin VB.Menu mnuSortTest 
         Caption         =   "Test Sorting"
      End
      Begin VB.Menu mnuGridADO 
         Caption         =   "Test Grid (ADO)"
      End
      Begin VB.Menu mnuAutoTestADO 
         Caption         =   "Test Auto (ADO)"
      End
      Begin VB.Menu mnuAutoTest 
         Caption         =   "Test Auto (DAO)"
      End
      Begin VB.Menu mnuAutoTestRDO 
         Caption         =   "Test Auto (RDO)"
      End
      Begin VB.Menu mnuTestListView 
         Caption         =   "Test Listview"
      End
      Begin VB.Menu mnuRepWizard 
         Caption         =   "Test Report Wizard"
      End
      Begin VB.Menu mnuTestWhere 
         Caption         =   "Test Where Control"
      End
      Begin VB.Menu mnuDebugSQL 
         Caption         =   "Test Debug SQL"
      End
      Begin VB.Menu mnuDebugSQLADOTestJet 
         Caption         =   "Test Debug SQL (ADO) - Jet"
      End
      Begin VB.Menu mnuDebugSQLTestOracle 
         Caption         =   "Test Debug SQL (ADO) - Oracle"
      End
      Begin VB.Menu mnuDebugSQLTestSQLServer 
         Caption         =   "Test Debug SQL (ADO) - SQL"
      End
      Begin VB.Menu mnuImportTest 
         Caption         =   "Test Import"
      End
      Begin VB.Menu mnuImportWizTest 
         Caption         =   "Test Import Wizard"
      End
      Begin VB.Menu mnuExitBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindowList 
      Caption         =   "&Window"
      WindowList      =   -1  'True
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

Private Sub MDIForm_Load()
  gbMDILoaded = True
  Set clsEvent = gEvents
End Sub

Private Sub clsEvent_DebugMenuItem(Name As String, Index As Long, Parent As TCSMenuItems)
  On Error GoTo clsEvent_DebugMenuItem_err
  Select Case Parent
    Case MNU_BREAK
      Call ExitApp(True)
    Case Else
      ECASE " clsEvent_DebugMenuItem"
  End Select

clsEvent_DebugMenuItem_end:
  Exit Sub
  
clsEvent_DebugMenuItem_err:
  Call ErrorMessage(ERR_ERROR, Err, "DebugMenuItem", "ERR_DEBUGMENU", "Error processing the debug menu event " & Name & ".")
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

Private Sub mnuAddMenu_Click()
  Call AddMenus(Me.hWnd)
End Sub

Private Sub mnuAutoTest_Click()
  frmGrid.Show
End Sub

Private Sub mnuAutoTestADO_Click()
  frmGridADO.Show
End Sub

Private Sub mnuAutoTestRDO_Click()
  frmGridRDO.Show
End Sub

Private Sub mnuDebugSQL_Click()
   Call SQLTestDebug
End Sub

Private Sub mnuDebugSQLADOTest_Click()
  
End Sub



Private Sub mnuDebugSQLADOTestJet_Click()
  Call SQLADOTestDebug(DB_TARGET_JET)
End Sub

Private Sub mnuDebugSQLTestOracle_Click()
  Call SQLADOTestDebug(DB_TARGET_ORACLE)
End Sub

Private Sub mnuDebugSQLTestSQLServer_Click()
  Call SQLADOTestDebug(DB_TARGET_SQLSERVER)
End Sub

Private Sub mnuErrorTest_Click()
  Call TestError
End Sub

Private Sub mnuExit_Click()
  Call ExitApp
End Sub

Private Sub mnuExportDB_Click()
  frmExportToXML.Show
End Sub

Private Sub mnuGridADO_Click()
  frmGridTest.Show
End Sub

Private Sub mnuHelpAbout_Click()
  Call AppAbout
End Sub

Private Sub mnuImportTest_Click()
  Call ImportTest
End Sub

Private Sub mnuImportWizTest_Click()
  Call ImportWizardTest
End Sub

Private Sub mnuRepWizard_Click()
  Set RW = New ReportWizard
  Set RW.ReportInterface = New TestUDM
  Set RW.IRepProcess = New TestRepInterface
  RW.Title = "UDM Test"
  RW.IgnoreUserError = True
  RW.StartReportWizard
End Sub

Private Sub mnuSortTest_Click()
  frmSortTest.Show
End Sub

Private Sub mnuStatus_Click()
  frmStatus.Show
End Sub

Private Sub mnuStringTest_Click()
  frmStringTest.Show
End Sub

Private Sub mnuTCSNotes_Click()
  frmTestNotes.Show
End Sub

Private Sub mnuTestGo_Click()
  Dim X As TCSFileread, buffer As String
  Dim dirlist As StringList
  Dim s As String, sval As String
  
  Set dirlist = FindFiles("C:\WINNT", "*.dll", True)
  Set X = New TCSFileread
  If Not X.OpenFile("Filename") Then
  ' failed to open
  End If
  Do While X.GetLine(buffer)
    
  Loop
  Set X = Nothing


  
  
  s = FileSaveAsDlg("XXX", "Text Files|*.txt|All Files|*.*", "C:\caesar")
    
  s = " ""hello"" , ""bye"" "
  Debug.Print GetDelimitedValue(sval, s, 1)
  Debug.Print sval
End Sub

Private Sub mnuTestLDAP_Click()
  Call TestLDAP
End Sub

Private Sub mnuTestListView_Click()
  frmListViewTest.Show
End Sub

Private Sub mnuTestMAPI_Click()
  frmMAPIMail.Show
End Sub

Private Sub mnuTestWhere_Click()
  frmWhere.Show
End Sub

Private Sub TB_Main_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call gTBCtrl.ExecuteButton(Button.Key)
End Sub

Private Sub TB_Main_DblClick()
  'Me.TB_Main.Customize
End Sub
' END --- DO NOT CHANGE CODE ABOVE
