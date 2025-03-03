VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A7CE771F-05B2-43CF-9650-ED841A9049FA}#1.0#0"; "atc3FolderBrowser.ocx"
Begin VB.Form F_ImportTracking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Tracking"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin atc3FolderBrowser.FolderBrowser fb 
      Height          =   420
      Left            =   90
      TabIndex        =   8
      Top             =   4005
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   741
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4455
      TabIndex        =   7
      Top             =   4455
      Width           =   1140
   End
   Begin VB.CheckBox chkImportTracking 
      Caption         =   "Enable import tracking"
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   4455
      Width           =   2805
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5670
      TabIndex        =   3
      Top             =   4455
      Width           =   1140
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   3015
      TabIndex        =   2
      Top             =   4455
      Width           =   1140
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   2220
      Left            =   45
      TabIndex        =   1
      Top             =   1530
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   3916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Tracked import saved folder"
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   3735
      Width           =   2535
   End
   Begin VB.Label lblCurrentEmployer 
      Caption         =   "lblCurrentEmployer"
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   1170
      Width           =   6675
   End
   Begin VB.Label lblInfo 
      Caption         =   $"F_ImportTracking.frx":0000
      Height          =   915
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   6675
   End
End
Attribute VB_Name = "F_ImportTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IEnumFiles
Private m_GuidEmployer As String
Private m_EmployerFileName As String
Private Sub chkImportTracking_Click()
  p11d32.Importing.Tracking = ChkBoxToBool(chkImportTracking)
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
  Unload Me
End Sub

Private Sub cmdRestore_Click()
  Call Restore
End Sub


Private Sub Restore()
  Dim LVI As ListItem
  Dim sDestFile As String
  Dim Unzip As cUnzip
  On Error GoTo err_Err
  
  Set LVI = lvFiles.SelectedItem
  If LVI Is Nothing Then GoTo err_End
  
  sDestFile = p11d32.WorkingDirectory & m_EmployerFileName
  'try ant get an exclusive lock on the file using file open
  If IsFileOpen(sDestFile, False) Then Call Err.Raise(ERR_EMPLOYER_INVALID, "Restore", "The employer file '" & sDestFile & "' is open. To perform a restore no one can be using the file.")
  Set Unzip = New cUnzip
  Unzip.ZipFile = LVI.Tag
  Unzip.UnzipFolder = p11d32.WorkingDirectory
  Unzip.ExtractOnlyNewer = False
  Unzip.OverwriteExisting = True
  Call Unzip.Unzip
  Call MsgBox("File successfully restored to " & sDestFile, vbInformation)
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Restore", "Restore", Err.Description)
  Resume err_End
End Sub

Private Sub cmdOK_Click()
  Call cmdCancel_Click
End Sub

Private Sub Form_Load()
  Call SettingsToScreen
End Sub
Public Sub SettingsToScreen()
  Dim ben As IBenefitClass
  Dim ey As Employer
  Dim LVI As ListItem
  Dim sEmployer As String
  On Error GoTo err_Err
  
  If (p11d32.Employers.Count = 0) Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "SettingsToScreen", "There are no employer files available")
  End If
    
  'runs the fixes add adds the guid
  Set ey = p11d32.Importing.PreImport(True)
  Set ben = ey
  Call ben.Kill
    
  fb.Directory = p11d32.Importing.TrackingPath
    
  chkImportTracking.value = BoolToChkBox(p11d32.Importing.Tracking)
  lblCurrentEmployer.Caption = ""
  
  sEmployer = "'" & ben.Name & "'"
  m_EmployerFileName = ben.value(employer_FileName)
  m_GuidEmployer = ben.value(employer_GUID_db)
  
  
  Call ben.Kill
  If (Len(ben.value(employer_GUID_db)) = 0) Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, ErrorSource(Err, "SettingsToScreen"), "The employer " & sEmployer & " needs to be opened first. Please open the employer")
  End If
  lblCurrentEmployer = "Import restores available for " & sEmployer
  Call ColumnWidths(Me.lvFiles, 20, 20, 60)
  
  Call UpdateList
  
err_End:
  Exit Sub
err_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Setting To Screen", "Failed to do settings to screen")
  Resume err_End
End Sub
Private Sub UpdateList()

  Call lvFiles.listitems.Clear
  Call EnumFiles("", p11d32.Importing.TrackingPath, "*" & S_FILE_EXTENSION_BAK, Me)
  
End Sub
Private Sub IEnumFiles_File(ByVal vData As Variant, ByVal sPathAndFile As Variant, ByVal sFile As String)
  Dim sGuid As String
  Dim sTime As String, sMinutes As String, sHour As String, sSeconds As String
  Dim sDate As String, sDay As String, sMonth As String, sYear As String
  Dim sComment As String
  Dim iLen As Long
  Dim LVI As ListItem
  Dim p0 As Long
  Dim p1 As Long
  
  On Error GoTo err_Err
  
  iLen = Len(sFile)
  sFile = Left$(sFile, iLen - Len(S_FILE_EXTENSION_BAK))
  iLen = Len(sFile)
  
  p0 = iLen - 1
  sSeconds = Mid$(sFile, p0, 2)
  p0 = p0 - 2
  sMinutes = Mid$(sFile, p0, 2)
  p0 = p0 - 2
  sHour = Mid$(sFile, p0, 2)
  p0 = p0 - 2
  sDay = Mid$(sFile, p0, 2)
  p0 = p0 - 2
  sMonth = Mid$(sFile, p0, 2)
  p0 = p0 - 4
  sYear = Mid$(sFile, p0, 4)
  p0 = p0 - 2
  p1 = InStrRev(sFile, "_", p0)
  sGuid = Mid$(sFile, p1 + 1, p0 - p1)

  sComment = Left$(sFile, p1 - 1)
  
  If (StrComp(sGuid, m_GuidEmployer, vbTextCompare) <> 0) Then GoTo err_End
  
  
  Set LVI = lvFiles.listitems.Add()
  LVI.Tag = sPathAndFile
  sTime = sHour & ":" & sMinutes & ":" & sSeconds
  sDate = sDay & "/" & sMonth & "/" & sYear
  
  LVI.Text = sDate
  Call LVI.ListSubItems.Add(, , sTime)
  Call LVI.ListSubItems.Add(, , sComment)
  
err_End:
  Exit Sub
err_Err:
  Resume err_End
End Sub
Private Sub fb_Ended()
  On Error GoTo err_Err
  
  p11d32.Importing.TrackingPath = fb.Directory
  Call UpdateList
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Ended", "Ended", Err.Description)
  Resume err_End
End Sub

Private Sub fb_Started()
  fb.Directory = p11d32.Importing.TrackingPath
End Sub

