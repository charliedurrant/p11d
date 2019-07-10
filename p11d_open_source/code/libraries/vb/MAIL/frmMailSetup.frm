VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMailSetup 
   Caption         =   "Mail Setup"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8715
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   4500
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   28
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   22
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame fraProfile 
      Caption         =   "Default profile/Login name"
      Height          =   975
      Left            =   5640
      TabIndex        =   20
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtLoginName 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame fraMailApplication 
      Caption         =   "Mail Application"
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Microsoft Outlook"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Tag             =   "6"
         Top             =   960
         Width           =   2745
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Pegasus"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   19
         Tag             =   "9"
         Top             =   2760
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Novell GroupWise v5.x +"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Tag             =   "7"
         Top             =   2400
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Lotus Notes v4.5 +"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Tag             =   "5"
         Top             =   2040
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Lotus Notes"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Tag             =   "4"
         Top             =   1680
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Lotus cc:Mail"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Tag             =   "3"
         Top             =   1320
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Microsoft Exchange"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Tag             =   "2"
         Top             =   585
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Microsoft Mail"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Tag             =   "1"
         Top             =   240
         Width           =   2500
      End
      Begin VB.OptionButton optMailApplication 
         Caption         =   "Other (Specify Mail System)"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Tag             =   "0"
         Top             =   3120
         Width           =   2500
      End
   End
   Begin VB.Frame fraMailSystem 
      Caption         =   "Mail System"
      Height          =   4335
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optMailSystem 
         Caption         =   "Unknown"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3720
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "NotesAPI"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Tag             =   "11"
         Top             =   3360
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "RAS"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Tag             =   "12"
         Top             =   3000
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "Active Messaging"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Tag             =   "9"
         Top             =   2640
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "VINES"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Tag             =   "20"
         Top             =   2280
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "SMTP/POP"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Tag             =   "10"
         Top             =   1920
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "CSERVE"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1560
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "MAPI (Extended) - prevents outlook asking to send emails"
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Tag             =   "3"
         Top             =   960
         Width           =   2115
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "VIM"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "2"
         Top             =   600
         Width           =   2000
      End
      Begin VB.OptionButton optMailSystem 
         Caption         =   "MAPI"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "1"
         Top             =   240
         Width           =   2000
      End
   End
   Begin VB.Frame fraConfiguration 
      Caption         =   "Check configuration"
      Height          =   2775
      Left            =   5640
      TabIndex        =   23
      Top             =   1200
      Width           =   3015
      Begin VB.CommandButton cmdKillOutlook 
         Caption         =   "Kill Outlook"
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdDebug 
         Caption         =   "Debug mail system"
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveSettings 
         Caption         =   "Save settings"
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdTestSettings 
         Caption         =   "Test settings"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdDetectSettings 
         Caption         =   "Detect settings"
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMailSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_Mail As Mail
Private m_DebugAppPathFile As String
Private m_DebugOutputPathFile As String

'Old values for MAIL class
Private m_MailApplication_old As MAIL_APPLICATION
Private m_MailSystem_old As IDSM_MAIL_SYSTEM
Private m_LoginName_old As String

'New values for MAIL class
Private m_MailApplication_new As MAIL_APPLICATION
Private m_MailSystem_new As IDSM_MAIL_SYSTEM
Private m_LoginName_new As String

Private m_MailApplicationChanged As Boolean

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szexeFile As String * 260
End Type
'-------------------------------------------------------
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
ByVal lFlags As Long, lProcessID As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


Private Sub cmdCancel_Click()
  Call SaveComponentChanges(False)
  Call SaveFormChanges(False)
  Unload Me
End Sub
Private Sub cmdDebug_Click()
  sb.SimpleText = "Loading debug application."
  'Mail debug tool will always install to this directory:
  Call Shell(m_DebugAppPathFile, vbNormalFocus)  'RK need to derive from installation
  
End Sub

Private Sub cmdKillOutlook_Click()
  On Error GoTo err_err
  
  Call KillProcess("Outlook.exe")
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "Kill Outlook", "Kill Outlook", Err.Description)
  Resume err_end
End Sub

Public Sub KillProcess(NameProcess As String)
  Const PROCESS_ALL_ACCESS = &H1F0FFF
  Const TH32CS_SNAPPROCESS As Long = 2&
  Dim uProcess  As PROCESSENTRY32
  Dim RProcessFound As Long
  Dim hSnapshot As Long
  Dim SzExename As String
  Dim ExitCode As Long
  Dim MyProcess As Long
  Dim AppKill As Boolean
  Dim AppCount As Integer
  Dim i As Integer
  Dim WinDirEnv As String
        
  If NameProcess <> "" Then
     AppCount = 0

     uProcess.dwSize = Len(uProcess)
     hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
     RProcessFound = ProcessFirst(hSnapshot, uProcess)

     Do
       i = InStr(1, uProcess.szexeFile, Chr(0))
       SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
       WinDirEnv = Environ("Windir") + "\"
       WinDirEnv = LCase$(WinDirEnv)
   
       If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
          AppCount = AppCount + 1
          MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
          AppKill = TerminateProcess(MyProcess, ExitCode)
          Call CloseHandle(MyProcess)
       End If
       RProcessFound = ProcessNext(hSnapshot, uProcess)
     Loop While RProcessFound
     Call CloseHandle(hSnapshot)
  End If

End Sub


Private Sub cmdOk_Click()
  If CheckChanges Then
    If MsgBox("Your Mail settings have changed." & _
              "Do you wish to save them?", vbYesNo, "Save changes") = vbYes Then
      Call SaveComponentChanges(True)
      Call SaveFormChanges(True)
    Else
      Call SaveComponentChanges(False)
    End If
  End If
  Unload Me
End Sub

Private Sub cmdSaveSettings_Click()
  Call SaveComponentChanges(True)
  Call SaveFormChanges(True)
End Sub

Private Sub Form_Load()
  
  Call GetDebugPaths(m_DebugAppPathFile, m_DebugOutputPathFile)
  Call CheckDebugAppPathFile
  m_MailApplicationChanged = False
  
  'Save initial settings
  m_MailApplication_old = m_Mail.MailApplication
  m_MailSystem_old = m_Mail.IDSMail.MailSystem
  m_LoginName_old = m_Mail.IDSMail.LoginName
  
  m_MailApplication_new = m_MailApplication_old
  m_MailSystem_new = m_MailSystem_old
  m_LoginName_new = m_LoginName_old
  'Apply initial settings
  
  Call SetMailApplication(m_MailApplication_new)
  Call SetMailSystem(m_MailSystem_new)
  
  txtLoginName.Text = m_Mail.IDSMail.LoginName
  
  'If m_Mail.MailApplication = MA_OTHER Then
  '  optMailApplication(MA_OTHER).Value = True
  '  Call SetMailApplication(MA_OTHER)
    
  'Else
  '  Call SetMailApplication(m_Mail.MailApplication)
  '
  'End If
  'Call SetMailSystem(m_Mail.MailSystem)
  sb.SimpleText = "Current settings loaded"

End Sub
Private Sub EnableMailApplication(bEnableDisable As Boolean)
  Dim i As Long
  For i = optMailApplication.LBound To optMailApplication.UBound
    Let optMailApplication.Item(i).Enabled = bEnableDisable
  Next i
End Sub
Private Sub EnableMailSystem(bEnableDisable As Boolean)
  Dim i As Long
  For i = optMailSystem.LBound To optMailSystem.UBound
    If StrComp(optMailSystem(i).Tag, m_MailSystem_new, vbTextCompare) = 0 Then
      optMailSystem(i).Enabled = True
    Else
      optMailSystem(i).Enabled = bEnableDisable
    End If
  Next i
End Sub
Private Sub SetMailSystem(MailSystem As IDSM_MAIL_SYSTEM, Optional bSetByApplication As Boolean = False)
  Dim i As Long
  On Error GoTo SetMailSystem_Err
  
  m_MailSystem_new = MailSystem
  
  'If Mail system set by application then restrict options
  If bSetByApplication Then
    If MailSystem = IDSM_SYS_UNKNOWN Then
      EnableMailSystem (False)
    Else
      EnableMailSystem (False)
    End If
  End If
  
  'Set Mail System
  For i = optMailSystem.LBound To optMailSystem.UBound
    If StrComp(optMailSystem(i).Tag, MailSystem, vbTextCompare) = 0 Then
      optMailSystem(i).Value = True
      optMailSystem(i).Enabled = True
      Exit For
    End If
  Next i


  Call RedemptionEnableExtendedMapi

SetMailSystem_End:
  Exit Sub

SetMailSystem_Err:
  Resume SetMailSystem_End
End Sub

Private Sub SetMailApplication(MailApplication As MAIL_APPLICATION)
  Dim i As Long
  On Error GoTo SetMailApplication_Err
 
  m_MailApplication_new = MailApplication
  
  'Set Mail Application
  For i = optMailApplication.LBound To optMailApplication.UBound + 1
    If StrComp(optMailApplication(i).Tag, MailApplication, vbTextCompare) = 0 Then
      optMailApplication(i).Value = True
      Exit For
    End If
  Next i
  
  If m_MailApplicationChanged Then
    Select Case MailApplication
    Case MA_OTHER
      Call SetMailSystemOptions
'      Call SetMailSystem(IDSM_SYS_UNKNOWN, False)
    Case MA_MICROSOFT_MAIL, MA_MICROSOFT_EXCHANGE, MA_NOVELL_GROUPWISE_V5_PLUS
      Call SetMailSystem(IDSM_SYS_MAPI, True)
    Case MA_MICROSOFT_OUTLOOK
      Call SetMailSystem(IDSM_SYS_MAPI_EXTENDED, True)
    Case MA_LOTUS_CC_MAIL
      Call SetMailSystem(IDSM_SYS_VIM, True)
    Case MA_LOTUS_NOTES_VIM
      Call SetMailSystem(IDSM_SYS_VIM, True)
    Case MA_LOTUS_NOTES_API
      Call SetMailSystem(IDSM_SYS_NOTES, True)
    Case MA_PEGASUS
      Call SetMailSystem(IDSM_SYS_SMTP_POP, True)
    End Select
    m_MailApplicationChanged = False
  Else
    Call SetMailSystem(m_MailSystem_new, True)
  End If
  
  Call RedemptionEnableExtendedMapi
  
  
SetMailApplication_End:
  Exit Sub

SetMailApplication_Err:
  Resume SetMailApplication_End
End Sub


Private Sub cmdDetectSettings_Click()
  Call DetectDefaultSettings
End Sub
Private Sub cmdTestSettings_Click()
 m_Mail.TestMailMode = True
 Call SetupTestForm(Me)
 m_Mail.TestMailMode = False
End Sub

Public Sub SetupTestForm(OwnerForm As Form)
  Dim frm As Form
  If OwnerForm Is Nothing Then
    Set OwnerForm = frmMailSetup
  End If
  sb.SimpleText = "Loading test screen..."
  
  'Use current setup settings for component test
  Call SaveComponentChanges(True)
  Set frm = New frmMailTest
  Set frm.m_Mail = m_Mail
  frm.m_DebugOutputPathFile = m_DebugOutputPathFile

  Call frm.Show(vbModal, OwnerForm)
  If m_Mail.Success Then
    sb.SimpleText = "Successful test completed"
  Else
    sb.SimpleText = "Test Settings not completed/successful"
  End If
  'Use restore original component settings following test
  Call SaveComponentChanges(False)
End Sub

Public Sub DetectDefaultSettings()
  Dim DetectedLoginName As String, DetectedMailSystem As IDSM_MAIL_SYSTEM, DetectedMailApplication As MAIL_APPLICATION
  On Error GoTo DetectDefaultSettings_Err

    sb.SimpleText = "Detecting default settings.."

   'Reset options
    Call SetMailApplication(MA_MICROSOFT_OUTLOOK)
    Call SetMailSystem(IDSM_SYS_MAPI_EXTENDED, True)
    txtLoginName.Text = ""
   
   'Autodetect Native Mail System
    m_Mail.IDSMail.SetNativeSystem
    DetectedMailSystem = m_Mail.IDSMail.MailSystem
    
   'Query for Mail application
    Call QueryMailApplication(m_Mail.IDSMail, DetectedMailSystem, DetectedMailApplication)
  
   'Query Default profile
    Call QueryDefaultProfile(m_Mail.IDSMail, DetectedLoginName)
      
   'Display detected options on screen
    If (DetectedMailApplication = MA_MICROSOFT_OUTLOOK) Then
      DetectedMailSystem = IDSM_SYS_MAPI_EXTENDED
    End If
    
   
    Call SetMailApplication(DetectedMailApplication)
    Call SetMailSystem(DetectedMailSystem)
    m_LoginName_new = DetectedLoginName
    txtLoginName.Text = DetectedLoginName

   'Enable all possible options for Mail System
    Call SetMailSystemOptions

    sb.SimpleText = "Default settings loaded"

DetectDefaultSettings_End:
  Exit Sub

DetectDefaultSettings_Err:
  Err.Raise Err.Number, "DetectDefaultSettings", Err.Description
  Resume DetectDefaultSettings_End
  Resume
End Sub

Private Sub SaveFormChanges(SaveChanges As Boolean)
  If SaveChanges Then
     'Reset form level variables
     m_MailApplication_old = m_MailApplication_new
     m_MailSystem_old = m_MailSystem_new
     m_LoginName_old = m_LoginName_new
  End If
End Sub
Private Sub SaveComponentChanges(SaveChanges As Boolean)
  If SaveChanges Then
     sb.SimpleText = "Saving settings"
     'Save settings back to component
     m_Mail.MailApplication = m_MailApplication_new
     m_Mail.IDSMail.MailSystem = m_MailSystem_new
     m_Mail.IDSMail.LoginName = m_LoginName_new
     
     'RK Save to registry?
     Call m_Mail.RegistrySettings(REGISTRY_KEY_WRITE)
     sb.SimpleText = "Settings saved"
  Else
     'Restore original settings
     m_Mail.MailApplication = m_MailApplication_old
     m_Mail.IDSMail.MailSystem = m_MailSystem_old
     m_Mail.IDSMail.LoginName = m_LoginName_old
     
     Call m_Mail.RegistrySettings(REGISTRY_KEY_WRITE)
  End If
End Sub
Private Function CheckChanges() As Boolean
  Dim bDirty As Boolean
   bDirty = False
     If m_MailApplication_new <> m_MailApplication_old Then bDirty = True
     If m_MailSystem_new <> m_MailSystem_old Then bDirty = True
     If m_LoginName_new <> m_LoginName_old Then bDirty = True
   CheckChanges = bDirty
End Function

Private Sub optMailSystem_Click(Index As Integer)
  Call SetMailSystem(optMailSystem(Index).Tag)
End Sub

Private Sub optMailApplication_Click(Index As Integer)
  If m_MailApplication_new <> optMailApplication(Index).Tag Then
    m_MailApplicationChanged = True
  End If
  Call SetMailApplication(optMailApplication(Index).Tag)
End Sub
Private Sub RedemptionEnableExtendedMapi()
  If m_MailApplication_new = MA_MICROSOFT_OUTLOOK Then
    optMailSystem(3).Enabled = True
    optMailSystem(1).Enabled = True
    optMailSystem(0).Enabled = True
  End If
End Sub

Private Sub SetMailSystemOptions()
  Dim i As Long
  On Error GoTo SetMailSystemOptions_Err
 
  'Enable all available options for mail system
  For i = optMailSystem.LBound To optMailSystem.UBound
    If optMailSystem(i).Tag = IDSM_SYS_UNKNOWN Then
      optMailSystem(i).Enabled = True
    ElseIf m_Mail.IDSMail.QueryMailSystem(optMailSystem(i).Tag) = True Then
      optMailSystem(i).Enabled = True
    Else
      optMailSystem(i).Enabled = False
    End If
  Next i
  
  Call RedemptionEnableExtendedMapi
  
SetMailSystemOptions_End:
  Exit Sub

SetMailSystemOptions_Err:
  Resume SetMailSystemOptions_End
End Sub

'AM Core function
Function GetSysDirectory() As String
  Dim sRes As String
  Dim retval As Long
  
  On Error GoTo GetSysDirectory_err
  sRes = String$(TCSBUFSIZ, 0)
  retval = GetSystemDirectory(sRes, TCSBUFSIZ)
  If retval = 0 Then
    sRes = WINDIR
  Else
    sRes = Left$(sRes, retval)
  End If
  
GetSysDirectory_end:
  GetSysDirectory = sRes
  Exit Function
  
GetSysDirectory_err:
  Resume GetSysDirectory_end
End Function

Private Sub CheckDebugAppPathFile()
  Dim i As Long
  On Error GoTo CheckDebugAppPathFile_Err
  If Not FileExists(m_DebugAppPathFile) Then
    cmdDebug.Enabled = False
    m_DebugAppPathFile = ""
    Err.Raise ERR_DEBUGAPP_NOT_FOUND, "CheckDebugAppPathFile", "IDSMail debug application not found: " & m_DebugAppPathFile
  End If
CheckDebugAppPathFile_End:
  Exit Sub

CheckDebugAppPathFile_Err:
  Err.Raise Err.Number, "CheckDebugAppPathFile", Err.Description
  Resume CheckDebugAppPathFile_End
End Sub

'Public Function ConvertMailApplicationToString(MyMailApplication As MAIL_APPLICATION) As String
'  Dim i As Long
'  For i = optMailApplication.LBound To optMailApplication.UBound
'    If StrComp(optMailApplication(i).Tag, MailApplication, vbTextCompare) = 0 Then
'       ConvertMailApplicationToString = optMailApplication(i).Caption
'      Exit For
'    End If
'  Next i
'
'ConvertMailApplicationToString_end:
'  Exit Function
'
'ConvertMailApplicationToString_err:
'  Resume QueryMailApplication_end
'End Function


Private Sub txtLoginName_Change()

End Sub

Private Sub txtLoginName_Validate(Cancel As Boolean)
  m_LoginName_new = txtLoginName.Text
End Sub
