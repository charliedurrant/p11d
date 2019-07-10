VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerComp 
   Caption         =   "Compare system files"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Caption         =   "View"
      Height          =   1095
      Left            =   45
      TabIndex        =   4
      Top             =   4230
      Width           =   1695
      Begin VB.OptionButton optDisplay 
         Caption         =   "Different"
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   585
         Width           =   1095
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "All"
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   1845
      Top             =   4815
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vercomp.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vercomp.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vercomp.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vercomp.frx":0336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClipCopy 
      Caption         =   "&Copy to clipboard"
      Height          =   420
      Left            =   4560
      TabIndex        =   3
      Top             =   4965
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   6585
      TabIndex        =   2
      Top             =   4965
      Width           =   1275
   End
   Begin VB.ListBox lst 
      Height          =   4155
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvDLLS 
      Height          =   4875
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   8599
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "System Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Loader Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "System date"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Loader date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVerComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type VERSION_DETAILS
  Initialised As Boolean
  vList As ObjectList
End Type
  
Private Enum VERCOMP_SECTIONS
  [_VCS_FIRST_ITEM] = 1
  VCS_ALWAYSINSTALL = [_VCS_FIRST_ITEM]
  'entries here
  VCS_INSTALL
  [_VCS_LAST_ITEM] = VCS_INSTALL
End Enum

Private Enum VERCOMP_DATA
  VCD_DESCRIPTION = 1
  VCD_VERSION_DATA
End Enum

Private Enum ICO_DLL
  ICO_LESS_THAN = 1
  ICO_GREATER_THAN
  ICO_EQUALS
  ICO_QUESTION
End Enum

Private Const S_INI_SECTION_SETTINGS As String = "SETTINGS"
Private m_VersionDetails([_VCS_FIRST_ITEM] To [_VCS_LAST_ITEM]) As VERSION_DETAILS
Private Const SNG_RESIZE_BORDER As Single = 100

Private Function OSSpecificPath(ByVal sPath As String, OS As OS_TYPE) As String
  Select Case OS
    Case OS_WIN95
      OSSpecificPath = sPath & "Win95\"
    Case OS_WIN98, OS_WINME
      OSSpecificPath = sPath & "Win98\"
    Case OS_NT4
      OSSpecificPath = sPath & "NT4\"
    Case OS_WIN2000
      OSSpecificPath = sPath & "Win2000\"
    Case OS_WINXP
      OSSpecificPath = sPath & "WinXP\"
    Case OS_WIN2003
      OSSpecificPath = sPath & "Win2003\"
    Case Else
      ECASE_SYS ("Unknown OS type in OSSpecificPath")
  End Select
End Function

Public Sub Initialise()
  Dim i As VERCOMP_SECTIONS
  Dim v As Variant
  
  On Error GoTo Initialise_ERR
  Set lvDLLS.SmallIcons = iml
  For i = [_VCS_FIRST_ITEM] To [_VCS_LAST_ITEM]
    m_VersionDetails(i).Initialised = False
    v = GetData(i, VCD_DESCRIPTION)
    lst.AddItem v
    lst.ItemData(lst.ListCount - 1) = i
  Next i
  If lst.ListCount > 1 Then lst.ListIndex = 0
  Me.Show vbModal
  
Initialise_END:
  Exit Sub
  
Initialise_ERR:
  Resume Initialise_ERR
End Sub

Private Function FillVerList(ByVal lv As ListView, ByVal VCS As VERCOMP_SECTIONS) As String
  Dim li As ListItem
  Dim i, j As Long
  Dim VD As VersionData, sVerlist As String, sState As String
  
  On Error GoTo FillVerList_ERR
  If Not lv Is Nothing Then lv.ListItems.Clear
  Call SetCursorEx(vbHourglass, "")
  If Not m_VersionDetails(VCS).Initialised Then
    Set m_VersionDetails(VCS).vList = GetData(VCS, VCD_VERSION_DATA)
    m_VersionDetails(VCS).Initialised = True
  End If
  sVerlist = "State" & vbTab & "Name" & vbTab & "System Version" & vbTab & "Loader Version" & vbTab & "System datetime" & vbTab & "Loader datetime" & vbCrLf
  If Not m_VersionDetails(VCS).vList Is Nothing Then
    For j = 1 To m_VersionDetails(VCS).vList.Count
      Set VD = m_VersionDetails(VCS).vList(j)
      If (VD.CompareVersions = 0) And VD.LoaderFileExists And VD.SystemFileExists And (optDisplay(1).Value = True) Then GoTo NEXT_ITEM
      If Not lv Is Nothing Then
        Set li = lv.ListItems.Add(, , VD.DLLName)
        li.SubItems(1) = VD.CurrentVersion
        li.SubItems(2) = VD.LoaderVersion
        li.SubItems(3) = VD.CurrentDateTime
        li.SubItems(4) = VD.LoaderDateTime
      End If
      sState = "?"
      If VD.LoaderFileExists = False Or VD.SystemFileExists = 0 Then
        If Not lv Is Nothing Then li.SmallIcon = ICO_QUESTION
      Else
        Select Case VD.CompareVersions
          Case -1
            If Not lv Is Nothing Then li.SmallIcon = ICO_LESS_THAN
            sState = "<"
          Case 1
            If Not lv Is Nothing Then li.SmallIcon = ICO_GREATER_THAN
            sState = ">"
          Case 0
            If Not lv Is Nothing Then li.SmallIcon = ICO_EQUALS
            sState = "="
          Case Else
            ECASE_SYS ("Versioning for " & VD.DLLName & " is invalid.")
        End Select
      End If
      sVerlist = sVerlist & sState & vbTab & VD.DLLName & vbTab & VD.CurrentVersion & vbTab & VD.LoaderVersion & vbTab & VD.CurrentDateTime & vbTab & VD.LoaderDateTime & vbCrLf
NEXT_ITEM:
     Next j
   End If
   
FillVerList_END:
  FillVerList = sVerlist
  Call ClearCursorEx(True)
  Exit Function
  
FillVerList_ERR:
  Resume FillVerList_END
End Function

Private Property Get LoaderDLLPath(ByVal sIniSection As String, ByVal sDefaultValue) As String
  LoaderDLLPath = FullPathEx(mHomeDirectory & GetIniEntryEx(S_INI_SECTION_SETTINGS, sIniSection, sDefaultValue, LoaderFilePath))
End Property

Private Function Drive(ByVal sPath As String) As String
  Dim i As Long
  
  If Len(sPath) > 2 Then
    If StrComp(left$(sPath, 2), "\\") = 0 Then
      Drive = left(sPath, InStr(3, sPath, "\", vbBinaryCompare) - 1)       'canonical
    Else
      Drive = left(sPath, 2)
    End If
  Else
    ECASE_SYS ("Drive can not be obtained from " & sPath & ".")
  End If
End Function

'apf cd to fix
Private Property Get MSDaoPath() As String
  Dim sDir As String
  
  Call SplitPathEx(GetWindowsDirectoryEx, sDir, "", "")
  MSDaoPath = Drive(GetWindowsDirectoryEx) & "\Program files\common files\microsoft shared\dao\"
End Property

Private Function GetData(ByVal VCS As VERCOMP_SECTIONS, ByVal vcd As VERCOMP_DATA) As Variant
  Dim v As Variant
  Dim VerData As ObjectList
  Dim s As String
  Dim i As Long, j As Long
  Dim VD As VersionData
  Dim sWinSysPath As String
  Dim sLoaderDLLPath As String
    
  On Error GoTo GetData_ERR
  If vcd = VCD_VERSION_DATA Then Set VerData = New ObjectList
  sWinSysPath = FullPathEx(GetSysDirectoryEx)
  
  Select Case VCS
    Case VCS_INSTALL
      Select Case vcd
        Case VCD_DESCRIPTION
          GetData = "System components"
        Case VCD_VERSION_DATA
          sLoaderDLLPath = LoaderDLLPath("InstallPath", "Setup\System")
          For i = 1 To GetIniKeyNamesExInternal(v, "INSTALL", LoaderFilePath())
            Set VD = New VersionData
            VD.DLLName = v(i)
            'get system files
            s = GetIniEntryEx("INSTALL", v(i), "", LoaderFilePath)
            If Len(s) = 0 Then GoTo NEXT_ITEM
            If InStr(1, s, "WinSysPath", vbTextCompare) > 0 Then
              VD.DLLPathSystem = sWinSysPath
            ElseIf InStr(1, s, "MSDAOPath", vbTextCompare) > 0 Then
              VD.DLLPathSystem = MSDaoPath
            Else
              ECASE_SYS ("Unknown dll path in file" & LoaderFilePath() & " entry " & s & ".")
            End If
            
            VD.SystemFileExists = FileExistsEx(VD.DLLPathSystem & VD.DLLName, False, False)
            
            If Not FileExistsEx(sLoaderDLLPath & VD.DLLName, False, False) Then
              'then try for NT, Win95,Win98 specific
              Call FileExistsVersionSpecific(VD, sLoaderDLLPath)
            Else
              VD.LoaderFileExists = True
              VD.DLLPathLoader = sLoaderDLLPath
            End If
            Call VerData.Add(VD)
NEXT_ITEM:
          Next
      End Select
      
    Case VCS_ALWAYSINSTALL
      Select Case vcd
        Case VCD_DESCRIPTION
          GetData = "Always install"
        Case VCD_VERSION_DATA
          sLoaderDLLPath = LoaderDLLPath("AlwaysPath", "Setup")
          s = GetIniEntryEx("ALWAYSINSTALL", "dlls", "", LoaderFilePath)
          For i = 1 To GetDelimitedValuesEx(v, s, True, True, ";", """")
            Set VD = New VersionData
            VD.DLLPathSystem = sWinSysPath
            VD.DLLName = v(i)
            VD.DLLPathLoader = sLoaderDLLPath
            VD.LoaderFileExists = FileExistsEx(VD.DLLPathLoader & VD.DLLName, False, False)
            VD.SystemFileExists = FileExistsEx(VD.DLLPathSystem & VD.DLLName, False, False)
            Call VerData.Add(VD)
          Next
      End Select
  End Select
    
  If vcd = VCD_VERSION_DATA Then
    For i = 1 To VerData.Count
      Set VD = VerData(i)
      Call GetVersionData(VD)
    Next
    Set GetData = VerData
  End If
    
GetData_END:
  Exit Function
GetData_ERR:
  Resume GetData_END
End Function

Private Sub FileExistsVersionSpecific(VD As VersionData, ByVal sLoaderPath As String)
  Dim lMajor As Long, lMinor As Long, OS As OS_TYPE, lBuild As Long, s As String
  
  Call GetWindowsVersion(lMajor, lMinor, lBuild, OS, s)
  VD.LoaderFileExists = FileExistsEx(OSSpecificPath(sLoaderPath, OS) & VD.DLLName, False, False)
  If VD.LoaderFileExists Then VD.DLLPathLoader = OSSpecificPath(sLoaderPath, OS)
End Sub

Private Sub GetVersionData(VD As VersionData)
  Dim sPathAndFile As String
  Dim sPropertyResult As String
  
  On Error GoTo GetVersionData_ERR
  If VD Is Nothing Then
    Call ECASE_SYS("VD is nothing in GetVersionData")
    GoTo GetVersionData_END
  End If
  
  If (VD.LoaderFileExists Or VD.SystemFileExists) = False Then GoTo GetVersionData_END
  
  sPathAndFile = VD.DLLPathSystem & VD.DLLName
  
  sPropertyResult = VersionQueryMap(sPathAndFile, VQT_FILE_VERSION)
  VD.CurrentVersion = sPropertyResult
  VD.CurrentDateTime = FileDateTime(sPathAndFile)
  
  VD.CompareVersions = VerCompEx(sPathAndFile, VD.DLLPathLoader & VD.DLLName)
  sPathAndFile = VD.DLLPathLoader & VD.DLLName
  
  sPropertyResult = VersionQueryMap(sPathAndFile, VQT_FILE_VERSION)
  VD.LoaderVersion = sPropertyResult
  VD.LoaderDateTime = FileDateTime(sPathAndFile)
  
GetVersionData_END:
  Exit Sub
  
GetVersionData_ERR:
  Resume GetVersionData_END
End Sub

Private Property Get LoaderFilePath() As String
  LoaderFilePath = mHomeDirectory & mAppExeName & ".LOD"
End Property

Private Sub cmdClipCopy_Click()
  Dim sVerlist As String
  
  sVerlist = FillVerList(Nothing, lst.ItemData(lst.ListIndex))
  If OpenClipboard(0) Then
    Call EmptyClipboard
    Call SetAnyClipboardDataEx(vbCFText, sVerlist)
    Call CloseClipboard
  End If
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Resize()
  If (Me.ScaleHeight < 2700) Or (Me.ScaleWidth < 500) Then Exit Sub
  lst.top = SNG_RESIZE_BORDER
  lst.left = SNG_RESIZE_BORDER
  fra.left = SNG_RESIZE_BORDER
  lst.Height = Me.ScaleHeight - (SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + fra.Height)
  fra.top = lst.top + lst.Height + SNG_RESIZE_BORDER
  lvDLLS.left = lst.left + lst.Width + SNG_RESIZE_BORDER
  lvDLLS.Width = Me.ScaleWidth - (SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + lst.Width)
  
  lvDLLS.top = SNG_RESIZE_BORDER
  lvDLLS.Height = Me.ScaleHeight - (SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + cmdOK.Height)
  
  cmdOK.top = lvDLLS.top + lvDLLS.Height + SNG_RESIZE_BORDER
  cmdOK.left = Me.ScaleWidth - (SNG_RESIZE_BORDER + SNG_RESIZE_BORDER + cmdOK.Width)
  
  cmdClipCopy.Height = cmdOK.Height
  cmdClipCopy.top = cmdOK.top
  cmdClipCopy.left = cmdOK.left - cmdClipCopy.Width - SNG_RESIZE_BORDER
End Sub

Private Sub lst_Click()
  If lst.ListIndex <> -1 Then Call FillVerList(lvDLLS, lst.ItemData(lst.ListIndex))
End Sub

Private Sub optDisplay_Click(Index As Integer)
  Call lst_Click
End Sub
