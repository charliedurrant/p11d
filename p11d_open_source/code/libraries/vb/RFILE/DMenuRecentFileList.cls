VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMenuRecentFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ILibraryVersion
Private m_SetupOK As Boolean

Private m_MaxMenuCaptionStringLength As Long
Private m_IniPathAndFile As String

Private m_vbMenuRecentFile As VBMenu
Private m_vbMenuItemRecentFile As VBMenuItem

Private m_MaxNumberOfFiles As Long
Private m_NoOfRecentFilesAvailable As Long

Private m_IRecentFileList As IRecentFileList
Private m_RecentFiles() As RECENT_FILE

Public Sub Kill()
  Dim i As Long
  
  Set m_IRecentFileList = Nothing
  If m_SetupOK Then
    Call WriteRecentFiles
    Call LoadvbMenuItems(m_vbMenuRecentFile, m_vbMenuItemRecentFile, 0)
  End If
  Set m_vbMenuRecentFile = Nothing
  Set m_vbMenuItemRecentFile = Nothing
  m_SetupOK = False
End Sub

Public Sub RecentFileClick(ByVal MenuIndex As Long)
  Dim mnu As VBMenuItem
  
  On Error GoTo RecentFileClick_ERR
  Call xSet("RecentFileClick")
  Set mnu = m_vbMenuRecentFile.Item(vbMenuName(MenuIndex))
  Select Case m_IRecentFileList.Validate(m_RecentFiles(MenuIndex).CanonicalPathAndFile)
    Case RFL_OK
      Call RegigFileList(MenuIndex)
    Case RFL_DISABLE
      mnu.Enabled = False
    Case RFL_REMOVE
      Call RegigFileList(MenuIndex, True)
    Case Else
      Call ECASE("Invalid Return value from recent files list")
  End Select
  Call WriteRecentFilesToMenu

RecentFileClick_END:
  Call xReturn("RecentFileClick")
  Exit Sub
  
RecentFileClick_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RecentFileClick", "Recent File Click", "Error on clicking recent file Index " & MenuIndex)
  Resume RecentFileClick_END
End Sub

Public Property Let MaxRecentFiles(ByVal NewVal As Long)
  If NewVal > 0 Then
    If Not (m_vbMenuRecentFile Is Nothing) And Not (m_IRecentFileList Is Nothing) Then
      Call WriteRecentFiles
      Call Setup(m_vbMenuRecentFile, m_vbMenuItemRecentFile, m_IRecentFileList, NewVal, m_MaxMenuCaptionStringLength)
    End If
  End If
End Property

Public Sub AddRecentFile(ByVal sPathAndFile As String)
  Dim i As Long
  
  On Error GoTo AddRecentFile_ERR
  Call xSet("AddRecentFile")
  If Len(sPathAndFile) = 0 Then Call Err.Raise(ERR_NO_FILE, "Add Recent File", "No recent file to add.")
  For i = 1 To m_NoOfRecentFilesAvailable
    If (StrComp(m_RecentFiles(i).NonCanonicalPathAndFile, sPathAndFile, vbTextCompare) = 0) Or _
       (StrComp(m_RecentFiles(i).CanonicalPathAndFile, sPathAndFile, vbTextCompare) = 0) Then
      Call RegigFileList(i)
      Exit Sub
    End If
  Next i
  Call RegigFileList(m_MaxNumberOfFiles)
  If m_NoOfRecentFilesAvailable < m_MaxNumberOfFiles Then m_NoOfRecentFilesAvailable = m_NoOfRecentFilesAvailable + 1
  m_RecentFiles(1).NonCanonicalPathAndFile = sPathAndFile
  m_RecentFiles(1).CanonicalPathAndFile = GetCanonicalPathName(sPathAndFile)
  m_RecentFiles(1).Enabled = True
  m_RecentFiles(1).Valid = True
  Call WriteRecentFilesToMenu

AddRecentFile_END:
  Call xReturn("AddRecentFile")
  Exit Sub
AddRecentFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddRecentFile", "Add Recent File", "Error adding recent file.")
  Resume AddRecentFile_END
End Sub

Private Sub Class_Initialize()
  m_SetupOK = False
End Sub

Private Sub Class_Terminate()
  Call Kill
End Sub

Private Function ReadRecentFiles() As Long
  Dim i As Long, rCount As Long
  Dim bNonUNCFileExists As Boolean, bTryUNCFileExists As Boolean, bSucceededUNCFileExist As Boolean
  Dim sIniEntry As String
    
  On Error GoTo ReadRecentFiles_ERR
  Call xSet("ReadRecentFiles")
  bTryUNCFileExists = True
  bSucceededUNCFileExist = False
  ReDim m_RecentFiles(1 To m_MaxNumberOfFiles)
  i = 1: rCount = 1
  Do
    m_RecentFiles(i).Valid = False
    sIniEntry = GetIniEntry(S_INI_SECTION_RECENT_FILES, S_INI_ENTRY_RECENT_FILE & i, "", m_IniPathAndFile)
    If Len(sIniEntry) > 0 Then
      m_RecentFiles(rCount).CanonicalPathAndFile = Left$(sIniEntry, InStr(sIniEntry, ",") - 1)
      m_RecentFiles(rCount).NonCanonicalPathAndFile = Right$(sIniEntry, Len(sIniEntry) - InStr(sIniEntry, ","))
      bNonUNCFileExists = FileExists(m_RecentFiles(rCount).NonCanonicalPathAndFile)
      If bNonUNCFileExists Then
        m_RecentFiles(rCount).Enabled = True
      Else
        If bTryUNCFileExists Then
          m_RecentFiles(rCount).Enabled = FileExists(m_RecentFiles(rCount).CanonicalPathAndFile)
          bSucceededUNCFileExist = bSucceededUNCFileExist Or m_RecentFiles(rCount).Enabled
          bTryUNCFileExists = bSucceededUNCFileExist
        Else
          m_RecentFiles(rCount).Enabled = False
        End If
      End If
      m_RecentFiles(rCount).Valid = True
      rCount = rCount + 1
    End If
    i = i + 1
  Loop Until i > m_MaxNumberOfFiles
  ReadRecentFiles = rCount - 1
  
ReadRecentFiles_END:
  Call xReturn("ReadRecentFiles")
  Exit Function
  
ReadRecentFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReadRecentFiles", "Read Recent Files", "Error reading recent files from " & m_IniPathAndFile)
  Resume ReadRecentFiles_END
End Function

Private Function WriteRecentFiles() As Long
  Dim i As Long
  Dim sIniEntry As String
  
  On Error GoTo WriteRecentFiles_ERR
  Call xSet("WriteRecentFiles")
  For i = 1 To m_MaxNumberOfFiles
    If m_RecentFiles(i).Valid Then
      sIniEntry = m_RecentFiles(i).CanonicalPathAndFile & "," & m_RecentFiles(i).NonCanonicalPathAndFile
      Call WriteIniEntry(S_INI_SECTION_RECENT_FILES, S_INI_ENTRY_RECENT_FILE & i, sIniEntry, m_IniPathAndFile)
    End If
  Next
  
WriteRecentFiles_END:
  Call xReturn("WriteRecentFiles")
  Exit Function
  
WriteRecentFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteRecentFiles", "Write Recent Files", "Error writing recent files.")
  Resume WriteRecentFiles_END
End Function

Private Sub WriteRecentFilesToMenu()
  Dim i As Long
  Dim mnu As VBMenuItem
  
  On Error GoTo WriteRecentFilesToMenu_ERR
  Call xSet("WriteRecentFilesToMenu")
  For i = 1 To m_MaxNumberOfFiles
    Set mnu = m_vbMenuRecentFile.Item(vbMenuName(i))
    If i <= m_NoOfRecentFilesAvailable Then
      mnu.Caption = "&" & i & "  " & GetMenuCaption(m_RecentFiles(i).NonCanonicalPathAndFile, m_MaxMenuCaptionStringLength)
      mnu.Enabled = m_RecentFiles(i).Enabled
      mnu.Visible = True
    Else
      mnu.Visible = False
    End If
  Next
    
WriteRecentFilesToMenu_END:
  Call xReturn("WriteRecentFilesToMenu")
  Exit Sub
  
WriteRecentFilesToMenu_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteRecentFilesToMenu", "Write Recent Files To Menu", "Error writing the recent files to the menu array.")
  Resume WriteRecentFilesToMenu_END
End Sub

Private Sub RegigFileList(ByVal IndexToMove As Long, Optional ByVal RemoveFile As Boolean)
  Dim i As Long
  Dim RF As RECENT_FILE
    
  On Error GoTo RegigFileList_ERR
  Call xSet("RegigFileList")
  If RemoveFile Then
    For i = IndexToMove To m_MaxNumberOfFiles
      m_RecentFiles(i) = m_RecentFiles(i + 1)
    Next i
    m_RecentFiles(m_NoOfRecentFilesAvailable).Valid = False
    m_NoOfRecentFilesAvailable = m_NoOfRecentFilesAvailable - 1
  Else
    If IndexToMove > 1 Then
      RF = m_RecentFiles(IndexToMove)
      For i = IndexToMove To 2 Step -1
        m_RecentFiles(i) = m_RecentFiles(i - 1)
      Next i
      m_RecentFiles(1) = RF
    End If
  End If

RegigFileList_END:
  Call xReturn("RegigFileList")
  Exit Sub
  
RegigFileList_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RegigFileList", "Regig File List", "Error regigging file list.")
  Resume RegigFileList_END
End Sub

Public Sub Setup(ByVal mnuRecentFile As VBMenu, ByVal mnuItemRecentFile As VBMenuItem, ByVal Irfl As IRecentFileList, ByVal MaxNoOfFiles As Long, ByVal MaxMenuCaptionStringLength As Long)
  Dim mnu As VBMenuItem
  On Error GoTo Setup_ERR
  Call xSet("RecentFileListSetup")
  Call RegisterLibrary(Me)
  If mnuRecentFile Is Nothing Then Err.Raise ERR_NO_MENU, "RecentFileListSetup", "No mnuRecentFile entry on menu."
  If Irfl Is Nothing Then Err.Raise ERR_NO_IRFL, "Setup", "No RecentFileList Interface."
  If MaxNoOfFiles < 1 Then Err.Raise ERR_MAX_FILES, "Setup", "Max no of files cannot be less than 1."
  m_MaxMenuCaptionStringLength = MaxMenuCaptionStringLength
  Set m_vbMenuRecentFile = mnuRecentFile
  Set m_vbMenuItemRecentFile = mnuItemRecentFile
  Set m_IRecentFileList = Irfl
  m_MaxNumberOfFiles = MaxNoOfFiles
  m_IniPathAndFile = GetWindowsDirectory() & "\" & AppExeName & ".ini"
  
  Call LoadvbMenuItems(m_vbMenuRecentFile, m_vbMenuItemRecentFile, m_MaxNumberOfFiles)
  m_NoOfRecentFilesAvailable = ReadRecentFiles()
  Call WriteRecentFilesToMenu
  Call mnuRecentFile.Refresh
  m_SetupOK = True
  
Setup_END:
  Call xReturn("RecentFileListSetup")
  Exit Sub
  
Setup_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RecentFileListSetup", "Setup recent file list", "Error setting up recent file list.")
  m_SetupOK = False
  Resume Setup_END
End Sub

Public Property Get MenuPrefix() As String
  MenuPrefix = S_MENU_PREFIX
End Property

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "Recent File List (DMenu)"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

