VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'* Sets the cursor, default is the hourglass
'*
'* return value:
'* none
Public Sub SetCursor(Optional lCursor As MousePointerConstants = vbHourglass)
  Call gTracer.xSet("SetCursor")
  mpstack(curmp) = Screen.MousePointer
  curmp = curmp + 1
  Screen.MousePointer = lCursor
  Call gTracer.xReturn("SetCursor")
End Sub

'* Clears the current cursor and restores the prior cursor on the cursor stack
'*
'* return value:
'* none
Public Sub ClearCursor()
  Call gTracer.xSet("ClearCursor")
  If curmp > 0 Then
    curmp = curmp - 1
    Screen.MousePointer = mpstack(curmp)
  Else
    Screen.MousePointer = vbDefault
  End If
  Call gTracer.xReturn("ClearCursor")
End Sub

'* WinExec an application and show error messages
'* if no path is specified windows will search the current dir, windows directory, system directory and the application path
'*
'* sFile              [path] & file name of program
'*
'* return value:
'* double representing the program's task ID if successful, otherwise it returns zero
Public Function ShellProgram(sFile As String, Optional winstyle As VbAppWinStyle = vbNormalFocus, Optional Wait As Boolean = False) As Long
  Dim pid As Long, hproc As Long, hErr As Long, ExitCode As Long
  
  On Error GoTo shellprogram_err
  Call gTracer.xSet("ShellProgram")
  pid = CLng(Shell(sFile, winstyle))
  If Wait Then
    hproc = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    If hproc <> 0 Then
      Do
        hErr = GetExitCodeProcess(hproc, ExitCode)
        Call Sleep(100)
      Loop Until (ExitCode <> STILL_ACTIVE) Or (hErr = 0)
    End If
  End If
  ShellProgram = pid
  
ShellProgram_end:
  Call gTracer.xReturn("ShellProgram")
  Exit Function

shellprogram_err:
  Call gCore.ErrorMessage(ERR_ERROR, "ShellProgram", "ERR_WINEXEC", "Failed to winexec " & sFile)
  Resume ShellProgram_end
End Function

' apf check
'* Write an entry to the ini file specified or of the current application ([AppExeName].INI) if none is specified
'* Careful here with empty strings for sSection and sKey
'* If this is empty then the whole section in the ini file is deleted
'* If sKey is empty then sSection is deleted
'*
'* ssection         section - i.e. [STARTUP]
'* sKey             Key - i.e. Load=
'* sString          Value to be written
'* [sPathAndFile]   Path and file name of ini file (Default = [Windows directory]\[APPEXENAME].INI)
'*
'* return value:
'*                  True if success
Public Function WriteIniEntry(sSection As String, sKey As String, sString As String, Optional sPathAndFile = "") As Boolean
  Dim sFname As String
  Dim l As Long
  
  Call gTracer.xSet("WriteIniEntry")
  If sPathAndFile = "" Then
    sFname = GetWindowsDirectory() & "\" & gCore.AppExeName & ".INI"
  Else
    sFname = sPathAndFile
  End If
  l = WritePrivateProfileString(sSection, sKey, sString, sFname)
  WriteIniEntry = (l <> 0&)
  Call gTracer.xReturn("WriteIniEntry")
End Function

'* Read an entry to the ini file specified or of the current application ([AppExeName].INI)
'*
'* sSection             section in ini file
'* sKey                 key
'* [sDef]               default value for the key (Default = "")
'* [sPathAndFile]       Path and filename of ini file (Default = [Windows directory]\[APPEXENAME].INI)
'*
'* Return value
'*                      The value for the key in the ini file, or sDef if the Key is not
'*                      present
'*
'* note: reads from an ini file with entries formatted as
'* [section]
'* key=value
'*
Public Function GetIniEntry(sSection As String, sKey As String, Optional ByVal sDef As String = "", Optional sPathAndFile = "") As String
  Dim sFname As String
  Dim sBuffer As String
  Dim retval As Long
  
  Call gTracer.xSet("GetIniEntry")
  If Len(sPathAndFile) = 0 Then
    sFname = GetWindowsDirectory() & "\" & gCore.AppExeName & ".INI"
  Else
    sFname = sPathAndFile
  End If
  sBuffer = String$(bufsiz, 0)
  retval = GetPrivateProfileString(sSection, sKey, "", sBuffer, bufsiz, sFname)
  If retval = 0 Then
    sBuffer = sDef
  Else
    sBuffer = gString.RTrimChar(sBuffer, vbNullChar)
  End If
  GetIniEntry = sBuffer
  Call gTracer.xReturn("GetIniEntry")
End Function

Public Function GetWinIniEntry(sSection As String, sKey As String, Optional ByVal sDef As String = "") As String
  Dim sFname As String
  Dim sBuffer As String
  Dim retval As Long
  
  Call gTracer.xSet("GetWinIniEntry")
  sBuffer = String$(bufsiz, 0)
  retval = GetProfileString(sSection, sKey, sDef, sBuffer, bufsiz)
  If retval = 0 Then
    sBuffer = ""
  Else
    sBuffer = gString.RTrimChar(sBuffer, vbNullChar)
  End If
  GetWinIniEntry = sBuffer
  Call gTracer.xReturn("GetWinIniEntry")
End Function

'* Returns true if password is entered correctly
'*
'*
'* return value
'* true password is entered correctly
Public Function GetTCSPassword() As Boolean
  Static lPassword As Long
  Dim sToday As String
  
  Call gTracer.xSet("GetTCSPassword")
  GetTCSPassword = False
  sToday = Format$(Now, "dd mmmm yyyy")
  'Initialise the dialog box
  With frmPassw
    .txtPassword = CStr(lPassword)
    .txtPassword.SelStart = 0
    .txtPassword.SelLength = Len(.txtPassword.Text)
    .lblInfoDate = "Please enter the TCS password for " & sToday & "."
    .lblContact = gCore.GetStatic("contact")
  End With
  
  'Show password dialog box until correct password or
  'the dialog box is cancelled
  Do Until GetTCSPassword = True
    frmPassw.Show vbModal
    If (frmPassw.PasswordOk) Then
      'Check password is correct
      If IsNumeric(frmPassw.txtPassword.Text) Then
        lPassword = CLng(frmPassw.txtPassword.Text)
      Else
        lPassword = -1
      End If
      If lPassword = GetPassword() Then
        GetTCSPassword = True
      Else
        'If password is incorrect then show info box and then show
        'password dialog again
        Call gCore.ErrorMessage(ERR_INFO, "GetTCSPassword", "ERR_INVALID PASSWORD", "The password you have entered is invalid")
        frmPassw.txtPassword.SelStart = 0
        frmPassw.txtPassword.SelLength = Len(frmPassw.txtPassword.Text)
      End If
    Else
      Exit Do
    End If
  Loop
  Unload frmPassw
  Set frmPassw = Nothing
  Call gTracer.xReturn("GetTCSPassword")
End Function

'* get the Password for a specific date
Public Function GetPassword(Optional ddate As Date = UNDATED) As Long
  Call gTracer.xSet("GetPassword")
  If ddate = UNDATED Then ddate = Now
  GetPassword = Abs(2 ^ (WeekDay(ddate, vbMonday)) + IIf(Day(ddate) Mod 2, -1, 1) * (Day(ddate) - (Len(Format(ddate, "dddd")) * Len(Format(ddate, "mmmm")))))
  Call gTracer.xReturn("GetPassword")
End Function

Public Sub SystemStub()
  Dim l As Long
  l = 5
End Sub

'* Show the TCS debug popup with a password check
'*
'* return value: true if succesful
Public Function ShowDebugPopup() As Boolean
  Call frmErr.PopupMenu(frmErr.mnuDebug)
  ShowDebugPopup = True
End Function

'* GetCanonicalPathName
'* doc please
Public Function GetCanonicalPathName(ByVal sPath) As String
  Dim sBuffer As String, sFileName As String
  Dim lBufSiz As Long, retval As Long

On Error GoTo GetCanonicalPathName_err
  Call gTracer.xSet("GetCanonicalPathName")
  lBufSiz = 1024
  sBuffer = String$(lBufSiz, 0)
  sFileName = String$(lBufSiz, 0)
  
  retval = GetFullPathName(sPath, lBufSiz, sBuffer, sFileName)
  If retval <> 0 Then
    GetCanonicalPathName = gString.RTrimChar(sBuffer, vbNullChar)
  End If
  
GetCanonicalPathName_end:
  Call gTracer.xReturn("GetCanonicalPathName")
  Exit Function
  
GetCanonicalPathName_err:
  Resume GetCanonicalPathName_end
End Function

'* Get the free and total disk space of the drive specified
'*
'* sRootPath           Root Path of drive
'* lTotal              Total space on this drive
'* lFree               Free space on drive
'*
'* return value
'* true if successful
Public Function GetDiskSpace(sRootPath As String, fTotal As Single, fFree As Single, Optional lMem As MemoryUnit = BYTES) As Boolean
  Dim lSectorsPerCluster As Long, lBytesPerSector As Long, lNumberOfFreeClusters As Long, lTotalNumberOfClusters As Long
  Dim retval As Long
  
  gTracer.xSet ("GetDiskSpace")
  GetDiskSpace = False
  retval = GetDiskFreeSpace(sRootPath, lSectorsPerCluster, lBytesPerSector, lNumberOfFreeClusters, lTotalNumberOfClusters)
  If retval Then
    fTotal = CSng(lTotalNumberOfClusters) * CSng(lSectorsPerCluster) * CSng(lBytesPerSector)
    fFree = CSng(lNumberOfFreeClusters) * CSng(lSectorsPerCluster) * CSng(lBytesPerSector)
    Select Case lMem
    Case GIGABYTES
      fTotal = fTotal / 1.073742E+09!
      fFree = fFree / 1.073742E+09!
    Case MEGABYTES
      fTotal = fTotal / 1048576!
      fFree = fFree / 1048576!
    End Select
    GetDiskSpace = True
  End If
  gTracer.xReturn ("GetDiskSpace")
End Function

'* Display the (windows) standard folder browse dialog
'*
'* hWndOwner           Handle to parent window (normally mdimain.hwnd)
'* sPath               String for path to be stored in
'* [sCaption]          Dialog box caption (default = "Select a folder")
'*
'* return value
'* true if successful
Public Function lBrowseFolder(ByVal hWndOwner As Long, sPath As String, Optional ByVal sPrompt As String = "Select a folder") As Long
Dim lpIDList As Long, lResult As Long
Dim udtBI As BrowseInfo
Dim sFN As String
  On Error GoTo lBrowseFolder_ERR
  sFN = "lBrowseFolder"
  gTracer.xSet (sFN)
  With udtBI
    .hWndOwner = hWndOwner
    .lpszTitle = sPrompt
    '1 = BIF_RETURNONLYDIRS
    .ulFlags = 1
  End With
  
  lpIDList = SHBrowseForFolder(udtBI)
  If lpIDList Then
    sPath = String$(bufsiz, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    If lResult Then
      Call CoTaskMemFree(lpIDList)
      sPath = gString.RTrimChar(sPath, Chr(0))
      lBrowseFolder = True
    Else
      lBrowseFolder = False
    End If
  Else
    lBrowseFolder = False
  End If
  
lBrowseFolder_END:
  gTracer.xReturn (sFN)
  Exit Function
lBrowseFolder_ERR:
  lBrowseFolder = False
  Call gCore.ErrorMessage(ERR_ERROR, sFN, "Error " & Err.LastDllError, Err.Description)
  Resume lBrowseFolder_END
End Function

'* Display the (windows) standard FileOpen Dialog box
'*
'*
'* return value
Public Function FileOpenDlg(ByVal sTitle As String, ByVal sFilter As String, ByVal sInitDir As String, Optional MultiSelect As Boolean = False) As String
  Dim cd As CommonDialog
  
  On Error GoTo FileOpenDlg_Err
  Call gTracer.xSet("FileOpenDlg")
  Set cd = frmAbout.cd
  'cdlOFNAllowMultiselect
  'cdlOFNFileMustExist
  'cdlOFNHideReadOnly
  'cdlOFNLongNames
  'cdlOFNNoDereferenceLinks
  cd.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
  If MultiSelect Then cd.Flags = cd.Flags + cdlOFNAllowMultiselect
  cd.DialogTitle = sTitle
  cd.Filter = sFilter
  cd.InitDir = sInitDir
  cd.CancelError = True
  cd.ShowOpen
  FileOpenDlg = cd.filename
  
FileOpenDlg_End:
  Call gTracer.xReturn("FileOpenDlg")
  Exit Function

FileOpenDlg_Err:
  FileOpenDlg = ""
  Resume FileOpenDlg_End
End Function

'*   Places the network user name into a string.
'*
'*  sUser         String to place User name in
'*  [lBufSize]    Default to buffer size - size of data to retrieve
'*
'*  return value
'*  True if succeeds, 0 if code fails, if error with network
'*  returns the dll error (see DLLErrors enumeration)
'*  ERROR_NOT_CONNECTED
'*  ERROR_MORE_DATA
'*  ERROR_NO_NETWORK
'*  ERROR_EXTENDED_ERROR
' note: apf/rmc - do we need to enum valid network connections??
Function GetNetUser(Optional bShowErrors As Boolean = False) As String
  Dim lRetVal As Long, lBufSiz As Long
  Dim sUser As String
    
  On Error GoTo GetNetUser_err
  Call gTracer.xSet("GetNetUser")
  lBufSiz = bufsiz
  sUser = String$(lBufSiz, 0)
  lRetVal = WNetGetUser("", sUser, lBufSiz)
  If lRetVal = 0 Then
    sUser = gString.RTrimChar(sUser, vbNullChar)
    GetNetUser = sUser
  Else
    GetNetUser = ""
  End If
  
GetNetUser_end:
  Call gTracer.xReturn("GetNetUser")
  Exit Function
  
GetNetUser_err:
  GetNetUser = ""
  If bShowErrors Then
    Call gCore.ErrorMessage(ERR_ERROR, "GetNetUser", "ERR_GETNETUSER", "")
  End If
  Resume GetNetUser_end
End Function
  
'* retrieves current windows directory (e.g C:\WINDOWS)
Function GetWindowsDirectory() As String
  Dim sRes As String
  Dim retval As Long
  On Error GoTo GetWindowsDirectory_err
  
  gTracer.xSet ("GetWindowsDirectory")
  sRes = String$(bufsiz, 0)
  retval = GetWindowsDirectory32(sRes, bufsiz)
  If retval = 0 Then
    sRes = WINDIR
  Else
    sRes = gString.RTrimChar(sRes, vbNullChar)
  End If
  
GetWindowsDirectory_end:
  GetWindowsDirectory = sRes
  gTracer.xReturn ("GetWindowsDirectory")
  Exit Function
GetWindowsDirectory_err:
  Call gCore.ErrorMessage(ERR_ERROR, "GetWindowsDirectory", "ERR_GETWINDOWSDIRECTORY", "Unable to retrieve Windows directory." & vbCr & "using " & WINDIR)
  Resume GetWindowsDirectory_end
End Function
  
'*  Given a drive letter return either the network resource and path (lrestype = 0) or just network resource (lrestype =1)
'*
'*  sdrv          String to place drive letter in & :
'*  sRes          String where result is placed
'*  [lrestype]    See above default = 0
'*  [lBufSiz]     if function returns ERR_MORE_DATA then increase size of buffer
'*
'*  return value
'*  True if succeeds, 0 if code fails, if error with network or resource does not exist
'*  returns the dll error (see DLLErrors enumeration)
Function lGetNetDrvRes&(sDrv$, sRes$, Optional lResType& = 0, Optional lBufSize = bufsiz)
Dim lRetVal&
Dim sFN$
  On Error GoTo lGetNetDrvRes_ERR
  sFN = "lGetNetDrvMap"
  gTracer.xSet (sFN)
  
  sRes = String$(lBufSize, 0)
  lRetVal = WNetGetConnection(sDrv, sRes, lBufSize - 1)
  If lRetVal = 0 Then
    lGetNetDrvRes = True
    sRes = gString.RTrimChar(sRes, Chr(0))
    If lResType = 1 Then
    sRes = gString.LTrimChar(sRes, "\")
    lRetVal = InStr(sRes, "\")
    If lRetVal Then
      sRes = Left(sRes, lRetVal - 1)
    End If
  End If
  Else
    lGetNetDrvRes = Err.LastDllError
  End If
  
lGetNetDrvRes_END:
  gTracer.xReturn (sFN)
  Exit Function
lGetNetDrvRes_ERR:
  lGetNetDrvRes = False
  Call gCore.ErrorMessage(ERR_ERROR, sFN, "Error " & Err.Number, Err.Description)
  Resume lGetNetDrvRes_END
End Function

'* Locks updates
Function LockWindowUpdate(Optional ByVal hWnd As Long = 0) As Boolean
  Static winlocked As Long
  Dim l As Long
  
  If hWnd <> 0 Then
    If winlocked = 0 Then
      l = LockWindowUpdateW32(hWnd)
    End If
    winlocked = winlocked + 1
  Else
    winlocked = winlocked - 1
    If winlocked <= 0 Then
      winlocked = 0
      l = LockWindowUpdateW32(hWnd)
    End If
  End If
  LockWindowUpdate = (l <> 0)
End Function

Public Function MultiDialog(ByVal sTitle As String, ByVal sMessage As String, ParamArray buttons() As Variant) As Long
  Dim fm As Form
  Dim cmb As CommandButton
  Dim l As Long, lDelta As Long
  
  Set fm = frmMessage
  For l = LBound(buttons) To UBound(buttons)
    Load fm.cmdArray(l)
    Set cmb = fm.cmdArray(l)
    With cmb
      .Top = fm.fmeMessage.Top + 100 + (l * (cmb.Height + 100))
      .Caption = buttons(l)
      .Visible = True
    End With
  Next l
  fm.Caption = sTitle
  fm.lblMessage.Caption = sMessage
  fm.Show 1
  MultiDialog = fm.lClicked
  Set fm = Nothing
End Function

Public Sub SetText(lbl As Object, sText As String)
  
  On Error GoTo SetText_err
  Call gTracer.xSet("SetText")
  If TypeOf lbl Is Label Then
    lbl.Caption = sText
  ElseIf TypeOf lbl Is TextBox Then
    lbl.Text = sText
  Else
    Call gCore.ECASE("SetText - only valid on textboxes and labels")
    Exit Sub
  End If
  lbl.Refresh
  
SetText_end:
  Call gTracer.xReturn("SetText")
  Exit Sub
  
SetText_err:
  Resume SetText_end
End Sub

'* True for success, False (+ raised error otherwise)
Public Function PopUpButtons(hWndToolBarOCX As Long, Optional bTopDivider As Boolean = False) As Boolean
  Dim hWndToolbarDLL As Long, lStyle As Long
  Dim hErr As Long
  
  Call gTracer.xSet("PopUpButtons")
  
  lStyle = TBSTYLE_FLAT
  If bTopDivider = 0 Then
    lStyle = lStyle Or CCS_NODIVIDER
  End If
  
  hWndToolbarDLL = FindWindowEx(hWndToolBarOCX, 0, "ToolbarWindow32", vbNullString)
  If hWndToolbarDLL Then
    Call SendMessage(hWndToolbarDLL, TB_SETSTYLE, 0, ByVal lStyle)
    PopUpButtons = True
  Else
    hErr = Err.LastDllError
    Err.Raise ERR_FINDWINDOW, "PopUpButtons", "Unable to Find Window ToolbarWindow32" & vbCrLf & "LastDLL Error: " & CStr(hErr)
  End If
   
PopUpButtons_end:
  Call gTracer.xReturn("PopUpButtons")
  Exit Function
End Function

'* Flush the ini buffer
'*
'* [sPathAndFile]               Name of ini file to flush default = windows directory\exename.ini
'*
'* return value:
'* True if success
Public Function FlushIniBuffer(Optional sPathAndFile As String) As Boolean
  Dim l As Long, sFname As String
  Call gTracer.xSet("FlushIniBuffer")
  If sPathAndFile = "" Then
    sFname = GetWindowsDirectory() & "\" & gCore.AppExeName & ".INI"
  Else
    sFname = sPathAndFile
  End If
  l = WritePrivateProfileString("", "", "", sFname)
  FlushIniBuffer = (l <> 0&)
  Call gTracer.xReturn("FlushIniBuffer")
End Function

