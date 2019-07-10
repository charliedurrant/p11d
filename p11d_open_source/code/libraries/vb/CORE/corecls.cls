VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoreClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Class level constants
'* GLOBAL_LBCOLCHARS          ?????
'* BUFSIZ                     Default buffer size
'* TCSCORE_ERROR              TCSCORE error numbers start at TCSCORE_ERROR
'* TCSCLIENT_ERROR            Your user application error numbers start at TCSCLIENT_ERROR

Public Enum DLLErrors
  ERROR_NOT_CONNECTED = 2250&
  ERROR_MORE_DATA = 234&
  ERROR_NO_NETWORK = 1222&
  ERROR_EXTENDED_ERROR = 1208&
  ERROR_BAD_DEVICE = 1200&
  ERROR_CONNECTION_UNAVAIL = 1201&
  ERROR_NO_NET_OR_BAD_PATH = 1203&
End Enum

Public Enum COM_TYPE
  WIN32_INPROC = 1
  WIN32_SERVERPROC = 2
  WIN32_ANY = 3
  WIN16_SERVERPROC = 4
  WIN16_ANY = 4
  WIN_ANY = 7
End Enum

'Error type constants
'*  For use with ErrorMessage
'*  ERR_ERROR                 Throw a TCS error
'*  ERR_ERRORSILENT           Log the error but do not display the message
'*  ERR_INFO                  Information box
'*  ERR_INFOSILENT            Log the info but do not show box
'*  ERR_NOERROR               No error
'*  Error messages are logged to APPPATH\APPEXENAME.ERR
'*  Info messages are logged to APPPATH\APPEXENAME.LOG
Public Enum errnumbersenum
  ERR_ERROR = 1
  ERR_ERRORSILENT
  ERR_INFO = 1024
  ERR_INFOSILENT
  ERR_ALLOWIGNORE = 2048
  ERR_ALLOWRETRY = 4096
  ERR_ALLOWCANCEL = 8192
  ERR_ALLOWOTHER = 16384
  ERR_NOERROR = -1
End Enum

'*  Return Disk / Memory sizes in the following
'*  BYTES
'*  MEGABYTES
'*  GIGABYTES
Public Enum MemoryUnit
  BYTES
  MEGABYTES
  GIGABYTES
End Enum

Public Enum TCSMenuItems
  MNU_NOTHING
  MNU_BREAK
  MNU_SYSTEM
  MNU_DATABASE
  MNU_APPLICATION
End Enum

Public Enum DATECONVERT_TYPE ' note: see VTEXT if changed
  CONVERT_DATETYPEUNKNOWN = 0
  CONVERT_FIXEDDATE = 256
  CONVERT_DELIMITED
  CONVERT_FIXEDDATETIME
End Enum

Public Enum DATABASE_TARGET
  DB_TARGET_JET = 0
  DB_TARGET_SQLSERVER
  DB_TARGET_ORACLE
End Enum

Public Enum VER_QUERY_TYPE
  VQT_PRODUCT_VERSION = 1
  [_VQT_FIRST_ITEM] = VQT_PRODUCT_VERSION
  VQT_PRODUCT_NAME
  VQT_COMPANY_NAME
  VQT_FILE_DESCRIPTION
  VQT_FILE_VERSION
  VQT_INTERNAL_NAME
  VQT_LEGAL_COPYRIGHT
  VQT_ORIGINAL_FILE_NAME
  VQT_COMMENTS
  VQT_LEGAL_TRADEMARKS
  VQT_PRIVATE_BUILD
  'insert items here
  VQT_SPECIAL_BUILD
  [_VQT_LAST_ITEM] = VQT_SPECIAL_BUILD
End Enum

Public Enum BUTTON_CONSTANTS
  VK_LBUTTON = &H1
  VK_RBUTTON = &H2
End Enum


Public Enum CSIDL_SPECIAL_FOLDERS
  CSIDL_DESKTOP = &H0 '// The Desktop - virtual folder
  CSIDL_PROGRAMS = 2 '// Program Files
  CSIDL_CONTROLS = 3 '// Control Panel - virtual folder
  CSIDL_PRINTERS = 4 '// Printers - virtual folder
  CSIDL_DOCUMENTS = 5 '// My Documents
  CSIDL_FAVORITES = 6 '// Favourites
  CSIDL_STARTUP = 7 '// Startup Folder
  CSIDL_RECENT = 8 '// Recent Documents
  CSIDL_SENDTO = 9 '// Send To Folder
  CSIDL_BITBUCKET = 10 '// Recycle Bin - virtual folder
  CSIDL_STARTMENU = 11 '// Start Menu
  CSIDL_DESKTOPFOLDER = 16 '// Desktop folder
  CSIDL_DRIVES = 17 '// My Computer - virtual folder
  CSIDL_NETWORK = 18 '// Network Neighbourhood - virtual folder
  CSIDL_NETHOOD = 19 '// NetHood Folder
  CSIDL_FONTS = 20 '// Fonts folder
  CSIDL_SHELLNEW = 21 '// ShellNew folder
End Enum

Private Const fileattrs As Long = vbNormal + vbReadOnly + vbArchive
Private Const sysfileattrs As Long = fileattrs + vbHidden + vbSystem
Private Const dirattrs As Long = vbDirectory
Private mClassTerminate As Boolean
Private mCoreFirst As Boolean

Private Sub Class_Initialize()
  mCoreFirst = False
  If mCoreInitCount = 0 Then
    Set gCore = Me
    Set gTCSEventClass = New TCSEventClass
    Set gMenusCollection = New Collection
    Set gstaticdata = New Collection
    Set LibraryVersions = New Collection
  End If
  mCoreInitCount = mCoreInitCount + 1
End Sub

' terminate lib
Private Sub Class_Terminate()
  mCoreInitCount = mCoreInitCount - 1
  mClassTerminate = True
  Call CoreShutDown
  mClassTerminate = False
End Sub

'##D Saves the TCS core DLL statics. No return value.
'##C CLF
Public Sub CoreShutDown()
  On Error GoTo coreshutdown_err
   
  'Call MsgBox("CoreShutDown TERM: " & CStr(mClassTerminate) & " Usage count: " & CStr(mCoreInitCount))
  If (mCoreInitCount = 0) Or ((Not mClassTerminate) And mCoreFirst) Then
    If Not mCoreShutDownDone Then
      mCoreShutDownDone = True
      Call TimerFrequency(False)
      Call KillAllTimers
      Call KillKeyboardHook
      Call SaveStatics
      Call Tracer_report
      Call Tracer_Cleanup(True)
    End If
  End If
  
  If mCoreInitCount = 0 Then
    Set gTCSEventClass = Nothing
    Set gMenusCollection = Nothing
    Set gstaticdata = Nothing
    Set LibraryVersions = Nothing
    Set gCore = Nothing
    Set vbg = Nothing
    Set m_IPreErrorFilter = Nothing
    Set m_IPostErrorProcess = Nothing
    Set m_IDebugMenu = Nothing
    Set m_INotifyTCSPassword = Nothing 'cad
  End If
  
coreshutdown_exit:
  Exit Sub

coreshutdown_err:
  Call MsgBox("ERR_CORESHUTDOWN", vbCritical + vbOKOnly, "Error shutting down the Core Library" & vbCr & contactstr)
  Resume coreshutdown_exit
End Sub

'##D Internal Property used by TCSDB to decide whether to create a default workspace
'##D Set during CoreSetup
'##C CLF
Public Property Get InitialiseDefaultWS() As Boolean
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "InitialiseDefaultWS", "Cannot access TCSCORE properties before calling CoreSetup"
  InitialiseDefaultWS = mTCS_InitialiseDefaultWS
End Property

'##D Setup the TCS Core DLL
'##LD Invokes the tcs core dll and logs the fact that the DLL has
'##LD been initialised.
'##V CommandLine Use command$ for this parameter
'##C    CLF
'##RV   True is succesful
Public Function CoreSetup(ByVal CommandLine As String, ByVal AppVBG As IUnknown, Optional ByVal DatabaseAccess As Boolean = True) As Boolean
  #If DEBUGVER Then
    Call Tracer_XSet("CoreSetup")
  #End If
    If ghInstance = 0 Then mCoreFirst = True
    CoreSetup = CoreSetupEx(CommandLine, AppVBG, DatabaseAccess, False, mCoreFirst)
  #If DEBUGVER Then
    Call Tracer_XReturn("CoreSetup")
  #End If
End Function

'##D Setup the TCS Core DLL from another DLL (only required if a DLL is to be called from a non VB Standard App)
'##LD Invokes the tcs core dll and logs the fact that the DLL has been initialised.
'##V CommandLine Command line sent to setup core
'##C    CLF
'##RV True is succesful
Public Function DLLSetup(ByVal CommandLine As String, Optional ByVal DatabaseAccess As Boolean = True) As Boolean
  #If DEBUGVER Then
    Call Tracer_XSet("DLLSetup")
  #End If
    If ghInstance = 0 Then mCoreFirst = True
    DLLSetup = CoreSetupEx(CommandLine, Nothing, DatabaseAccess, True, mCoreFirst)
  #If DEBUGVER Then
    Call Tracer_XReturn("DLLSetup")
  #End If
End Function

'##D adds a string variable to the common static list maintained by the program, retrieved using getstatic
'##StaticName Name of the static to save
'##V DefaultValue default value (default="")
'##V OverrideValue override value (default="")
'##V Persistent is this a persistent static (saved in msg file) (default=true)
'##LD This is like the MSG file in Abacus - there is one copy per program and it is
'##LD called APPNAME.MSG
'##LD By default when a variable is change
'##C SF
Public Sub AddStatic(ByVal StaticName As String, Optional ByVal DefaultValue As String = "", Optional OverrideValue As Variant, Optional ByVal Persistent As Boolean = True)
'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  Call AddStaticEx(StaticName, DefaultValue, OverrideValue, Persistent)
End Sub

'##D retrives a string variable from the common static list maintained by the program
'##V StaticName name of static to retrieve
'##RV Value of the static
'##C SF
Public Function GetStatic(ByVal StaticName As String) As Variant
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  GetStatic = GetStaticEx(StaticName, True)
End Function

Public Property Get MsgFileName() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  MsgFileName = mStaticFileName
End Property

Public Property Get HomeDirectory() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  HomeDirectory = mHomeDirectory
End Property

Public Property Let HomeDirectory(ByVal NewValue As String)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  If Not FileExistsEx(NewValue, True, False) Then Err.Raise 380, "HomeDirectory", "Home directory '" & LCase$(NewValue) & "' does not exist"
  mHomeDirectory = FullPathEx(NewValue)
  mStaticFileName = mHomeDirectory & mAppExeName & ".MSG"
End Property

Public Property Let FormattedErrorStrings(ByVal NewValue As Boolean)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  mFormattedErrorStrings = NewValue
End Property

'##D saves static list
'##C SF
Public Sub SaveStatics()
  Dim sItem As staticcls, ret As Boolean
  #If DEBUGVER Then
    Call logfunction3(Now, 0, "SAVE_STATICS", "Start Save Statics", "SaveStatics", LogFileExt)
  #End If
  For Each sItem In gstaticdata
    If sItem.bPersist Then
      ret = WriteIniEntry(STATICS_SECTION, sItem.Name, CStr(sItem.Value), mStaticFileName)
      #If DEBUGVER Then
        If Not ret Then Call logfunction3(Now, 0, "SAVE_STATICS", "Failed to save static: " & sItem.Name, "SaveStatics", LogFileExt)
      #End If
    End If
  Next
  #If DEBUGVER Then
    Call logfunction3(Now, 0, "SAVE_STATICS", "End Save Statics", "SaveStatics", LogFileExt)
  #End If
End Sub


'##D Main Error Message Function - does not allow cascading errors
'##V ErrType Error type (ERR_INFO, ERR_ERROR etc)
'##V ErrObj Err object, can be nothing for information messages
'##V FunctionName Name of function calling errormessage
'##V ErrTitle Error Title
'##V ErrMessage Message String
'##RV Returns true if Retry/Cancel button choosen
'##C EMF
Public Function ErrorMessage(ByVal ErrType As errnumbersenum, ErrObj As ErrObject, ByVal FunctionName As String, ByVal ErrTitle As String, ByVal ErrMessage As String) As Boolean
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  ErrorMessage = ErrorMessageEx(ErrType, ErrObj, FunctionName, ErrTitle, ErrMessage, False)
End Function

'##D Errormessage function for error messages recorded with ErrorMessagePush
'##C EMF
Public Function ErrorMessagePop(ByVal ErrType As errnumbersenum, ErrObj As ErrObject, ByVal FunctionName As String, ByVal ErrTitle As String, ByVal ErrMessage As String) As Boolean
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  ErrorMessagePop = ErrorMessageEx(ErrType, ErrObj, FunctionName, ErrTitle, ErrMessage, True)
End Function

'##D Errormessage function to repopulate the Err object for error messages recorded with ErrorMessagePush
'##C EMF
Public Sub ErrorMessagePopErr(ErrObj As ErrObject)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  Call PopErrorMessageErr(ErrObj)
End Sub

'##D Function to build the Source string of an errormessage
'##C EMF
Public Function ErrorSource(ErrObj As ErrObject, ByVal FunctionName As String) As String
  ErrorSource = ErrorSourceEx(ErrObj, FunctionName)
End Function

'##D Records the current error whcih can then be displyed using ErrorMessagePop
'##LD This is required where your error handler calls functions which have error handlers themselves (On Error handler clears the current error information)
'##C EMF
Public Sub ErrorMessagePush(ErrObj As ErrObject)
  Call PushErrorMessage(ErrObj)
End Sub


'##D Quick error raising function. Use for debugging code, eg. case else ECASE ("Item not found")
'##V SystemMessage Error Message
'##C EMF
Public Sub ECASE(ByVal SystemMessage As String)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  Call ECASE_SYS(SystemMessage)
End Sub

'##D Same as ECASE, however only displays message if the program is running in the IDE
'##V SystemMessage Error Message
'##C EMF
Public Sub ECASE_DEBUG(ByVal SystemMessage As String)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  Call ECASE_SYS(SystemMessage, True)
End Sub

'##D Gets DOS path of program
'##RV None
'##C FNF
Public Property Get AppPathDos() As String
  Static sAppPathDos As String
  
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  If Not Len(sAppPathDos) Then
    sAppPathDos = GetDrivePathEx(mAppPath)
  End If
  AppPathDos = sAppPathDos
End Property

'##D Gets path of program
'##RV None
'##C FNF
Public Property Get AppPath() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppPath = mAppPath
End Property

'##D Application name
'##C SQF
Public Property Get AppName() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppName = mAppName
End Property

'##D Application executable name
'##C SQF
Public Property Get AppExeName() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppExeName = mAppExeName
End Property

'##D Application command parameter string
'##C SQF
Public Property Get AppCmdParam() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppCmdParam = mAppCmdParam
End Property

'##D TCS Core dll version number
'##C SQF
Public Property Get TCSCoreVersion() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  TCSCoreVersion = mTCSCoreVersion
End Property

Public Property Get InErrorMessage() As Boolean
  InErrorMessage = inErrMainMsg
End Property

'##D Prevents error message box from appearing but logs the error
'##C EMF
Public Property Get SilentError() As Boolean
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  SilentError = mSilentError
End Property

Public Property Let SilentError(ByVal NewValue As Boolean)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  mSilentError = NewValue
End Property

Public Property Let ErrorOtherButtonCaption(ByVal NewValue As String)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  m_OtherCaption = NewValue
End Property

Public Property Let RetryCancelCaption(ByVal NewValue As String)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  m_OtherCaption = NewValue
End Property

'##D Obsolete use DatabaseTarget instead
'##C DF
Public Property Get SQLServerStrings() As Boolean
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  SQLServerStrings = (mDBTarget = DB_TARGET_SQLSERVER)
End Property

Public Property Let SQLServerStrings(ByVal NewValue As Boolean)
  Call ECASE("SQLServerStrings property obsolete: Use DatabaseTarget property instead." & vbCrLf)
End Property

'##D Sets/Gets the current database target (SQL/Access/Oracle)
'##LD Setting the value of this property changes the behaviour of Database and SQL functions
'##C DF
Public Property Get DatabaseTarget() As DATABASE_TARGET
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  DatabaseTarget = mDBTarget
End Property

Public Property Let DatabaseTarget(ByVal NewValue As DATABASE_TARGET)
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  mDBTarget = NewValue
End Property

'##D Checks for the existence of a file or directory, wildcards accepted
'##FileName path and (filename)
'##V Directory look for directories (default = false)
'##V ShowErrors report/log errors (default = false)
'##RV True if file/directory exists
Public Function FileExists(ByVal FileName As String, Optional ByVal Directory As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("FileExists")
#End If
  FileExists = FileExistsEx(FileName, Directory, ShowErrors)
#If DEBUGVER Then
   Call Tracer_XReturn("FileExists")
#End If
End Function

'##D Makes Directory recursively
'##V Directory Name of Directory
'##RV True if directory created or already present
'##C FNF
Public Function xMkdir(ByVal Directory As String) As Boolean
  Dim DirString As String
  Dim p0 As Long, p1 As Long
  
  On Error GoTo xMkdir_err
  If left$(Directory, 2) = "\\" Then
    p0 = 3
    p0 = InStr(p0, Directory, "\")
    If p0 = 0 Then Err.Raise ERR_MKDIR, "xMkdir", "Unable to create directory " & Directory
    p0 = InStr(p0 + 1, Directory, "\")
    If p0 = 0 Then Err.Raise ERR_MKDIR, "xMkdir", "Unable to create directory " & Directory
    p0 = p0 + 1
  Else
    p0 = 4
  End If
  Do
    p1 = InStr(p0, Directory, "\")
    If p1 = 0 Then
      DirString = Mid$(Directory, 1)
    Else
      DirString = Mid$(Directory, 1, p1 - 1)
      p0 = p1 + 1
    End If
    Call MkDir(DirString)
  Loop Until p1 = 0
  xMkdir = True
  
xMkdir_end:
  Exit Function
  
xMkdir_err:
  xMkdir = False
  If Err.Number = 75 Then Resume Next
  Resume xMkdir_end
End Function

'##D Get DOS Path of a Directory/Path
'##RV The DOS path (includes driver letter)
'##C FNF
Public Function GetDOSPath(ByVal DirectoryOnly As String) As String
  GetDOSPath = GetDrivePathEx(DirectoryOnly)
End Function

Public Function GetDOSFullPath(ByVal FullPath As String) As String
  Dim sDir As String, sFile As String, sExt As String
  
  Call SplitPathEx(FullPath, sDir, sFile, sExt)
  GetDOSFullPath = GetDrivePathEx(sDir) & sFile & sExt
End Function

'##D xKill wraps Kill function - erases file
'##V FullPath Full Path of file to be erased
'##RV False if failed to kill files, no error is raised if no files erased due to no files matching mask
'##C FNF
Public Function xKill(ByVal FullPath As String) As Boolean
  xKill = xKillEx(FullPath)
End Function

'##D Count the number of files in directory, does not test for existence of path
'##V FullPath the path of the directory to be searched
'##RV Number of files
'##C FNF
Public Function CountFiles(ByVal FullPath As String) As Long
  Dim sFile As String
#If DEBUGVER Then
  Call Tracer_XSet("CountFiles")
#End If
  If FileExistsEx(FullPath, True, False) Then FullPath = FullPathEx(FullPath)
  sFile = Dir$(FullPath)
  Do While Len(sFile) > 0
    CountFiles = CountFiles + 1
    sFile = Dir
  Loop
#If DEBUGVER Then
  Call Tracer_XReturn("CountFiles")
#End If
End Function

Public Function FindFiles(ByVal FileDirectory As String, ByVal FileMask As String, Optional ByVal SubDirs As Boolean = False) As StringList
  Set FindFiles = FindFilesEx(FileDirectory, FileMask, SubDirs, False)
End Function

Public Function FindFiles2(ByVal FileDirectory As String, ByVal FileMask As String, Optional ByVal SubDirs As Boolean = False, Optional ByVal IncludeDirectories As Boolean = False) As StringList
  Set FindFiles2 = FindFilesEx(FileDirectory, FileMask, SubDirs, IncludeDirectories)
End Function

Public Function isDir(ByVal FileName As String) As Boolean
  isDir = FileExistsEx(FileName, True, False)
End Function

'##D Centre a form to the screen
'##V FormToCentre Form to be centred
'##RV none
'##C FDF
Public Sub CentreForm(FormToCentre As Object)
  Call CentreInForm(FormToCentre, Nothing)
End Sub

'##D Centre a form within another form
'##V FormToCentre Form to be centred
'##V CentreOnForm Form on which FormToCentre is to be centred
'##C FDF
Public Sub CentreInForm(FormToCentre As Object, CentreOnForm As Object)
  Call CentreInFormEx(FormToCentre, CentreOnForm)
End Sub

'##D Invoke the standard TCS about box
'##RV None
'##C SQF
Public Function AppAbout() As Boolean

  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  frmAbout.Caption = "About " & mAppName
  frmAbout.lblApplication = mAppName
  frmAbout.lblAppVersion = "Version " & mAppVersion
  frmAbout.lblCopyright = "Copyright " & Chr$(169) & " 1998-" & Year(Now)
  Call SetHelpAboutText(m_HelpAboutText)
  frmAbout.Show vbModal
  Set frmAbout = Nothing
End Function

'##D Returns true if a command line parameter is used and optionally the parameter specified
'##V Param Command line parameter to search for or parameter number (1 based)
'##V Buffer String in which the value is returned
'##V vCmd Command line string to be searched
'##RV True if command parameter found, false otherwise
'##C SQF
Public Function GetCmdParam(ByVal Param As Variant, Optional buffer As Variant, Optional vCmd As Variant) As Boolean
  #If DEBUGVER Then
    Call Tracer_XSet("GetCmdParam")
  #End If
    'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
    GetCmdParam = GetCmdParamEx(Param, buffer, vCmd)
  #If DEBUGVER Then
    Call Tracer_XReturn("GetCmdParam")
  #End If
End Function

'##D Show the TCS debug popup with a password check
'##RV true if successful
'##C TCSMF
Public Function ShowDebugPopup() As Boolean
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  ShowDebugPopup = ShowDebugPopupex
End Function


'##D Add menu item to debug menu
'##V Item              Caption of Item (must be unique)
'##V ParentMenu          Parent menu of the item (enumerated) default = MNU_APPLICATION
'##V Enabled         Is the menu item enabled
'##RV true if succesful
'##C TCSMF
Public Function AddMenuItem(ByVal Name As String, ByVal Item As String, Optional ByVal ParentMenu As TCSMenuItems = MNU_APPLICATION, Optional ByVal Enabled As Boolean = True) As Boolean
  Dim m As MenuClass
  
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  On Error GoTo AddMenuItem_err
  Set m = New MenuClass
  m.Name = Name
  m.Caption = Item
  m.Parent = ParentMenu
  m.Enabled = Enabled
  Call gMenusCollection.Add(m, Name)
  AddMenuItem = True
  
AddMenuItem_end:
  Exit Function
  
AddMenuItem_err:
  AddMenuItem = False
  Resume AddMenuItem_end
End Function

'##D Update menu item in debug menu - you must provide the enabled state, you can change the parent
'##V Item Caption of Item
'##V Enabled Enabled property - NOT OPTIONAL (also not very useful!!) !
'##V Parentmenu Parent menu of the item (enumerated) default = MNU_NOTHING
'##RV True if succesful
'##C TCSMF
Public Function UpdateMenuItem(ByVal Name As String, Optional ByVal Item As String = "", Optional ByVal ParentMenu As TCSMenuItems = MNU_NOTHING, Optional ByVal Enabled As Boolean = False) As Boolean
  Dim m As MenuClass
  
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  On Error GoTo UpdateMenuItem_err
  Set m = gMenusCollection.Item(Name)
  If Len(Item) <> 0 Then m.Caption = Item
  If ParentMenu <> MNU_NOTHING Then m.Parent = ParentMenu
  m.Enabled = Enabled
  UpdateMenuItem = True
  
UpdateMenuItem_end:
  Exit Function
  
UpdateMenuItem_err:
  UpdateMenuItem = False
  Resume UpdateMenuItem_end
End Function

'##D Remove menu item from debug menu
'##V Name Caption of Item (must be unique)
'##RV true if successful
Public Function RemoveMenuItem(ByVal Name As String) As Boolean

  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  On Error GoTo RemoveMenuItem_err
  
  gMenusCollection.Remove (Name)
  RemoveMenuItem = True
  
RemoveMenuItem_end:
  Exit Function
  
RemoveMenuItem_err:
  RemoveMenuItem = False
  Resume RemoveMenuItem_end
End Function

'*  The global reference to the TCS event class
'*
Public Property Get gEvents() As TCSEventClass
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  Set gEvents = gTCSEventClass
End Property

'##D Returns Minimum Value in an Sequence of Parameters
'##C NF
Public Function Min(ParamArray vals() As Variant) As Variant
  Dim i As Long
  
  Min = vals(LBound(vals))
  For i = (LBound(vals) + 1) To UBound(vals)
    If vals(i) < Min Then Min = vals(i)
  Next i
End Function

'##D Returns Maximum Value in a Sequence of Parameters
'##C NF
Public Function Max(ParamArray vals() As Variant) As Variant
  Dim i As Long
  
  Max = vals(LBound(vals))
  For i = (LBound(vals) + 1) To UBound(vals)
    If vals(i) > Max Then Max = vals(i)
  Next i
End Function

'##D Truncates a number to an integer
'##d Number to be truncated
'##C NF
Public Function Trunc(ByVal d As Double) As Long
  Trunc = CLng(Int(d + DBL_EPSILON))
End Function


'##D Set the function timer and place this function on the call stack
'##V FunctionName      Function name
'##RV None
'##C MF
Public Sub xSet(FunctionName As String)
#If DEBUGVER Then
  Call Tracer_XSet(FunctionName)
#End If
End Sub

'##D Unset the function timer and remove last function on the call stack
'##V FunctionName      Function name
'##RV none
'##C MF
Public Sub xReturn(FunctionName As String)
#If DEBUGVER Then
  Call Tracer_XReturn(FunctionName)
#End If
End Sub

'##D Modifies a date by changing any specific field
'##V ChangeDate           Date to be modified
'##V Day               Overide day (Default -1)
'##V Month              Overide month (Default -1)
'##V Year               Overide year (Default -1)
'##V Hour              Overide hours (Default -1)
'##V Minute             Overide minutes (Default -1)
'##V Seconds            Overide seconds (Default -1)
'##LD Passing a -1 (the default value) means leaves that portion of the date unchanged
'##RV The changed date
'##C DTF
Public Function SetDateTime(ByVal ChangeDate As Date, Optional ByVal Day As Integer = -1, Optional ByVal Month As Integer = -1, Optional ByVal Year As Integer = -1, Optional ByVal Hour As Integer = 0, Optional ByVal Minute As Integer = 0, Optional ByVal Seconds As Integer = 0) As Date
  Dim nday0 As Integer, nmonth0 As Integer, nyear0 As Integer
  Dim nhour0 As Integer, nminute0 As Integer, nsecond0 As Integer
  
#If DEBUGVER Then
  Call Tracer_XSet("SetDateTime")
#End If
  ' date
  If Day <> -1 Then
    nday0 = Day
  Else
    nday0 = DatePart("d", ChangeDate)
  End If
  If Month <> -1 Then
    nmonth0 = Month
  Else
    nmonth0 = DatePart("m", ChangeDate)
  End If
  If Year <> -1 Then
    nyear0 = Year
  Else
    nyear0 = DatePart("yyyy", ChangeDate)
  End If
  
  ' time
  If Hour <> -1 Then
    nhour0 = Hour
  Else
    nhour0 = DatePart("h", ChangeDate)
  End If
  If Minute <> -1 Then
    nminute0 = Minute
  Else
    nminute0 = DatePart("n", ChangeDate)
  End If
  If Seconds <> -1 Then
    nsecond0 = Seconds
  Else
    nsecond0 = DatePart("s", ChangeDate)
  End If

  SetDateTime = DateSerialEx(nyear0, nmonth0, nday0) + TimeSerial(nhour0, nminute0, nsecond0)
#If DEBUGVER Then
  Call Tracer_XReturn("SetDateTime")
#End If
End Function

'##D Attempt to convert a string of the form DD/MM/YY HH:NN:SS to a date
'##LD If the string is not a valid date then returns DefaultDate
'##C DTF
Public Function TryConvertDate(ByVal DateString As Variant, Optional ByVal DefaultDate As Date = UNDATED) As Date
  On Error GoTo TryConvertDate_err
  
#If DEBUGVER Then
  Call Tracer_XSet("TryConvertDate")
#End If

  TryConvertDate = ConvertDateEx(DateString, CONVERT_DELIMITED, "DMY", "/", ":")
  
TryConvertDate_end:
#If DEBUGVER Then
  Call Tracer_XReturn("TryConvertDate")
#End If
  Exit Function
  
TryConvertDate_err:
  TryConvertDate = DefaultDate
  Resume TryConvertDate_end
End Function

'##D Attempt to convert a string of the form DD/MM/YY HH:NN:SS to a date
'##LD If the string is not a valid date then returns DefaultDate if present, or DateString if DefaultDate is not specified
'##C DTF
' note: see VTEXT if changed
Public Function TryConvertDateDMY(ByVal DateString As Variant, Optional ByVal DefaultDate As Variant) As Variant
  On Error GoTo TryConvertDateDMY_err
  
#If DEBUGVER Then
  Call Tracer_XSet("TryConvertDateDMY")
#End If
    
  TryConvertDateDMY = ConvertDateEx(DateString, CONVERT_DELIMITED, "DMY", "/", ":")
  
TryConvertDateDMY_end:
#If DEBUGVER Then
  Call Tracer_XReturn("TryConvertDateDMY")
#End If
  Exit Function
  
TryConvertDateDMY_err:
  If IsMissing(DefaultDate) Then
    TryConvertDateDMY = DateString
  Else
    TryConvertDateDMY = DefaultDate
  End If
  Resume TryConvertDateDMY_end
End Function

'##D Converts Fixed/Delimited date string to dates
'##V ConvStr order DMY or MDY (or DMYHNS) for delimited dates, mask for fixed dates
'##C DTF
Public Function ConvertDate(ByVal DateString As String, ByVal ConvType As DATECONVERT_TYPE, ConvStr As String, Optional ByVal DateDelimit As String = "/", Optional ByVal TimeDelimit As String = ":") As Date
  ConvertDate = ConvertDateEx(DateString, ConvType, ConvStr, DateDelimit, TimeDelimit)
End Function
     
'##D remove multiple occurrances of characters in a string e.g "XX  X  Y" -> "XX X Y"
'##V String1 String to be compressed
'##V Char Character to be removed when multiply occurring
'##RV Compressed String
'##C STF
Public Function CompressString(ByVal String1 As String, ByVal Char As String) As String
  CompressString = CompressStringEx(String1, Char)
End Function

'##D Construct a new string by using the insertion point within a string
'##V string1 string to be manipulated
'##V Insert insertion point in that string
'##V char ASCII character in question
'##LD Construct a new string by using the insertion point within a string
'##LD For use when typing to detect BACKSPACE, DEL etc
'##RV true if text inserted at end or False otherwise
'##C STF
Public Function ConstructString(String1 As String, ByVal Insert As Integer, ByVal Char As Integer) As Boolean
  Dim i As Integer

#If DEBUGVER Then
  Call Tracer_XSet("ConstructString")
#End If
  String1 = ConstructStringEx(String1, ToUpper(Char), Insert, 0)
  ConstructString = ((Len(String1) - 1) = Insert)
#If DEBUGVER Then
  Call Tracer_XReturn("ConstructString")
#End If
End Function

'##C STF
Public Function ConstructString2(ByVal CurrentString As String, ByVal Char As Integer, ByVal SelStart As Integer, ByVal SelLength As Integer) As String
  ConstructString2 = ConstructStringEx(CurrentString, Char, SelStart, SelLength)
End Function

'##D Prepares a string to use in SQL
'##V vString source string
'##RV Valid SQL string
'##C DF
Public Function StrSQL(ByVal vString As Variant) As String
#If DEBUGVER Then
  Call Tracer_XSet("StrSQL")
#End If
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  StrSQL = StrSQLEx(vString)
#If DEBUGVER Then
  Call Tracer_XReturn("StrSQL")
#End If
End Function

'##D Prepares a date to use in SQL
'##RV Valid SQL date as a string
'##C DF
Public Function DateSQL(ByVal vDate As Variant) As String
#If DEBUGVER Then
  Call Tracer_XSet("DateSQL")
#End If
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  If IsNull(vDate) Then
    DateSQL = "Null"
  Else
    If mDBTarget = DB_TARGET_JET Then
      DateSQL = "#" & Format$(vDate, "MM/DD/YYYY") & "#"
    ElseIf mDBTarget = DB_TARGET_SQLSERVER Then
      DateSQL = "'" & Format$(vDate, "YYYYMMDD") & "'"
    ElseIf mDBTarget = DB_TARGET_ORACLE Then
      DateSQL = "TO_DATE('" & Format$(vDate, "YYYYMMDD") & "','YYYYMMDD')"
    End If
  End If
#If DEBUGVER Then
  Call Tracer_XReturn("DateSQL")
#End If
End Function

'##D Prepares a date time to use in SQL
'##RV Valid SQL date as a string
'##C DF
Public Function DateTimeSQL(ByVal vDateTime As Variant) As String
#If DEBUGVER Then
  Call Tracer_XSet("DateTimeSQL")
#End If
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  If IsNull(vDateTime) Then
    DateTimeSQL = "Null"
  Else
    If mDBTarget = DB_TARGET_JET Or mDBTarget = DB_TARGET_SQLSERVER Then
      DateTimeSQL = "'" & Format$(vDateTime, "YYYYMMDD HH:NN:SS") & "'"
    ElseIf mDBTarget = DB_TARGET_ORACLE Then
      DateTimeSQL = "TO_DATE('" & Format$(vDateTime, "YYYYMMDD HH:NN:SS") & "','YYYYMMDD HH24:MI:SS')"
    End If
  End If
#If DEBUGVER Then
  Call Tracer_XReturn("DateTimeSQL")
#End If
End Function

'##C DF
Public Function NumSQL(ByVal vNumber As Variant) As String
#If DEBUGVER Then
  Call Tracer_XSet("NumSQL")
#End If
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  If IsNull(vNumber) Then
    NumSQL = "Null"
  Else
    If VarType(vNumber) = vbBoolean Then
      If vNumber Then
        NumSQL = "-1"
      Else
        NumSQL = "0"
      End If
    Else
      NumSQL = CStr(vNumber)
    End If
  End If
#If DEBUGVER Then
  Call Tracer_XReturn("NumSQL")
#End If
End Function

'##D Takes a Variant converts into either a Date String of the format specified or the DefaultValue if is is not a date
'##C DTF
Public Function DateString2(ByVal v As Variant, Optional ByVal DefaultValue As String = "", Optional ByVal FormatString As String = "DD/MM/YYYY") As String
#If DEBUGVER Then
  Call Tracer_XSet("DateString2")
#End If
  'If ghInstance = 0 Then Err.Raise ERR_INITCORE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  DateString2 = DateStringEx2(v, DefaultValue, FormatString)
#If DEBUGVER Then
  Call Tracer_XReturn("DateString2")
#End If
End Function


'##D Takes a Variant converts into either a Date String of the form DD/MM/YYYY or the Empty String if is is not a date
'##C DTF
Public Function DateString(ByVal v As Variant) As String
#If DEBUGVER Then
  Call Tracer_XSet("DateString")
#End If
  DateString = DateStringEx2(v, "", "DD/MM/YYYY")
#If DEBUGVER Then
  Call Tracer_XReturn("DateString")
#End If
End Function

'##D Takes a variant converts into either a Date String of the form DD/MM/YYYY or the DefaultValue if is is not a date
'##C DTF
Public Function DateStringEx(ByVal v As Variant, ByVal DefaultValue As String) As String
#If DEBUGVER Then
  Call Tracer_XSet("DateStringEx")
#End If
  DateStringEx = DateStringEx2(v, DefaultValue, "DD/MM/YYYY")
#If DEBUGVER Then
  Call Tracer_XReturn("DateStringEx")
#End If
End Function

'##D Duplicate a single character in a string - use StrSQL for SQL strings
'##RV the string with the duplicated character
'##C STF
Public Function StrDupChar(ByVal String1 As String, ByVal CharDup As String) As String
#If DEBUGVER Then
  Call Tracer_XSet("StrDupChar")
#End If
  StrDupChar = StrDupCharEx(String1, CharDup)
#If DEBUGVER Then
  Call Tracer_XReturn("StrDupChar")
#End If
End Function

'##D Find a string in an array return offset of match
'##C STF
Public Function FindString(StrArray() As String, ByVal String1 As String, Optional ByVal NotFoundReturn As Long = -1) As Long
  Dim ub As Long, lb As Long, j As Long
  
  ub = UBound(StrArray)
  lb = LBound(StrArray)
  For j = lb To ub
    If StrComp(StrArray(j), String1, vbTextCompare) = 0 Then
      FindString = j
      Exit Function
    End If
  Next j
  FindString = NotFoundReturn
End Function


'##D Pad a string at front or back with a character to a specified length
'##V String1 original string
'##V Pad padding character
'##V Length length of padded string
'##V PadFront pad from front (default = false)
'##RV The padded string
'##C STF
Public Function xStrPad(ByVal String1 As String, ByVal Pad As String, ByVal Length As Long, Optional ByVal PadFront As Boolean = False) As String
#If DEBUGVER Then
  Call Tracer_XSet("xStrPad")
#End If
  xStrPad = xStrPadEx(String1, Pad, Length, PadFront)
#If DEBUGVER Then
  Call Tracer_XReturn("xStrPad")
#End If
End Function

'##D Same as InStr but starts from the end of the string
'##LD DO NOT USE, USE InStrRev instead !!
'##C STF
Public Function InStrBack(String1 As String, ByVal String2 As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
  Call ECASE("InStrBack obsolete - use InStrRev")
End Function

'##D Counts the occurences of a character / string in a string
'##V String1 string to be searched
'##V Char character to be found in string
'##V Compare Compare type
'##RV Number of occurences of that string
'##C STF
Public Function CountChar(String1 As String, ByVal Char As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
  Dim p0 As Long, p1 As Long, ilen As Long
  
#If DEBUGVER Then
  Call Tracer_XSet("CountChar")
#End If
  ilen = Len(Char)
  If ilen > 0 Then
    p0 = 1
    Do
      p1 = InStr(p0, String1, Char, Compare)
      If p1 > 0 Then
        CountChar = CountChar + 1
        p0 = p1 + ilen
      End If
    Loop Until p1 = 0
  End If
#If DEBUGVER Then
  Call Tracer_XReturn("CountChar")
#End If
End Function

'##D Extracts fields from a delimited string
'##V String1         string to extract FROM
'##V Char            character delimiter
'##V Fields          parameter fields to extract to in order
'##RV Fields in Parameter array Fields
'##C STF
Public Sub xStrExtract(String1 As String, ByVal Char As String, ParamArray Fields() As Variant)
  Dim tchar As String
  Dim i As Long
  Dim iStartPos As Long
  Dim iEndPos As Integer
  Dim sText As String
  
#If DEBUGVER Then
  Call Tracer_XSet("xStrExtract")
#End If
  tchar = left$(Char, 1)
  
  iStartPos = 1
  For i = LBound(Fields) To UBound(Fields)
    If iStartPos > Len(String1) Then GoTo endfunc
    iEndPos = InStr(iStartPos, String1, Char, vbTextCompare)
    If iEndPos = 0 Then
      Fields(i) = right$(String1, Len(String1) - iStartPos + 1)
      Exit For
    Else
      Fields(i) = Mid$(String1, iStartPos, iEndPos - iStartPos)
      iStartPos = iEndPos + 1
    End If
  Next i

endfunc:
#If DEBUGVER Then
  Call Tracer_XReturn("xStrExtract")
#End If
End Sub


'##D Returns a numbered field from a string of fields seperated by tokens
'##V String1            source string
'##V Token             seperator token
'##V TokenIndex         the index of the field to be extracted (1=Field1, 2=Field2)
'##V TrimString Trim spaces off the field (default = false)
'##RV the extracted string
'##C STF
Public Function GetStringToken(String1 As String, ByVal Token As String, Optional ByVal TokenIndex As Long = 1, Optional ByVal TrimString As Boolean = False) As String
  Dim pos As Long
  Dim lpos As Long
  Dim i As Long
  
  ' sort out apf tokenize XXX  vbTab YYY
#If DEBUGVER Then
  Call Tracer_XSet("GetStringToken")
#End If
  If Len(String1) = 0 Then GoTo GetStringToken_End
  
  pos = 0: i = 0
  Do
    lpos = pos
    pos = InStr(pos + 1, String1, Token, vbTextCompare)
    i = i + 1
  Loop Until (i = TokenIndex) Or (pos = 0)
  
  If pos = 0 Then
    If i <> TokenIndex Then
      GoTo GetStringToken_End
    End If
  End If
  
  If lpos = 0 Then
    If pos = 0 Then
      GetStringToken = String1
    Else
      GetStringToken = left$(String1, pos - 1)
    End If
  ElseIf pos = 0 Then
    GetStringToken = right$(String1, Len(String1) - lpos)
  ElseIf pos > 0 Then
    GetStringToken = Mid$(String1, lpos + 1, pos - lpos - 1)
  End If
  
  If TrimString Then GetStringToken = Trim$(GetStringToken)
  
GetStringToken_End:
#If DEBUGVER Then
  Call Tracer_XReturn("GetStringToken")
#End If
End Function

'##D Finds the first occurence of a character in String1 which is not in String2
'##V String1 source string
'##V String2 ignore character
'##V Start start point (Default = 1)
'##V Compare compare method (Default = vbBinaryCompare)
'##RV the position of the first character which is not String1 or if String2 is not in String1 then Len(String1)+1
Public Function NotInStr(String1 As String, ByVal String2 As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
#If DEBUGVER Then
  Call Tracer_XSet("NotInStr")
#End If
  NotInStr = NotInStrAny(String1, String2, Start, Compare)
#If DEBUGVER Then
  Call Tracer_XReturn("NotInStr")
#End If
End Function

'##D Parses a string of the form  WWW,XXX  , YYYY, "ZZZ""ZZZ",,"" into 5 values WWW XXX YYYY ZZZ"ZZZ (empty string) (empty string)
'##V DelimitedValue               Value parsed using m_EscapeChar and m_Delimiter
'##V DelimitedString buffer to parse
'##V offset               offset to begin parsing at
'##RV New offset in the buffer to parse, if > len(Buffer) then stop parsing string
'##C STF
Public Function GetDelimitedValue(DelimitedValue As String, DelimitedString As String, ByVal offset As Long, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """") As Long
#If DEBUGVER Then
  Call Tracer_XSet("GetDelimitedValue")
#End If
  GetDelimitedValue = GetDelimitedValueInt(DelimitedValue, DelimitedString, offset, True, Delimiter, EscapeChar)
#If DEBUGVER Then
  Call Tracer_XReturn("GetDelimitedValue")
#End If
End Function

'##C STF
Public Function GetDelimitedValueEx(DelimitedValue As String, DelimitedString As String, ByVal offset As Long, Optional ByVal TrimValues As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """") As Long
#If DEBUGVER Then
  Call Tracer_XSet("GetDelimitedValueEx")
#End If
  GetDelimitedValueEx = GetDelimitedValueInt(DelimitedValue, DelimitedString, offset, TrimValues, Delimiter, EscapeChar)
#If DEBUGVER Then
  Call Tracer_XReturn("GetDelimitedValueEx")
#End If
End Function

'##D Returns number of delimited items in an array of strings
'##LD ValueArray is an uninitialised array of strings - Will be initialised as 1-based
'##LD DelimitedString - Obvious
'##LD Delimiter - Again obvious. Default is Comma
'##LD IgnoreBlank - Determines whether blank elements are stored. Default is False
'##LD Returns Number of Delimited Items, Zero on Error
'##LD apf IgnoreBlanks !!!
'##C STF
Public Function GetDelimitedValues(ValueArray As Variant, DelimitedString As String, Optional IgnoreBlank As Boolean = False, Optional ByVal TrimValues As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal EscapeChar As String = """") As Long
#If DEBUGVER Then
  Call Tracer_XSet("GetDelimitedValues")
#End If
  GetDelimitedValues = GetDelimitedValuesEx(ValueArray, DelimitedString, IgnoreBlank, TrimValues, Delimiter, EscapeChar)
#If DEBUGVER Then
  Call Tracer_XReturn("GetDelimitedValues")
#End If
End Function


'##D Converts a directory into a path of the form X:\path\
'##C FNF
Public Function FullPath(ByVal sPath As String) As String
#If DEBUGVER Then
  Call Tracer_XSet("FullPath")
#End If
  FullPath = FullPathEx(sPath)
#If DEBUGVER Then
  Call Tracer_XReturn("FullPath")
#End If
End Function

'##C STF
Public Function ReplaceString(String1 As String, ByVal sFind As String, ByVal sReplace As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As String
  'ReplaceString = cReplaceString(String1, sFind, sReplace, Compare)
  ReplaceString = Replace(String1, sFind, sReplace, , , Compare)
End Function

'##D Replaces character in string with another
'##C STF
Public Function ReplaceChar(String1 As String, ByVal cFind As String, ByVal cReplace As String) As String
  'ReplaceChar = cReplaceString(String1, cFind, cReplace, vbTextCompare)
  ReplaceChar = Replace(String1, cFind, cReplace, , , vbTextCompare)
End Function

'##D Removes character from string
'##C STF
Public Function RemoveChar(String1 As String, ByVal cFind As String) As String
  'RemoveChar = cReplaceString(String1, cFind, "", vbTextCompare)
  RemoveChar = Replace(String1, cFind, "", , , vbTextCompare)
End Function

'##D Split fullpath into directory (including backslash) and filename and extention (.ext)
'#RV returns true if could split
'##C FNF
Public Sub SplitPath(ByVal FullPath As String, Optional Directory As Variant, Optional File As Variant, Optional Ext As Variant)
#If DEBUGVER Then
  Call Tracer_XSet("SplitPath")
#End If
  Call SplitPathEx(FullPath, Directory, File, Ext)
#If DEBUGVER Then
  Call Tracer_XReturn("SplitPath")
#End If
End Sub


'##D Sets the cursor, default is the hourglass
'##V StatusString A status string associated with the action
'##C FDF
Public Sub SetCursor(Optional ByVal Cursor As MousePointerConstants = vbHourglass, Optional ByVal StatusString As String = "")
#If DEBUGVER Then
  Call Tracer_XSet("SetCursor")
#End If
  Call SetCursorEx(Cursor, StatusString)
#If DEBUGVER Then
  Call Tracer_XReturn("SetCursor")
#End If
End Sub

'##D Clears the current cursor and restores the prior cursor on the cursor stack
'##RV Previous SetCursor status string
'##C FDF
Public Function ClearCursor() As String
#If DEBUGVER Then
  Call Tracer_XSet("ClearCursor")
#End If
  ClearCursor = ClearCursorEx(False)
#If DEBUGVER Then
  Call Tracer_XReturn("ClearCursor")
#End If
End Function

'##D Clears the current cursor and restores the prior cursor on the cursor stack, error if no cursor available
'##RV Previous SetCursor status string
'##C FDF
Public Function ClearCursorError() As String
  ClearCursorError = ClearCursorEx(True)
End Function

'##D Clears All Cursors
'##C FDF
Public Function ClearAllCursors() As String
#If DEBUGVER Then
  Call Tracer_XSet("ClearAllCursors")
#End If
  ClearAllCursors = ClearCursorEx(False, -1)
#If DEBUGVER Then
  Call Tracer_XReturn("ClearAllCursors")
#End If
End Function

'##D WinExec an application and show error messages
'##V ProgramPath Full path of program
'##RV double representing the program's task ID if successful, otherwise it returns zero
'##LD ##D WinExec an application and show error messages
'##LD if no path is specified windows will search the current dir, windows directory, system directory and the application path
'##C MF
Public Function ShellProgram(ByVal ProgramPath As String, Optional ByVal WinStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal Wait As Boolean = False) As Long
  Dim pid As Long, hproc As Long, hErr As Long, ExitCode As Long
  
  On Error GoTo shellprogram_err
#If DEBUGVER Then
  Call Tracer_XSet("ShellProgram")
#End If
  pid = CLng(Shell(ProgramPath, WinStyle))
  If Wait Then
    hproc = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    If hproc <> 0 Then
      Do
        hErr = GetExitCodeProcess(hproc, ExitCode)
        Call SleepW32(100)
      Loop Until (ExitCode <> STILL_ACTIVE) Or (hErr = 0)
    End If
  End If
  ShellProgram = pid
  
ShellProgram_end:
#If DEBUGVER Then
  Call Tracer_XReturn("ShellProgram")
#End If
  Exit Function

shellprogram_err:
  Call ErrorMessageEx(ERR_ERROR, Err, "ShellProgram", "ERR_WINEXEC", "Failed to winexec " & ProgramPath, False)
  Resume ShellProgram_end
End Function
'##C MF
Public Function ShellOpenFile(ByVal FilePath As String, Optional ByVal WinStyle As VbAppWinStyle = vbNormalFocus) As Long
  ShellOpenFile = ShellExecute(0, "open", FilePath, 0, "", WinStyle)
End Function

'##D Write an entry to the ini file specified or of the current application ([AppExeName].INI) if none is specified
'##V section         Ini file section name without brackets
'##V Key             Ini file key name
'##V Value           New value to be written
'##V IniFilePath    Path and file name of ini file (Default = [Windows directory]\[APPEXENAME].INI)
'##RV True if success
'##C IFF
Public Function WriteIniEntry(ByVal Section As String, ByVal Key As String, ByVal Value As String, Optional ByVal IniFilePath As String = "") As Boolean
  Dim sFname As String
  Dim l As Long
  
#If DEBUGVER Then
  Call Tracer_XSet("WriteIniEntry")
#End If
  If Len(IniFilePath) = 0 Then
    sFname = GetWindowsDirectory() & "\" & mAppExeName & ".INI"
  Else
    sFname = IniFilePath
  End If
  l = WritePrivateProfileString(Section, Key, Value, sFname)
  WriteIniEntry = (l <> 0&)
#If DEBUGVER Then
  Call Tracer_XReturn("WriteIniEntry")
#End If
End Function

'##D Read an entry from the ini file specified or of the current application ([AppExeName].INI)
'##V Section             section in ini file
'##V Key                 key
'##V Default               default value for the key (Default = "")
'##V inifilepath Path and filename of ini file (Default = [Windows directory]\[APPEXENAME].INI)
'##RV The value for the key in the ini file, or Default if the Key is not present
'##C IFF
Public Function GetIniEntry(ByVal Section As String, ByVal Key As String, Optional ByVal default As String = "", Optional ByVal IniFilePath As String = "") As String
#If DEBUGVER Then
  Call Tracer_XSet("GetIniEntry")
#End If
  GetIniEntry = GetIniEntryEx(Section, Key, default, IniFilePath)
#If DEBUGVER Then
  Call Tracer_XReturn("GetIniEntry")
#End If
End Function

'##D Returns an array KeyNames of all keys in a particular Section in an ini file
'##C IFF
Public Function GetIniKeyNamesEx(KeyNames As Variant, ByVal Section As String, Optional ByVal IniFilePath As String = "") As Long
#If DEBUGVER Then
  Call Tracer_XSet("GetIniKeyNamesEx")
#End If
  GetIniKeyNamesEx = GetIniKeyNamesExInternal(KeyNames, Section, IniFilePath)
#If DEBUGVER Then
  Call Tracer_XReturn("GetIniKeyNamesEx")
#End If
End Function

'##D Clears ini file entry
'##LD If no Key specified, clears complete Section otherwise clears that Key
'##C IFF
Public Function ClearIniEntry(ByVal Section As String, Optional ByVal Key As Variant, Optional ByVal Value As Variant, Optional ByVal IniFilePath As String = "") As Boolean
  Dim sFname As String
  Dim s As String
  Dim l As Long
  
#If DEBUGVER Then
  Call Tracer_XSet("ClearIniEntry")
#End If
  If Len(IniFilePath) = 0 Then
    sFname = GetWindowsDirectory() & "\" & mAppExeName & ".INI"
  Else
    sFname = IniFilePath
  End If
  If IsMissing(Key) Then
    l = WritePrivateProfileString(Section, vbNullString, vbNullString, sFname)
  ElseIf IsMissing(Value) Then
    s = CStr(Key)
    l = WritePrivateProfileString(Section, s, vbNullString, sFname)
  End If
  ClearIniEntry = (l <> 0&)
#If DEBUGVER Then
  Call Tracer_XReturn("ClearIniEntry")
#End If
End Function

'##D Similar to GetIniEntry, but acts on WIN.INI
'##C IFF
Public Function GetWinIniEntry(ByVal Section As String, ByVal Key As String, Optional ByVal default As String = "") As String
  Dim sFname As String
  Dim sBuffer As String, bsize As Long
  Dim retval As Long
  
#If DEBUGVER Then
  Call Tracer_XSet("GetWinIniEntry")
#End If

  bsize = TCSBUFSIZ
  Do
    bsize = bsize * 2
    sBuffer = String$(bsize, vbNullChar)
    retval = GetProfileString(Section, Key, default, sBuffer, bsize)
  Loop Until (retval = 0) Or (retval <> (bsize - 2))
  If retval = 0 Then
    sBuffer = ""
  Else
    sBuffer = RTrimChar(sBuffer, vbNullChar)
  End If
  
'  sBuffer = String$(TCSBUFSIZ, vbNullChar )
'  retval = GetProfileString(Section, Key, Default, sBuffer, TCSBUFSIZ)
'  If retval = 0 Then
'    sBuffer = ""
'  Else
'    sBuffer = RTrimChar(sBuffer, vbNullChar)
'  End If
  GetWinIniEntry = sBuffer
#If DEBUGVER Then
  Call Tracer_XReturn("GetWinIniEntry")
#End If
End Function

'##D Flush the ini buffer
'##LD Ini files are cached in memory this forces the file to write to disk
'##V IniFilePath  Name of ini file to flush default = windows directory\exename.ini
'##RV True if success
'##C IFF
Public Function FlushIniBuffer(Optional ByVal IniFilePath As String) As Boolean
  Dim l As Long, sFname As String
  
#If DEBUGVER Then
  Call Tracer_XSet("FlushIniBuffer")
#End If
  If Len(IniFilePath) = 0 Then
    sFname = GetWindowsDirectory() & "\" & mAppExeName & ".INI"
  Else
    sFname = IniFilePath
  End If
  l = FlushPrivateProfileString(0, 0, 0, sFname)
  FlushIniBuffer = (l <> 0&)
#If DEBUGVER Then
  Call Tracer_XReturn("FlushIniBuffer")
#End If
End Function


'##D Returns true if password is entered correctly (No password required out of hours)
'##C TCSPF
Public Function GetTCSPassword() As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("GetTCSPassword")
#End If
  GetTCSPassword = GetTCSPasswordEx(False, GetStaticEx("contact"), "")
#If DEBUGVER Then
  Call Tracer_XReturn("GetTCSPassword")
#End If
End Function

'##D Returns true if password is entered correctly (Password required out of hours)
'##C TCSPF
Public Function GetTCSPasswordAlways() As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("GetTCSPasswordAlways")
#End If
  GetTCSPasswordAlways = GetTCSPasswordEx(True, GetStaticEx("contact"), "")
#If DEBUGVER Then
  Call Tracer_XReturn("GetTCSPasswordAlways")
#End If
End Function

'##D get the Password for a specific date
'##C TCSPF
Public Function GetPassword(Optional ByVal d0 As Date = UNDATED) As Long
  GetPassword = GetPassword_Daily(d0)
End Function
'##C TCSPF
Public Sub TCSPasswordPrompts(ByVal Title As String, ByVal Prompt As String)
  gPasswordTitle = "QUERY_PASSWORD"
  gPasswordPrompt = "Warning - You are about to enter a system function."
End Sub

'##D GetCanonicalPathName return full UNC pathname
'##V FilePath File Path
'##C FNF
Public Function GetCanonicalPathName(ByVal FilePath As String) As String
  Dim sBuffer As String, sFileName As String
  Dim uncpath As String
  Dim lBufSiz As Long, retval As Long
      
On Error GoTo GetCanonicalPathName_err
#If DEBUGVER Then
  Call Tracer_XSet("GetCanonicalPathName")
#End If
  lBufSiz = 1024
  sBuffer = String$(lBufSiz, vbNullChar)
  sFileName = String$(lBufSiz, vbNullChar)
  retval = GetFullPathName(FilePath, lBufSiz, sBuffer, sFileName)
  If retval <> 0 Then
    GetCanonicalPathName = RTrimChar(sBuffer, vbNullChar)
    If left$(GetCanonicalPathName, 2) <> "\\" Then
      If GetNetResource(uncpath, left$(GetCanonicalPathName, 2)) = 0 Then
        GetCanonicalPathName = UCase$(uncpath & right$(GetCanonicalPathName, Len(GetCanonicalPathName) - 2))
      End If
    End If
  End If
  
GetCanonicalPathName_end:
#If DEBUGVER Then
  Call Tracer_XReturn("GetCanonicalPathName")
#End If
  Exit Function
  
GetCanonicalPathName_err:
  GetCanonicalPathName = ""
  Resume GetCanonicalPathName_end
End Function

'##D Get the free and total disk space of the drive specified
'##V DriveRootPath Root Path of drive
'##V Total Total space on this drive
'##V Free Free space on drive
'##RV true if successful
'##C SQF
Public Function GetDiskSpace(ByVal DriveRootPath As String, Total As Single, Free As Single, Optional ByVal MemUnits As MemoryUnit = BYTES) As Boolean
  Dim dUserFree As Double, dTotal As Double, dFree As Double
  
#If DEBUGVER Then
  Call Tracer_XSet("GetDiskSpace")
#End If
  
  GetDiskSpace = GetDiskSpaceEx(DriveRootPath, dTotal, dUserFree, dFree, MemUnits)
  Total = dTotal
  Free = dUserFree
#If DEBUGVER Then
  Call Tracer_XReturn("GetDiskSpace")
#End If
End Function

'##D Return the path of the Special folder in the shell
'##V  CSIDL indentifier of the special shell folder
'##RV Pathname of folder selected
'##D FNF
Public Function GetSpecialFolder(ByVal CSIDL As CSIDL_SPECIAL_FOLDERS) As String
  GetSpecialFolder = GetSpecialFolderEx(CSIDL)
End Function

'##D Display the (windows) standard folder browse dialog
'##V hWndOwner Handle to parent window (normally mdimain.hwnd)
'##RV Pathname of folder selected
'##D FNF
Public Function BrowseForFolder(ByVal hwndOwner As Long, Optional ByVal InitialDirectory As String = "", Optional ByVal Title As String = "Select a folder") As String
  Dim lpIDList As Long, lResult As Long
  Dim hr As Long
  Dim BI As BrowseInfo
  
  On Error GoTo BrowseForFolder_ERR
#If DEBUGVER Then
  Call Tracer_XSet("BrowseForFolder")
#End If
  If StrComp(right$(InitialDirectory, 1), "\") = 0 Then
    If StrComp(right$(InitialDirectory, 2), ":\") <> 0 Then InitialDirectory = left$(InitialDirectory, Len(InitialDirectory) - 1)
  End If
  
  If Not FileExistsEx(InitialDirectory, True, False) Then InitialDirectory = mHomeDirectory
  
  BI.hwndOwner = hwndOwner
  BI.lpszTitle = Title
  '1 = BIF_RETURNONLYDIRS
  BI.lpfnCallBack = AddressOfFunc(AddressOf BrowseCallback)
  BI.lParam = InitialDirectory
  BI.ulFlags = 1
  lpIDList = SHBrowseForFolder(BI)
  
  If lpIDList Then
    InitialDirectory = String$(TCSBUFSIZ, vbNullChar)
    lResult = SHGetPathFromIDList(lpIDList, InitialDirectory)
    If lResult <> 0 Then
      BrowseForFolder = RTrimChar(InitialDirectory, vbNullChar)
    End If
  End If
  
BrowseForFolder_END:
  If lpIDList Then Call CoTaskMemFree(lpIDList)
#If DEBUGVER Then
  Call Tracer_XReturn("BrowseForFolder")
#End If
  Exit Function
  
BrowseForFolder_ERR:
  BrowseForFolder = ""
  Call ErrorMessageEx(ERR_ERROR, Err, "BrowseForFolder", "Error " & Err.LastDllError, Err.Description, False)
  Resume BrowseForFolder_END
End Function

'##D Display the (windows) standard FileOpen Dialog Box
'##C FNF
Public Function FileOpenDlg(ByVal Title As String, ByVal Filter As String, ByVal InitialDirectory As String, Optional ByVal MultiSelect As Boolean = False) As String
  Dim cd As CommonDialog, hWnd As Long
  
  On Error GoTo FileOpenDlg_Err
#If DEBUGVER Then
  Call Tracer_XSet("FileOpenDlg")
#End If
  Set cd = frmAbout.cd
  'cdlOFNAllowMultiselect
  'cdlOFNFileMustExist
  'cdlOFNHideReadOnly
  'cdlOFNLongNames
  'cdlOFNNoDereferenceLinks
  cd.Flags = cdlOFNNoChangeDir + cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
  If MultiSelect Then cd.Flags = cd.Flags + cdlOFNAllowMultiselect
  cd.DialogTitle = Title
  cd.Filter = Filter
  cd.InitDir = InitialDirectory
  cd.CancelError = True
  cd.FileName = ""
  
dofileopen:
  cd.ShowOpen
  If Len(Trim$(cd.FileName)) = 0 Then GoTo dofileopen
  FileOpenDlg = cd.FileName
  
FileOpenDlg_End:
#If DEBUGVER Then
  Call Tracer_XReturn("FileOpenDlg")
#End If
  Exit Function

FileOpenDlg_Err:
  FileOpenDlg = ""
  Resume FileOpenDlg_End
End Function

'##D Display the (windows) standard FileOpen Dialog Box
'##C FNF
Public Function FileOpenDlgFilter(ByRef FilterIndex As Long, ByVal Title As String, ByVal Filter As String, ByVal InitialDirectory As String, Optional ByVal MultiSelect As Boolean = False) As String
  Dim cd As CommonDialog, hWnd As Long
  
  On Error GoTo FileOpenDlg_Err
#If DEBUGVER Then
  Call Tracer_XSet("FileOpenDlg")
#End If
  Set cd = frmAbout.cd
  'cdlOFNAllowMultiselect
  'cdlOFNFileMustExist
  'cdlOFNHideReadOnly
  'cdlOFNLongNames
  'cdlOFNNoDereferenceLinks
  cd.Flags = cdlOFNNoChangeDir + cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
  If MultiSelect Then cd.Flags = cd.Flags + cdlOFNAllowMultiselect
  cd.DialogTitle = Title
  Filter = RTrimChar(Filter, "|")
  cd.Filter = Filter
  If (FilterIndex < 1) Or (FilterIndex > ((CountChar(Filter, "|") + 1) / 2)) Then FilterIndex = 1
  cd.FilterIndex = FilterIndex
  cd.InitDir = InitialDirectory
  cd.CancelError = True
  cd.FileName = ""
  
dofileopen:
  cd.ShowOpen
  If Len(Trim$(cd.FileName)) = 0 Then GoTo dofileopen
  FileOpenDlgFilter = cd.FileName
  FilterIndex = cd.FilterIndex
  
FileOpenDlg_End:
#If DEBUGVER Then
  Call Tracer_XReturn("FileOpenDlg")
#End If
  Exit Function

FileOpenDlg_Err:
  FileOpenDlgFilter = ""
  Resume FileOpenDlg_End
End Function

'##D Display the (windows) standard FileSave Dialog box
'##C FNF
Public Function FileSaveAsDlg(ByVal Title As String, ByVal Filter As String, ByVal InitialDirectory As String, Optional ByVal FileName As String = "") As String
  Dim dosdir As String
  Dim cd As CommonDialog
  
  On Error GoTo FileSaveAsDlg_Err
#If DEBUGVER Then
  Call Tracer_XSet("FileSaveAsDlg")
#End If
  dosdir = GetDOSPath(InitialDirectory)
  If Not IsUNCPath(dosdir) Then
    ChDrive dosdir
    ChDir dosdir
  End If
  Set cd = frmAbout.cd
  'cdlOFNAllowMultiselect
  'cdlOFNFileMustExist
  'cdlOFNHideReadOnly
  'cdlOFNLongNames
  'cdlOFNNoDereferenceLinks
  cd.Flags = cdlOFNNoChangeDir + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
  cd.DialogTitle = Title
  cd.Filter = Filter
  cd.InitDir = dosdir
  cd.CancelError = True
  cd.FileName = FileName
  cd.ShowSave
  FileSaveAsDlg = cd.FileName
  
FileSaveAsDlg_End:
#If DEBUGVER Then
  Call Tracer_XReturn("FileSaveAsDlg")
#End If
  Exit Function

FileSaveAsDlg_Err:
  FileSaveAsDlg = ""
  Resume FileSaveAsDlg_End
End Function

Public Function FileSaveAsDlgFilter(ByRef FilterIndex As Long, ByVal Title As String, ByVal Filter As String, ByVal InitialDirectory As String, Optional ByVal FileName As String = "") As String
  Dim dosdir As String
  Dim cd As CommonDialog
  
  On Error GoTo FileSaveAsDlgFilter_Err
#If DEBUGVER Then
  Call Tracer_XSet("FileSaveAsDlgFilter")
#End If
  dosdir = GetDOSPath(InitialDirectory)
  If Not IsUNCPath(dosdir) Then
    ChDrive dosdir
    ChDir dosdir
  End If
  FilterIndex = 0
  Set cd = frmAbout.cd
  'cdlOFNAllowMultiselect
  'cdlOFNFileMustExist
  'cdlOFNHideReadOnly
  'cdlOFNLongNames
  'cdlOFNNoDereferenceLinks
  cd.Flags = cdlOFNNoChangeDir + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
  cd.DialogTitle = Title
  Filter = RTrimChar(Filter, "|")
  cd.Filter = Filter
  If (FilterIndex < 1) Or (FilterIndex > ((CountChar(Filter, "|") + 1) / 2)) Then FilterIndex = 1
  cd.FilterIndex = FilterIndex
  cd.InitDir = dosdir
  cd.CancelError = True
  cd.FileName = FileName
  cd.ShowSave
  FilterIndex = cd.FilterIndex
  FileSaveAsDlgFilter = cd.FileName

FileSaveAsDlgFilter_End:
#If DEBUGVER Then
  Call Tracer_XReturn("FileSaveAsDlgFilter")
#End If
  Exit Function

FileSaveAsDlgFilter_Err:
  FileSaveAsDlgFilter = ""
  Resume FileSaveAsDlgFilter_End
End Function



'##D Returns the UserName of the current connection
'##RV (UNKNOWN) if call failed
'##C SQF
Public Function GetNetUser(Optional ByVal ShowErrors As Boolean = False) As String
#If DEBUGVER Then
  Call Tracer_XSet("GetNetUser")
#End If
  GetNetUser = GetNetUser_s(ShowErrors)
#If DEBUGVER Then
  Call Tracer_XReturn("GetNetUser")
#End If
End Function
  
'##D Retrieves computer name
'##C SQF
Public Function GetComputerName() As String
  GetComputerName = GetComputerName_s()
End Function

'##D retrieves current windows directory (e.g C:\WINDOWS)
'##C SQF
Public Function GetWindowsDirectory() As String
#If DEBUGVER Then
  Tracer_XSet ("GetWindowsDirectory")
#End If
  GetWindowsDirectory = GetWindowsDirectoryEx()
#If DEBUGVER Then
  Tracer_XReturn ("GetWindowsDirectory")
#End If
End Function

'##D retrieves current system directory (e.g C:\WinNT\system32)
'##C SQF
Public Function GetSysDirectory() As String
#If DEBUGVER Then
  Tracer_XSet ("GetSysDirectory")
#End If
  GetSysDirectory = GetSysDirectoryEx()
#If DEBUGVER Then
  Tracer_XReturn ("GetSysDirectory")
#End If
End Function
  
'##D Gives (full) network resource name
'##V NetResourceName String where result is placed
'##V LocalResourceName Drive letter (1..27) or printer port
'##V ResourceNameOnly See above default, false
'##V ShowErrors Show error message on error
'##LD Given a local resource name ( e.g G:,H: or LPT1: ) and return either the full network
'##LD resource name ( e.g for files \\server\volume\path ) or just the network resource
'##LD name (server or printer)
'##RV 0 if succeeds, error code if failure (see DLLErrors enumeration)
'##C SQF
Public Function GetNetResource(NetResourceName As String, ByVal LocalResourceName As String, Optional ByVal ResourceNameOnly As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Long
  Dim sRes As String, p As Long
  On Error GoTo GetNetResource_ERR
  
#If DEBUGVER Then
  Call Tracer_XSet("GetNetResource")
#End If
  
  NetResourceName = ""
  sRes = String$(TCSBUFSIZ, vbNullChar)
  GetNetResource = WNetGetConnection(UCase$(LocalResourceName), sRes, TCSBUFSIZ)
  If GetNetResource = 0 Then
    sRes = RTrimChar(sRes, vbNullChar)
    NetResourceName = sRes
    ' note UNC of type \\xxxx
    If ResourceNameOnly And (Len(NetResourceName) > 2) Then
      p = InStr(3, NetResourceName, "\")
      If p > 0 Then
        NetResourceName = Mid$(NetResourceName, 3, p - 3)
      End If
    End If
    NetResourceName = sRes
  End If
  
GetNetResource_END:
#If DEBUGVER Then
  Call Tracer_XReturn("GetNetResource")
#End If
  Exit Function
  
GetNetResource_ERR:
  If ShowErrors Then Call ErrorMessageEx(ERR_ERROR, Err, "GetNetResource", "Network Resource Name", "Unable to get the network resource name for the resource " & LocalResourceName, False)
  Resume GetNetResource_END
End Function

'##D Locks updates
'##C MF
Public Function LockWindowUpdate(Optional ByVal hWnd As Long = 0) As Boolean
  Static winlocked As Long
  Dim l As Long
        
  If IsRunningInIDEEx Then
    Debug.Print "Call to LockWindowUpdate disabled in IDE"
    Exit Function
  End If
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

Public Function IsRunningInIDE() As Boolean
  IsRunningInIDE = IsRunningInIDEEx()
End Function

Public Function MultiDialog(ByVal Title As String, ByVal Message As String, ParamArray buttons() As Variant) As Long
  Dim fm As Form
  Dim cmb As CommandButton
  Dim l As Long, lDelta As Long
  
  On Error GoTo MultiDialog_err
  Call SetCursorEx(vbArrow, "")
  Set fm = frmMessage
  For l = LBound(buttons) To UBound(buttons)
    Load fm.cmdArray(l)
    Set cmb = fm.cmdArray(l)
    With cmb
      .top = fm.fmeMessage.top + 100 + (l * (cmb.Height + 100))
      .Caption = buttons(l)
      .Visible = True
    End With
  Next l
  fm.Caption = Title
  fm.lblMessage.Caption = Message
  fm.Show 1
  MultiDialog = fm.lClicked
  
MultiDialog_end:
  Call ClearCursor
  Set fm = Nothing
  Exit Function
  
MultiDialog_err:
  MultiDialog = -1
  Resume MultiDialog_end
End Function

'##D Sets label/textbox text
'##V lbl label or textbox
'##V Text Text for label or textbox
'##C FDF
Public Sub SetText(lbl As Object, ByVal Text As String)
  On Error GoTo SetText_err
  
#If DEBUGVER Then
  Call Tracer_XSet("SetText")
#End If
  If TypeOf lbl Is Label Then
    lbl.Caption = Text
  ElseIf TypeOf lbl Is TextBox Then
    lbl.Text = Text
  Else
    Call ECASE_SYS("SetText - only valid on textboxes and labels")
    Exit Sub
  End If
  lbl.Refresh
  
SetText_end:
#If DEBUGVER Then
  Call Tracer_XReturn("SetText")
#End If
  Exit Sub
  
SetText_err:
  Resume SetText_end
End Sub

'* True for success, False (+ raised error otherwise)
Public Function PopUpButtons(ByVal hWndToolBarOCX As Long, Optional ByVal TopDivider As Boolean = False) As Boolean
  Dim hWndToolbarDLL As Long, lStyle As Long
  Dim hErr As Long
  
#If DEBUGVER Then
  Call Tracer_XSet("PopUpButtons")
#End If
  
  lStyle = TBSTYLE_FLAT
  If TopDivider = 0 Then
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
#If DEBUGVER Then
  Call Tracer_XReturn("PopUpButtons")
#End If
  Exit Function
End Function

'##D Queries the Current Printer for devicename and Port
'##RV True if success
'##C SQF
Public Function GetDefaultPrinter(DeviceName As String, PortName As String) As Boolean
  On Error GoTo GetDefaultPrinter_err
  
  DeviceName = Printer.DeviceName
  PortName = Printer.Port
  GetDefaultPrinter = True
  
GetDefaultPrinter_end:
  Exit Function
  
GetDefaultPrinter_err:
  GetDefaultPrinter = False
  DeviceName = "":  PortName = ""
  Resume GetDefaultPrinter_end
End Function

Public Property Get FatalError() As Boolean
  FatalError = m_FatalError
End Property

'##D Stop system for specified amount of time
'##C MF
Public Sub Sleep(ByVal Milliseconds As Long)
  Call SleepW32(Milliseconds)
End Sub

Public Function SetCapture(ByVal hWnd As Long) As Long
  SetCapture = SetCapture32(hWnd)
End Function

Public Sub ReleaseCapture()
  Call ReleaseCapture32
End Sub

'##C SQF
Public Property Get TCSCoreDebug() As Boolean
  #If DEBUGVER Then
    TCSCoreDebug = True
  #Else
    TCSCoreDebug = False
  #End If
End Property

'##C SQF
Public Function isCOMPresent(ByVal ProgID As String, ByVal ComType As COM_TYPE) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("isCOMPresent")
#End If
  isCOMPresent = isCOMPresentEx(ProgID, ComType)
#If DEBUGVER Then
  Call Tracer_XReturn("isCOMPresent")
#End If
End Function

'##D Get the path of the registered component identified by the ProgID
'##C SQF
Public Function GetComponentPath(ByVal ProgID As String) As String
  GetComponentPath = GetComponentPathEx(ProgID)
End Function

'##D Copies file to another location
'##LD This function replaces the normal FileCopy in VB, the function implements file copy as a copy and rename
'##LD and takes account of most long filename issues
'##V Source file
'##V Destination file
'##RV True if copy successful
'##C FNF
Public Function FileCopyEx(ByVal Source As String, ByVal Destination As String) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("FileCopyEx")
#End If
  FileCopyEx = FileCopyExN(Source, Destination, False)
#If DEBUGVER Then
  Call Tracer_XReturn("FileCopyEx")
#End If
End Function

'##D Copies file to another location
'##LD This function replaces the normal FileCopy in VB, the function implements file copy as a copy and rename
'##LD and takes account of most long filename issues.
'##LD Raises an error if Copy Fails
'##V Source file
'##V Destination file
'##C FNF
Public Sub FileCopyEx2(ByVal Source As String, ByVal Destination As String)
#If DEBUGVER Then
  Call Tracer_XSet("FileCopyEx2")
#End If
  Call FileCopyExN(Source, Destination, True)
#If DEBUGVER Then
  Call Tracer_XReturn("FileCopyEx2")
#End If
End Sub

'##C SORTF
Public Property Get SortElementComparisons() As Long
  SortElementComparisons = SortComparisons
End Property

'##C SORTF
Public Sub CombSortEx(v As Variant, ByVal lmin As Long, ByVal lmax As Long, SortFn As ISortFunction)
  Call SortAny(v, lmin, lmax, SortFn, COMB_SORT)
End Sub

'##C SORTF
Public Sub QSortEx(v As Variant, ByVal lmin As Long, ByVal lmax As Long, SortFn As ISortFunction)
  Call SortAny(v, lmin, lmax, SortFn, QUICK_SORT)
End Sub

'##C SORTF
Public Sub QSortAsc(v As Variant, ByVal lmin As Long, ByVal lmax As Long)
  Dim SortFn As SortAscending
  Set SortFn = New SortAscending
  Call SortAny(v, lmin, lmax, SortFn, QUICK_SORT)
End Sub

'##C SORTF
Public Sub QSortDesc(v As Variant, ByVal lmin As Long, ByVal lmax As Long)
  Dim SortFn As New SortDescending
  Set SortFn = New SortDescending
  Call SortAny(v, lmin, lmax, SortFn, QUICK_SORT)
End Sub

'##D Closes All Forms (if Form Tag is NOUNLOAD then do not unload form)
'##C FDF
Public Function CloseAllForms(FormsCollection As Object, ByVal ForceClose As Boolean) As Boolean
  Dim j As Long, i As Long
  Dim frm As Form
  
  On Error GoTo CloseAllForms_Err
#If DEBUGVER Then
  Call Tracer_XSet("CloseAllForms")
#End If
  j = 0
  For i = (FormsCollection.Count - 1) To 0 Step -1
    Set frm = FormsCollection(j)
    If (InStr(frm.Tag, "NOUNLOAD") = 0) Or ForceClose Then
      If Not (TypeOf frm Is MDIForm) Then
        Unload frm
        Set frm = Nothing
        j = j - 1
      End If
    End If
    j = j + 1
  Next i
  CloseAllForms = True
 
CloseAllForms_End:
#If DEBUGVER Then
  Call Tracer_XReturn("CloseAllForms")
#End If
  Exit Function

CloseAllForms_Err:
  CloseAllForms = False
  Resume CloseAllForms_End
End Function

'##D Generated a temporary file name
'##V DefaultDir Default Directory
'##V FilePrefix File Prefix, only uses the first 3 characters
'##C FNF
Public Function GetTempFileName(ByVal DefaultDir As String, ByVal FilePrefix As String) As String
  GetTempFileName = GetTempFileNameEx(DefaultDir, FilePrefix)
End Function

Public Function GetTempDirectory() As String
  GetTempDirectory = GetTempDirectoryEx()
End Function

Public Property Get STR_CR() As String
  STR_CR = vbCrLf
End Property

'##D Error filters can be used to capture error messages
'##LD Error filters implement the IErrorFilter interface
'##C EMF
Public Property Get ErrorFilter() As IErrorFilter
  Set ErrorFilter = m_IPreErrorFilter
End Property

' Sets ErrorLogRecordSet
Public Property Set ErrorFilter(NewValue As IErrorFilter)
  Set m_IPreErrorFilter = NewValue
End Property

'##D Similar to an error filter, however the error is displayed as per usual and the client application can use this as a Notify interface
'##C EMF
Public Property Get ErrorPostProcess() As IErrorPostProcess
  Set ErrorPostProcess = m_IPostErrorProcess
End Property

Public Property Set ErrorPostProcess(NewValue As IErrorPostProcess)
  Set m_IPostErrorProcess = NewValue
End Property

Public Property Get DebugMenuInterface() As IDebugMenu
  Set DebugMenuInterface = m_IDebugMenu
End Property
Public Property Set DebugMenuInterface(NewValue As IDebugMenu)
  Set m_IDebugMenu = NewValue
End Property

Public Sub RegisterLibrary(NewValue As ILibraryVersion)
  Dim lv As LibraryVersion
  
  If Not NewValue Is Nothing Then
    For Each lv In LibraryVersions
      If StrComp(lv.Name, NewValue.Name, vbBinaryCompare) = 0 Then Exit Sub
    Next lv
    Set lv = New LibraryVersion
    lv.Name = NewValue.Name
    lv.Version = NewValue.Version
    Call LibraryVersions.Add(lv)
  End If
End Sub

'##C TCSMF
Public Sub SetHelpAboutText(ByVal NewValue As String)
  Const USER_OFFSET As Single = 50
  Dim fBot As Single
  
  m_HelpAboutText = NewValue
  frmAbout.lblUser.Caption = ""
  frmAbout.lblUser.Height = 0
  frmAbout.lblUser.Caption = NewValue
  frmAbout.lblName.top = frmAbout.lblUser.top + frmAbout.lblUser.Height + USER_OFFSET
  frmAbout.lblVersion.top = frmAbout.lblName.top
  Call FillVersions(frmAbout.lblName, frmAbout.lblVersion)
  
  frmAbout.lblContact.Caption = GetStaticEx("CONTACT")
  frmAbout.lblContact.top = frmAbout.lblVersion.top + frmAbout.lblVersion.Height + USER_OFFSET
  fBot = frmAbout.lblContact.top + frmAbout.lblContact.Height
  If frmAbout.lblContact.Height > frmAbout.cmdOK.Height Then
    frmAbout.cmdOK.top = fBot - frmAbout.cmdOK.Height
  Else
    frmAbout.cmdOK.top = frmAbout.lblContact.top
  End If
  frmAbout.lblCopyright.top = frmAbout.cmdOK.top + frmAbout.cmdOK.Height + USER_OFFSET
  fBot = frmAbout.cmdOK.top + frmAbout.cmdOK.Height + frmAbout.lblCopyright.Height + USER_OFFSET
  frmAbout.Height = fBot + USER_OFFSET + 500
End Sub
'##D Starts Timer
'##C MF
Public Sub StartTimer(ByVal Interval As Long, tProc As ITimerProc)
  Call StartTimerEx(Interval, tProc)
End Sub

'##D Stops Timer
'##C MF
Public Sub StopTimer(tProc As ITimerProc)
  Call KillTimerEx(tProc)
End Sub

'##D Register a private clipboard format
'##C CF
'##V NewFormat Systemwide name of format to register
'##RV Handle to clipboard format used to refer to the private format
Public Function RegisterPrivateClipboard(ByVal NewFormat As String) As Long
  If Len(NewFormat) > 0 Then
    RegisterPrivateClipboard = RegisterClipboardFormat(NewFormat)
  End If
End Function

'##D Sets contents of clipboard (Text and multiple private formats)
'##C CF
'##V PrivateFormats an array of clipboard handle, clipboard data pairs
'##RV True if function is successful
Public Function SetAnyClipboardData(PublicText As String, ParamArray PrivateFormats() As Variant) As Boolean
  Dim i As Long
  Dim ClipFormat As Long, ClipData As String
  
  If OpenClipboard(0) Then
    Call EmptyClipboard
    SetAnyClipboardData = True
    SetAnyClipboardData = SetAnyClipboardData And SetAnyClipboardDataEx(vbCFText, PublicText)
    i = LBound(PrivateFormats)
    Do While i <= (UBound(PrivateFormats) - 1)
      ClipFormat = PrivateFormats(i)
      ClipData = PrivateFormats(i + 1)
      SetAnyClipboardData = SetAnyClipboardData And SetAnyClipboardDataEx(ClipFormat, ClipData)
      i = i + 2
    Loop
    Call CloseClipboard
  End If
End Function
   
'##D Gets contents of clipboard
'##V clipformat Can be any clipboard format (including private formats)
'##C CF
Public Function GetAnyClipboardData(ByVal ClipFormat As VBRUN.ClipBoardConstants) As String
  GetAnyClipboardData = GetAnyClipboardDataEx(ClipFormat)
End Function

'##D Displays a Message Box in centered on InForm
'##C FDF
Public Function DisplayMessage(InForm As Object, ByVal Message As String, ByVal Title As String, Optional ByVal OKText As String = "Ok", Optional ByVal CancelText As String = "Cancel") As Boolean
  DisplayMessage = DisplayMessageEx(InForm, Message, Title, OKText, CancelText)
End Function

'##D Displays a Message Box in centered on InForm for Given Period of Time, then closes the message box
'##C FDF
Public Function DisplayMessageTimeOut(InForm As Object, ByVal Message As String, ByVal Title As String, Optional ByVal TimeOut As Long = 5) As Boolean
  
  On Error GoTo DisplayMessageTimeOut_Err
  
  Load frmMSGOKCancel
  DisplayMessageTimeOut = frmMSGOKCancel.displayMsg(InForm, Message, Title, "", "", TimeOut)
  Unload frmMSGOKCancel
  Set frmMSGOKCancel = Nothing
  DoEvents
  
DisplayMessageTimeOut_End:
  Exit Function

DisplayMessageTimeOut_Err:
  DisplayMessageTimeOut = False
  Resume DisplayMessageTimeOut_End
End Function

'##D Displays a Message Box in centered on InForm
'##LD The message box acts like a floating status window, and is closed using DisplayMessageKill.
'##LD Multiple calls to DisplayMessagePopup change the message displayed
'##C FDF
Public Sub DisplayMessagePopup(InForm As Object, ByVal Message As String, ByVal Title As String)
  
  On Error GoTo DisplayMessagePopup_Err
  Load frmMSGOKCancel
  Call frmMSGOKCancel.displayMsg(InForm, Message, Title, "", "", NO_TIMEOUT, frmMSGOKCancel.Visible)
  
DisplayMessagePopup_End:
  Exit Sub

DisplayMessagePopup_Err:
  Resume DisplayMessagePopup_End
End Sub

'##D Unloads Message Box
'##C FDF
Public Sub DisplayMessageKill()

  On Error GoTo DisplayMessageKill_Err
  Unload frmMSGOKCancel
  Set frmMSGOKCancel = Nothing
  DoEvents
  
DisplayMessageKill_End:
  Exit Sub

DisplayMessageKill_Err:
  Resume DisplayMessageKill_End
End Sub

'##D Determines whether a variable is Boolean
'##LD Takes account of numbers and text true/false, on/off, yes/no
'##C GCF
Public Function IsBoolean(v As Variant) As Boolean
  IsBoolean = IsBooleanEx(v)
End Function

'##D Converts Variable to Boolean
'##LD Takes account of numbers and text true/false, on/off, yes/no
'##C GCF
Public Function CBoolean(v As Variant) As Boolean
  CBoolean = CBooleanEx(v)
End Function

'##C MF
Public Function InCollection(col As Object, vItem As Variant) As Boolean
  InCollection = InCollectionEx(col, vItem)
End Function

'##D Gets current time
'##RV Current time
'##C SQF
Public Function GetTime() As Long
  GetTime = timeGetTime
End Function

'##D Determines if a database is open
'##C FNF
Public Function IsDatabaseOpen(ByVal FileName As String, Optional ByVal Exclusive As Boolean = False) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("IsDatabaseOpen")
#End If
  IsDatabaseOpen = IsFileOpenEx(FileName, Exclusive)
#If DEBUGVER Then
  Call Tracer_XReturn("IsDatabaseOpen")
#End If
End Function

'##D Determines if a file is already open
'##V FileName File Name
'##RV True if file is open
'##C FNF
Public Function IsFileOpen(ByVal FileName As String, Optional ByVal Exclusive As Boolean = False) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("IsFileOpen")
#End If
  IsFileOpen = IsFileOpenEx(FileName, Exclusive)
#If DEBUGVER Then
  Call Tracer_XReturn("IsFileOpen")
#End If
End Function

'##D Rounds Given Number to nearest Integer
'##C NF
Public Function RoundN(ByVal Number As Double, Optional ByVal DecimalPlaces As Long = 2) As Double
  RoundN = RoundDouble(Number, DecimalPlaces, R_NORMAL)
End Function
'##D Rounds Up Given Number (Must not be negative)
'##C NF
Public Function RoundUp(ByVal Number As Double, Optional ByVal DecimalPlaces As Long = 2) As Double
  If Number < 0 Then Call Err.Raise(ERR_ROUND, "RoundUp", "Unable to round negative numbers up")
  RoundUp = RoundDouble(Number, DecimalPlaces, R_UP)
End Function

'##D Rounds Down Given Number (must not be negative)
'##C NF
Public Function RoundDown(ByVal Number As Double, Optional ByVal DecimalPlaces As Long = 2) As Double
  If Number < 0 Then Call Err.Raise(ERR_ROUND, "RoundDown", "Unable to round negative numbers down")
  RoundDown = RoundDouble(Number, DecimalPlaces, R_DOWN)
End Function

'##D Ends Application
'##C MF
Public Sub AppEnd()
  Call gTCSEventClass.AttemptEndApplication
End Sub
'##C TCSMF
Public Sub RefreshF12Menu()
  Call gTCSEventClass.FillDebugMenu(frmDebugMenu)
End Sub

'##D Converts a string to a Value of a specific data type
'##LD This is normally used for converting from saved representations ( web forms etc. ) to typed values
'##LD Returns null if type cannot be converted or it is the empty string
'##LD GetTypedValueAsString is used for converting back
'##C GCF
Public Function GetTypedValueFromString(ByVal sValue As String, ByVal dType As DATABASE_FIELD_TYPES) As Variant
  On Error GoTo GetTypedValueFromString_err
  sValue = Trim$(sValue)
  If Len(sValue) = 0 Then
    GetTypedValueFromString = Null
  Else
    GetTypedValueFromString = GetTypedValueEx(sValue, dType)
  End If
GetTypedValueFromString_end:
  Exit Function
  
GetTypedValueFromString_err:
  GetTypedValueFromString = Null
  Resume GetTypedValueFromString_end
End Function

'##D Converts a variant to a Value of a specific data type
'##LD If conversion fails then raises an error
'##C GCF
Public Function GetTypedValue(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As Variant
  GetTypedValue = GetTypedValueEx(v, dType)
End Function

'##D Converts a variant to a Value of a specific data type
'##LD If conversion fails then returns DefaultValue
'##C GCF
Public Function GetTypedValueDefault(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES, ByVal DefaultValue As Variant) As Variant
  On Error GoTo GetTypedValueDefault_err
  GetTypedValueDefault = GetTypedValueEx(v, dType)
  
GetTypedValueDefault_end:
  Exit Function
  
GetTypedValueDefault_err:
  GetTypedValueDefault = DefaultValue
  Resume GetTypedValueDefault_end
End Function

'##D Converts a variant to a Value of a specific data type
'##LD Deals with Null values, GetTypedValueFromString will convert back to a variant
'##C GCF
Public Function GetTypedValueAsString(ByVal vValue As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  Select Case dType
    Case TYPE_BOOL
      vValue = CBoolean(IsNullEx(vValue, False))
      vValue = CLng(vValue)
    Case TYPE_DOUBLE, TYPE_LONG
      vValue = IsNullEx(vValue, "")
    Case TYPE_DATE
      vValue = IsNullEx(vValue, UNDATED)
      If vValue = UNDATED Then
        vValue = ""
      Else
        vValue = DateString(vValue)
      End If
    Case Else
      vValue = IsNullEx(vValue, "")
  End Select
  GetTypedValueAsString = vValue
End Function

'##D Converts Value to a string
'##LD Does not deal with Null values, GetTypedValue will convert back to a variant
'##C GCF
Public Function GetValueAsString(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  If IsNull(v) Then Call Err.Raise(ERR_INVALID_TYPE, "GetValueAsString", "Unable to convert a Null Value to a string, use .")
  v = GetTypedValueEx(v, dType)
  Select Case dType
    Case TYPE_STR, TYPE_LONG, TYPE_DOUBLE
      GetValueAsString = CStr(v)
    Case TYPE_DATE
      GetValueAsString = DateString(v)
    Case TYPE_BOOL
      GetValueAsString = IIf(v, "True", "False")
    Case Else
      Call Err.Raise(ERR_INVALID_TYPE, "GetValueAsString", "Unrecognised type: " & CStr(dType))
  End Select
End Function

'##D Converts Value to String variable for use in SQL statements ( such as SELECT/INSERT )
'##LD Always converts to a string SQL value no matter what the underlying type
'##C GCF
Public Function GetSQLValueAsString(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  If Not (IsNull(v) Or IsEmpty(v)) Then v = GetValueAsString(v, dType)
  GetSQLValueAsString = StrSQL(v)
End Function

'##D Converts Value to SQL String variable for use in SQL statements ( such as SELECT/INSERT )
'##C GCF
Public Function GetSQLValue(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  If IsNull(v) Or IsEmpty(v) Then
    v = "Null"
  Else
    v = GetTypedValueEx(v, dType)
    Select Case dType
      Case TYPE_LONG, TYPE_DOUBLE
        v = NumSQL(v)
      Case TYPE_STR
        v = StrSQL(v)
      Case TYPE_DATE
        If IsDateTime(v) Then
          v = DateTimeSQL(v)
        Else
          v = DateSQL(v)
        End If
      Case TYPE_BOOL
        v = NumSQL(v)
    End Select
  End If
  GetSQLValue = v
End Function

'##D Adds Escape Characters to String
'##GCF
Public Function AddEscapeChars(ByVal s As String) As String
  If (InStr(s, ",") > 0) Or (InStr(s, """") > 0) Then
    If Not ((StrComp(left$(s, 1), """") = 0) And (StrComp(Mid$(s, Len(s), 1), """") = 0)) Then
      s = """" & StrDupChar(s, """") & """"
    End If
  End If
  AddEscapeChars = s
End Function

'##D Resets the ignore error number list
'##LD If an error message has the ERR_ALLOWIGNORE flag set then the user can ignore future errors of that type, this call allows you to reinitialise the ignore error list
'##C EMF
Public Sub ResetIgnoreError(Optional ByVal ErrorNumber As TCSERROR_CONSTANTS = -1)
  Call ClearIgnoreErrors(ErrorNumber)
End Sub
'##C MF
Public Property Set KeyboardHook(NewValue As IKeyboardHook)
  Set gKeyHook = NewValue
End Property

'##D Compares Version Numbers of two files sPathAndFile1 and sPathAndFile2
'##C SQF
Public Function VerComp(ByVal sPathAndFile1 As String, ByVal sPathAndFile2 As String) As Long
#If DEBUGVER Then
  Call Tracer_XSet("VerComp")
#End If
  VerComp = VerCompEx(sPathAndFile1, sPathAndFile2)
#If DEBUGVER Then
  Call Tracer_XReturn("VerComp")
#End If
End Function

'##D Get KeyState of keyboard control keys
'##C SQF
Public Function GetControlKeysState() As ShiftConstants
  GetControlKeysState = GetControlKeysStateEx()
End Function

'##C SQF
Public Function VerQuery(sPropertyResult As String, ByVal sPathAndFile As String, Optional ByVal VQT As VER_QUERY_TYPE = VQT_FILE_VERSION) As Boolean
#If DEBUGVER Then
 Call Tracer_XSet("VerQuery")
#End If
 sPropertyResult = VersionQueryMap(sPathAndFile, VQT)
 VerQuery = (Len(sPropertyResult) > 0)
#If DEBUGVER Then
 Call Tracer_XReturn("VerQuery")
#End If
End Function

Public Function GetTicks() As Long
  GetTicks = timeGetTime
End Function

Public Function IsNullEx(ByVal v As Variant, ByVal NullValue As Variant) As Variant
  If IsNull(v) Then
    IsNullEx = NullValue
  Else
    IsNullEx = v
  End If
End Function

'##D Is the variable a valid array which can be redim preserved and looped through
'##C STF
'##RV True if Array is valid
Public Function IsArrayEx(ByRef v As Variant) As Boolean
  IsArrayEx = IsArrayEx2(v)
End Function

Public Sub ActivatePrevInstance()
  Call ActivatePrevInstanceInt
End Sub

'##D Auto width the columns in a Listview control, contol must be populated first
'##C STF
Public Sub AutoWidthListView(ByVal lv As Object, Optional ByVal IncludeColumnHeaders As Boolean = False)
  Call AutoWidthListViewEx(lv, IncludeColumnHeaders)
End Sub

'##D Finds the position of the first occurance of any character in StringMatch in StringCheck
'##C STF
'##CCORE Public Function InStrAny(String1 As String, ByVal String2 As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

'##D Finds the position of the last occurance of any character in StringMatch in StringCheck
'##C STF
'##CCORE Public Function InStrAnyRev(StringCheck As Variant, ByVal StringMatch As String, Optional ByVal Start As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

'##D Finds the position of the first occurance of any character not in StringMatch in StringCheck
'##C STF
'##CCORE Public Function NotInStrAny(String1 As String, ByVal String2 As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

'##D Finds the position of the last occurance of any character not in StringMatch in StringCheck
'##C STF
'##CCORE Public Function NotInStrAnyRev(StringCheck As Variant, ByVal StringMatch As String, Optional ByVal Start As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

'##D Converts Character to upper case
'##C STF
'##CCORE Public Function ToUpper(ByVal KeyAscii As Integer) As Integer

'##D Converts Character to lower case
'##C STF
'##CCORE Public Function ToLower(ByVal KeyAscii As Integer) As Integer

'##D Is the Character printable
'##C STF
'##CCORE Public Function IsPrint(ByVal KeyAscii As Integer) As Boolean

'##D Is the Character in the standard Ascii character set
'##C STF
'##CCORE Public Function IsAscii(ByVal KeyAscii As Integer) As Boolean

'##D Returns a new string with the character char stripped from the left
'##s source string
'##char character to strip off, uses first character in string
'##RV the trimmed string; if char is empty then the source string is returned
'##C STF
'##CCORE Public Function LTrimChar(String1 As String, ByVal Char As String) As String

'##D Returns a new string with the character char stripped from the right
'##s source string
'##char character to strip off, uses first character in string
'##RV the trimmed string; if char is empty then the source string is returned
'##C STF
'##CCORE Public Function RTrimChar(String1 As String, ByVal Char As String) As String

'##D Returns a new string with the character char stripped from the left and right side
'##s source string
'##char character to strip off, if length greater than one then take the left most character
'##RV the trimmed string; if char is empty then the source string is returned
'##C STF
'##CCORE Public Function TrimChar(String1 As String, ByVal Char As String) As String

Public Function DataTypeName(ByVal dType As DATABASE_FIELD_TYPES) As String
  Dim dTypeName As String
  
  dTypeName = "Unrecognised data Type: " & CLng(dType)
  Select Case dType
    Case [_TYPE_INVALID]
      dTypeName = "Invalid"
    Case TYPE_LONG
      dTypeName = "Long"
    Case TYPE_STR
      dTypeName = "String"
    Case TYPE_DOUBLE
      dTypeName = "Double"
    Case TYPE_DATE
      dTypeName = "Date"
    Case TYPE_BOOL
      dTypeName = "Boolean"
    Case TYPE_BLOB
      dTypeName = "Binary large object"
  End Select
  DataTypeName = dTypeName
End Function

Public Function VarTypetoDatatype(ByVal vbType As VbVarType) As DATABASE_FIELD_TYPES
  VarTypetoDatatype = VarTypetoDatatypeEx(vbType)
End Function

Public Function IsNumberField(ByVal dbDataType As DATABASE_FIELD_TYPES) As Boolean
  IsNumberField = (dbDataType = TYPE_DOUBLE) Or (dbDataType = TYPE_LONG)
End Function

Public Property Get SQLDriverName() As String
  Select Case DatabaseTarget
    Case DB_TARGET_JET
           SQLDriverName = ""
    Case DB_TARGET_SQLSERVER
           SQLDriverName = "SQL Server"
    Case DB_TARGET_ORACLE
           SQLDriverName = "Oracle ODBC Driver"
    Case Else
           Call ECASE("SQLDriverName - Unknown DatabaseTarget: " & CStr(DatabaseTarget))
  End Select
End Property

Public Function DSNAttributes(ByVal Description As String, ByVal ServerName As String, ByVal DBName As String, ByVal NetDriver As String) As String
  Dim s As String
  
  s = "Description=" & Description & vbCr
  If DatabaseTarget = DB_TARGET_ORACLE Then
    s = s & "DBQ=" & ServerName & vbCr
  Else
    s = s & "SERVER=" & ServerName & vbCr & _
            "Database=" & DBName
  End If
  If Len(NetDriver) > 0 Then s = s & vbCr & "Network=" & NetDriver
  DSNAttributes = s
End Function

Public Function DSNConnectString(ByVal DSN_Name As String, ByVal UserID As String, ByVal UserPassword As String) As String
  Dim s As String
  
  s = "DSN=" & DSN_Name
  If Len(UserID) > 0 Then s = s & ";UID=" & UserID
  If Len(UserPassword) > 0 Then s = s & ";PWD=" & UserPassword
  DSNConnectString = s
End Function

Public Sub SetWindowZOrder(ByVal hWnd As Long, ByVal PositionFlag As Long)
  Call SetWindowZOrderEx(hWnd, PositionFlag)
End Sub
      
Public Property Let ForceErrorTopMost(ByVal NewValue As Boolean)
  mForceErrorTopMost = NewValue
End Property

Public Property Get ForceErrorTopMost() As Boolean
  ForceErrorTopMost = mForceErrorTopMost
End Property

Public Function GetDecimalPrecision(ByVal sFormat As String, Optional ByVal DecimalPlaceholder As String = ".") As Long
  Dim p0 As Long, p1 As Long
  
  GetDecimalPrecision = 0
  p0 = InStr(1, sFormat, DecimalPlaceholder, vbBinaryCompare)
  If p0 > 0 Then
    p1 = NotInStrAny(sFormat & ";", "#0123456789", p0 + 1, vbBinaryCompare) - 1
    GetDecimalPrecision = p1 - p0
  End If
  If GetDecimalPrecision > 14 Then Err.Raise ERR_PRECISION, "GetDecimalPrecision", "Numbers (Doubles) are limited to 14 significant digits after the decimal point"
End Function

Public Function CheckMouseButton(ByVal Button As BUTTON_CONSTANTS) As Boolean
  If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
    Button = IIf(Button = VK_LBUTTON, VK_RBUTTON, VK_LBUTTON)
  End If
  CheckMouseButton = (GetAsyncKeyState(Button) <> 0)
End Function
Public Function RegGetKeyValue(ByVal KeyName As String, ByVal ValueName As String) As String
  RegGetKeyValue = xRegGetKeyValue(KeyName, ValueName)
End Function

'cad for XDB
'##D Returns true if password is entered correctly (No password required out of hours)
'##C TCSPF
Public Function GetTCSPassword2(ByVal Reason As String) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("GetTCSPassword2")
#End If
  GetTCSPassword2 = GetTCSPasswordEx(False, GetStaticEx("contact"), Reason)
#If DEBUGVER Then
  Call Tracer_XReturn("GetTCSPassword2")
#End If
End Function
'##D Returns true if password is entered correctly (Password required out of hours)
'##C TCSPF
Public Function GetTCSPasswordAlways2(ByVal Reason As String) As Boolean
#If DEBUGVER Then
  Call Tracer_XSet("GetTCSPasswordAlways2")
#End If
  GetTCSPasswordAlways2 = GetTCSPasswordEx(True, GetStaticEx("contact"), Reason)
#If DEBUGVER Then
  Call Tracer_XReturn("GetTCSPasswordAlways2")
#End If
End Function
Public Property Get NotifyTCSPassword() As IBaseNotify
  Set NotifyTCSPassword = m_INotifyTCSPassword
End Property
Public Property Set NotifyTCSPassword(NewValue As IBaseNotify)
  Set m_INotifyTCSPassword = NewValue
End Property

