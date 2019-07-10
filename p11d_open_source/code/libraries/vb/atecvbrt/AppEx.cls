VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
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


Private mLibraryVersions As Collection
Private mStaticData As Collection
Private mAppPath As String
Private mAppExeName As String
Private mAppName As String
Private mAppVersion As String
Private mStaticFileName As String
Private mHomeDirectory As String
Private mInitOK As Boolean
Private mErrHelp As ErrHelper
Private mFileHelp As FileHelper
Private Const STATICS_SECTION As String = "STATICDATA"

Private Sub Class_Initialize()
  Set mLibraryVersions = New Collection
  Set mStaticData = New Collection
  Set mErrHelp = New ErrHelper
  Set mFileHelp = New FileHelper
  mInitOK = False
End Sub

Private Sub Class_Terminate()
  Set mLibraryVersions = Nothing
  Set mStaticData = Nothing
End Sub

Public Sub Initialize(ByVal AppVBG As IUnknown)
  Dim vbApp As VB.Global
  Dim s As String, s0 As String
     
  On Error GoTo CoreSetupEx_err
  Set vbApp = AppVBG
  Call GetAppSettings(mAppPath, mAppExeName, mAppVersion, mAppName, vbApp)
  mStaticFileName = mAppPath & "\" & mAppExeName & ".MSG"
  mHomeDirectory = FullPath(mAppPath)
  mInitOK = True
  
  Call AddStatic("Contact", CONTACTSTR)
  Call AddStatic("Version", , mAppVersion)
  Call AddStatic("ApplicationName", mAppName)
  s0 = GetStatic("ApplicationName")
  If Len(s0) > 0 Then
    mAppName = s0
  Else
    Call AddStatic("ApplicationName", , mAppName)
  End If
  Call ReadAllStatics
  Exit Sub
  
CoreSetupEx_err:
  mInitOK = False
  Err.Raise Err.Number, ErrorSourceEx(Err, "Initialise"), "Failed to initialise atecvbrt correctly" & vbCrLf & Err.Description
End Sub

Public Property Get LibraryVersions() As Collection
  Set LibraryVersions = mLibraryVersions
End Sub

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

Public Sub SaveStatics()
  Dim sItem As staticcls, ret As Boolean
  
  If Not mInitOK Then Err.Raise ERR_INITIALISE, "SaveStatics", "Cannot access atecvbrt.AppEx properties before calling Initialize"
  For Each sItem In mStaticData
    If sItem.bPersist Then
      ret = mFileHelp.WriteIniEntry(STATICS_SECTION, sItem.Name, CStr(sItem.Value), mStaticFileName)
    End If
  Next
End Sub

Public Sub ReadAllStatics()
  Dim Keys() As String, MaxKey As Long, sItem As staticcls
  Dim sValue As String
  Dim i As Long
  
  On Error GoTo ReadAllStatics_err
  
  MaxKey = mFileHelp.GetIniKeyNames(Keys, STATICS_SECTION, mStaticFileName)
  For i = 1 To MaxKey
    sValue = mFileHelp.GetIniEntry(STATICS_SECTION, Keys(i), "", mStaticFileName)
    If Not InCollectionEx(mStaticData, Keys(i)) Then
      Set sItem = New staticcls
      sItem.Name = Keys(i)
      sItem.Value = sValue
      sItem.bPersist = True
      mStaticData.Add sItem, Keys(i)
    End If
  Next i
  Exit Sub
  
ReadAllStatics_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "ReadAllStatics"), Err.Description
End Sub

'##D adds a string variable to the common static list maintained by the program, retrieved using getstatic
'##StaticName Name of the static to save
'##V DefaultValue default value (default="")
'##V OverrideValue override value (default="")
'##V Persistent is this a persistent static (saved in msg file) (default=true)
'##LD This is like the MSG file in Abacus - there is one copy per program and it is
'##LD called APPNAME.MSG
'##LD By default when a variable is change
'##C SF
Public Sub AddStatic(ByVal StaticName As String, Optional ByVal DefaultValue As Variant, Optional ByVal OverrideValue As Variant, Optional ByVal Persistent As Boolean = True)
  Dim sValue As String
  Dim sItem As staticcls, sItem2 As staticcls
  
  On Error GoTo AddStatic_err
  Set sItem = New staticcls
  StaticName = Trim$(StaticName)
  If Len(StaticName) = 0 Then Err.Raise ERR_STATIC, "AddStatic", "Cannot add a static with no name"
  sItem.Name = StaticName
  sItem.bPersist = Persistent
  
  If Not IsMissing(OverrideValue) Then
    sItem.Value = OverrideValue
  Else
    sItem.Value = GetStatic(StaticName, DefaultValue)
  End If
  If StrComp(StaticName, "ApplicationName", vbTextCompare) = 0 Then mAppName = sItem.Value
  If StrComp(StaticName, "Version", vbTextCompare) = 0 Then mAppVersion = sItem.Value
    
  If InCollectionEx(mStaticData, StaticName) Then
    Set sItem2 = mStaticData.Item(StaticName)
    sItem2.Name = sItem.Name
    sItem2.Value = sItem.Value
    sItem2.bPersist = sItem.bPersist
  Else
    mStaticData.Add sItem, StaticName
  End If
  Exit Sub
  
AddStatic_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "AddStatic"), Err.Description
End Sub

Public Function GetStatic(ByVal StaticName As String, Optional ByVal DefaultValue As Variant) As Variant
  Dim sItem As staticcls, sValue As String
  
  On Error GoTo GetStatic_err
  StaticName = Trim$(StaticName)
  Set sItem = mStaticData.Item(StaticName)
  GetStatic = sItem.Value
  Exit Function
  
GetStatic_err:
  If IsMissing(DefaultValue) Then
    sValue = mFileHelp.GetIniEntry(STATICS_SECTION, StaticName, "", mStaticFileName)
    If Len(sValue) = 0 Then Err.Raise ERR_STATIC, "GetStatic", "No value exists for named variable [" & StaticName & "]"
  Else
    sValue = mFileHelp.GetIniEntry(STATICS_SECTION, StaticName, DefaultValue, mStaticFileName)
  End If
  GetStatic = sValue
End Function

Public Property Get MsgFileName() As String
  If Not mInitOK Then Err.Raise ERR_INITIALISE, "MsgFileName", "Cannot access atecvbrt.AppEx properties before calling Initialize"
  MsgFileName = mStaticFileName
End Property

Public Property Get HomeDirectory() As String
  If Not mInitOK Then Err.Raise ERR_INITIALISE, "HomeDirectory", "Cannot access atecvbrt.AppEx properties before calling Initialize"
  HomeDirectory = mHomeDirectory
End Property

Public Property Let HomeDirectory(ByVal NewValue As String)
  If Not mInitOK Then Err.Raise ERR_INITIALISE, "HomeDirectory", "Cannot access atecvbrt.AppEx properties before calling Initialize"
  If Not FileExists(NewValue, True, False) Then Err.Raise 380, "HomeDirectory", "Home directory '" & LCase$(NewValue) & "' does not exist"
  mHomeDirectory = FullPath(NewValue)
  mStaticFileName = mHomeDirectory & mAppExeName & ".MSG"
End Property


'##D Errormessage function to repopulate the Err object for error messages recorded with ErrorMessagePush
'##C EMF
Public Sub ErrorMessagePopErr(ErrObj As ErrObject)
  'If ghInstance = 0 Then Err.Raise ERR_INITIALISE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  Call mErrHelp.Pop(ErrObj)
End Sub

'##D Records the current error whcih can then be displyed using ErrorMessagePop
'##LD This is required where your error handler calls functions which have error handlers themselves (On Error handler clears the current error information)
'##C EMF
Public Sub ErrorMessagePush(ErrObj As ErrObject)
  Call mErrHelp.Push(ErrObj)
End Sub

'##D Gets DOS path of program
'##RV None
'##C FNF
Public Property Get AppPathDos() As String
  Static sAppPathDos As String
  
  'If ghInstance = 0 Then Err.Raise ERR_INITIALISE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  If Not Len(sAppPathDos) Then
    sAppPathDos = GetDrivePathEx(mAppPath)
  End If
  AppPathDos = sAppPathDos
End Property


'##D Gets path of program
'##RV None
'##C FNF
Public Property Get AppPath() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITIALISE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppPath = mAppPath
End Property

'##D Application name
'##C SQF
Public Property Get AppName() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITIALISE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppName = mAppName
End Property

'##D Application executable name
'##C SQF
Public Property Get AppExeName() As String
  'If ghInstance = 0 Then Err.Raise ERR_INITIALISE, "GetStatic", "Cannot access TCSCORE properties before calling CoreSetup"
  AppExeName = mAppExeName
End Property

