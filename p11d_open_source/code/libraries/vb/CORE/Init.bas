Attribute VB_Name = "Init"
Option Explicit

Private Function GetAppSettings(mAppPath As String, mAppExeName As String, mAppVersion As String, mAppName As String) As Boolean
  Dim p As Long, AppFile As String
      
  If Len(mAppPath) = 0 Then
    If Not vbg Is Nothing Then mAppPath = vbg.App.Path
  End If
  If vbg Is Nothing Then
    AppFile = GetModuleName(0, True)
    Call SplitPathEx(AppFile, mAppPath, mAppExeName)
  Else
    mAppPath = vbg.App.Path
    mAppExeName = vbg.App.EXEName
  End If
  If right$(mAppPath, 1) = "\" Then mAppPath = left$(mAppPath, Len(mAppPath) - 1)
  p = InStr(1, mAppExeName, ".EXE", vbTextCompare)
  If p > 1 Then mAppExeName = left$(mAppExeName, p - 1)
      
  If vbg Is Nothing Then
    mAppVersion = VersionQueryMap(AppFile, VQT_FILE_VERSION)
    mAppName = mAppExeName
  Else
    #If DEBUGVER Then
      mAppVersion = vbg.App.Major & "." & vbg.App.Minor & "." & vbg.App.Revision & " Debug"
    #Else
      mAppVersion = vbg.App.Major & "." & vbg.App.Minor & "." & vbg.App.Revision
    #End If
    mAppName = vbg.App.Title
  End If
  
  If (Len(mAppVersion) = 0) Or (Len(mAppName) = 0) Then Call Err.Raise(ERR_INITCORE, Err, "GetAppSettings", "Unable to initialise core library. No AppVersion or AppName found")
  GetAppSettings = True
End Function
    
Public Function GetCmdParamEx(ByVal Param As Variant, Optional buffer As Variant, Optional ByVal vCmd As Variant, Optional ByVal RemoveParam As Boolean = False, Optional ByRef OutCMD As String) As Boolean
  Dim sCmd As String, sNext As String
  Dim sBuffer As String, sOutBuffer As String
  Dim p0 As Long, p1 As Long
  Dim InESC As Boolean
  Dim i As Long
  
  On Error GoTo GetCmdParamEx_err
  p0 = 1
  If IsMissing(vCmd) Then
    sCmd = Trim$(mAppCmdParam)
  Else
    sCmd = Trim$(vCmd)
  End If
  If Len(sCmd) = 0 Then GoTo GetCmdParamEx_end
  If VarType(Param) = vbString Then
    If Len(Param) = 0 Then GoTo GetCmdParamEx_end
next_param:
    p0 = InStr(p0, sCmd, Param, vbTextCompare)
    If p0 > 0 Then
      p1 = p0 + Len(Param)
      sNext = Mid$(sCmd, p1, 1)
      GetCmdParamEx = ((sNext = " ") Or (sNext = "=") Or (sNext = ""))
      If Not GetCmdParamEx Then
        p0 = p1
        GoTo next_param
      End If
    End If
  End If
  If (VarType(Param) = vbInteger) Or (VarType(Param) = vbLong) Then
    p0 = 1: i = 1: InESC = False
    Do While (i <> Param) And (p0 <= Len(sCmd))
      sNext = Mid$(sCmd, p0, 1)
      If sNext = """" Then
        InESC = Not InESC
      ElseIf (sNext = " ") And Not InESC Then
        p0 = NotInStrAny(sCmd, " ", p0, vbBinaryCompare) - 1
        i = i + 1
      End If
      p0 = p0 + 1
    Loop
    GetCmdParamEx = (i = Param)
  End If
  If GetCmdParamEx And (Not IsMissing(buffer)) Then
    ' p0 start of param
    p1 = InStr(p0, sCmd, "=")
    
    If (p1 = 0) And Not RemoveParam Then Err.Raise ERR_CMDPARAM, "GetCmdParamEx", "Expected to find /" & Param & "=Value found " & sCmd
    sBuffer = Trim$(Mid$(sCmd, p1 + 1))
    p1 = 1: InESC = False
    Do
      sNext = Mid$(sBuffer, p1, 1)
      If sNext = """" Then
        InESC = Not InESC
      ElseIf (sNext = " ") And Not InESC Then
        Exit Do
      Else
        sOutBuffer = sOutBuffer & sNext
      End If
      p1 = p1 + 1
    Loop Until p1 > Len(sBuffer)
    p1 = p1 + InStr(p0, sCmd, "=")
    buffer = sOutBuffer
  End If
  
GetCmdParamEx_end:
  If RemoveParam And (p0 > 0) And (p1 > p0) Then OutCMD = left$(sCmd, p0 - 1) & Mid$(sCmd, p1)
  Exit Function
  
GetCmdParamEx_err:
  GetCmdParamEx = False
  Call ErrorMessageEx(ERR_ERROR, Err, "GetCmdParamEx", "Unable to parse command line arguements", "Error parsing command line arguements", False)
  Resume GetCmdParamEx_end
  Resume
End Function

Public Function CoreSetupEx(CommandLine As String, ByVal AppVBG As IUnknown, ByVal DatabaseAccess As Boolean, ByVal DLLSetup As Boolean, ByVal bCoreFirst As Boolean) As Boolean
  Dim s As String, s0 As String
     
  On Error GoTo CoreSetupEx_err
  If (mCoreInitCount > 0) And Not bCoreFirst Then
    CoreSetupEx = True
    GoTo CoreSetupEx_end
  End If
  CoreSetupEx = False
  
  Set vbg = AppVBG
  If Not vbg Is Nothing Then
    ghInstance = vbg.App.hInstance
    ghThreadID = vbg.App.ThreadID
  Else
    ghInstance = App.hInstance
    ghThreadID = App.ThreadID
  End If
  gTCSEventClass.Name = "INTERNAL TCS"
  gPasswordTitle = "QUERY_PASSWORD"
  gPasswordPrompt = "Warning - You are about to enter a system function."
  mDBTarget = DB_TARGET_JET
  mTCS_InitialiseDefaultWS = DatabaseAccess
  mAppCmdParam = Trim$(CommandLine)
  Call TimerFrequency(True)
  #If DEBUGVER Then
    mCoreTrace = GetCmdParamEx("/TRACE", , mAppCmdParam)
    mTCSCoreVersion = App.Major & "." & App.Minor & "." & App.Revision & " Debug"
  #Else
    mCoreTrace = False
    mTCSCoreVersion = App.Major & "." & App.Minor & "." & App.Revision
  #End If
  
  Call GetCmdParamEx("/TRACE", , mAppCmdParam, True, mAppCmdParam)
  'APF NOT REQUIRED mAppCmdParam = sRemCmdParam(mAppCmdParam, "/ZHOME")
  'mAppCmdParam = sRemCmdParam(mAppCmdParam, "/TRACE")
  
  If GetAppSettings(mAppPath, mAppExeName, mAppVersion, mAppName) Then
    mStaticFileName = mAppPath & "\" & mAppExeName & ".MSG"
    mHomeDirectory = FullPathEx(mAppPath)
    Call AddStaticEx("Contact", contactstr)
    Call AddStaticEx("Version", , mAppVersion)
    Call AddStaticEx("ApplicationName", mAppName)
    s0 = GetStaticEx("ApplicationName")
    If Len(s0) > 0 Then
      mAppName = s0
    Else
      Call AddStaticEx("ApplicationName", , mAppName)
    End If
    Call ReadAllStatics
    mSilentError = False
    If Not DLLSetup Then Call InitKeyboardHook
    CoreSetupEx = True
  End If
  
CoreSetupEx_end:
  'MsgBox hInstance & vbCrLf & mAppName & vbCrLf & mAppPath
  Exit Function
  
CoreSetupEx_err:
  Call MsgBox("Unable to initialise the core library correctly." & vbCr & vbCr & contactstr, vbCritical + vbOKOnly, "Error in coresetup")
  CoreSetupEx = False
  Resume CoreSetupEx_end
End Function


