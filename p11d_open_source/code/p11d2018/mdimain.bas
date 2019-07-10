Attribute VB_Name = "MDIModule"
 
Option Explicit

Public Sub EnterSerialNumber()
  On Error GoTo EnterSerialNumber_Err
  
  Call xSet("EnterSerialNumber")
  If F_SerialNumber.Start Then
  End If
  Set F_SerialNumber = Nothing

EnterSerialNumber_End:
  Call xReturn("EnterSerialNumber")
  Exit Sub
EnterSerialNumber_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EnterSerialNumber", "Error in EnterSerialNumber", "Undefined error.")
  Resume EnterSerialNumber_End
End Sub

Public Sub ExitApp(Optional bForce As Boolean = False)
  Static inexitapp As Boolean
  If (Not inexitapp) Then
    inexitapp = True
    gbAllowAppExit = True
    gbForceExit = bForce
    Call UnLoadSplash
    If app.StartMode = vbSModeStandalone Then
      If Not fatalError Then Unload MDIMain
    Else
      Call UserAppShutDown
    End If
    If gbAllowAppExit Then
      Call CoreShutDown
      Set gPreAlloc = Nothing
      If fatalError Then End
      Set MDIMain = Nothing
    End If
    gbForceExit = False
    inexitapp = False
  End If
End Sub

Public Function GetAppYear() As Long
  Dim s As String
  Dim i As Long
  
  On Error GoTo GetAppYear_ERR
  
  GetAppYear = -1
  If Len(app.EXEName) < 4 Then Call Err.Raise(ERR_APP_NAME, "AppYear", "App EXE name less than 2 characters.")
  s = Right$(app.EXEName, 4)
  If Not IsNumeric(s) Then Call Err.Raise(ERR_APP_NAME, "AppYear", "Right 2 chars of App EXE name are not numeric.")
  GetAppYear = CLng(s)
GetAppYear_END:
  Exit Function
GetAppYear_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetAppYear", "Get App Year", "Error getting the appyear.")
  Resume GetAppYear_END
  Resume
End Function

Public Function GetAppYearShort(ByVal AppYear As Long) As String
  Dim l As Long
  Dim s As String
  
  l = AppYear - 2000
  If l > 9 Then
    s = CStr(l)
  Else
    s = "0" & CStr(l)
  End If
  GetAppYearShort = s
  
End Function

Public Function GetVersionString(ByVal bCoreInfo As Boolean) As String
  Dim s As String
  Static bStaticSet As Boolean
  On Error Resume Next
  s = app.major & "." & app.minor & "." & app.Revision
  If bCoreInfo Then
    If TCSCoreDebug Then s = s & " (Debug)"
    If (Not bStaticSet) Then
     Call AddStatic("Version", s, s, False)
     bStaticSet = True
    End If
  End If
  GetVersionString = s
End Function

Public Function IsFormLoaded(sFormName As String) As Boolean
'RK This functionality may be in the library, but cannot find it
  On Error GoTo IsFormLoaded_Err
  Call xSet("IsFormLoaded")
  Dim lngCount As Long
  
  For lngCount = 0 To Forms.Count - 1
    If Forms(lngCount).Name = sFormName Then
       IsFormLoaded = True
       Exit Function
    End If
  Next lngCount
  
  IsFormLoaded = False
  
IsFormLoaded_End:
  Call xReturn("IsFormLoaded")
  Exit Function
IsFormLoaded_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsFormLoaded", "Error in IsFormLoaded", "Undefined error.")
  Resume IsFormLoaded_End
End Function

Sub Main()
  Dim AppName As String, Version As String
  ' defaults only
  Dim s As String
  
  On Error GoTo main_err
  
  AppName = UCASE$(app.EXEName)
  Version = GetVersionString(False)
  
  
  'apf cd  if app fails to load please give errormessage
  If Not CoreSetup(Command$(), VB.Global, False) Then Call ExitApp(True)
  
  s = AppName
  Call AddStatic("ApplicationName", s, s, True)
  s = "For help, please contact " & app.companyName & " on " & S_TELEPHONE
  Call AddStatic("Contact", s, s, True)
  
  If app.StartMode = vbSModeStandalone Then
    'move bitmap screen
    Call SplashScreens
  End If
  DatabaseTarget = DB_TARGET_JET
  
  
  Set sql = New SQLQUERIES
  
  Set gPreAlloc = New PreAllocate
  
  Call gPreAlloc.AllocObjects(PREALLOC_PARSER + PREALLOC_REP + PREALLOC_AUTO)
  
  Set p11d32 = New p11d32
  
  If Not p11d32.Initialise Then GoTo main_end
  
  Version = GetVersionString(True)
  'SerialNumber checked on initsettings RK will not unload properly
  If Not CBoolean(p11d32.LicenceType) Then
     Set F_SerialNumber = Nothing
     GoTo main_end
  End If
  
  If app.StartMode = vbSModeStandalone Then
    'MDIMain.Show
    Call p11d32.JuneReleaseWarning
    Call p11d32.Help.ShowForm(MDIMain)
    Call UnLoadSplash
    DoEvents
    Call p11d32.StartScreen
  End If

  Exit Sub

main_end:
  Call UnLoadSplash
  Call ExitApp(True)
  Exit Sub
main_err:
  Call ErrorMessagePush(Err)
  Call ErrorMessagePop(ERR_ERROR, Err, "Main", "Error Starting " & AppName & " Version " & Version, "Startup initialisation failed")
  Resume main_end
  Resume
End Sub

Private Sub SplashScreens()
  On Error GoTo SplashScreens_ERR
  
  Exit Sub
'    frmSplash.Show
    Call p11d32.Help.ShowForm(frmSplash)
    DoEvents
    If Not IsRunningInIDE() Then
      Call Sleep(2000)
    End If
   frmSplash.Hide
  
SplashScreens_END:
  Exit Sub
SplashScreens_ERR:
  Resume SplashScreens_END
  Resume
End Sub

Private Sub UnLoadSplash()
  On Error Resume Next
  If app.StartMode = vbSModeStandalone Then
    Unload frmSplash
    Set frmSplash = Nothing
  End If
End Sub

Public Function UserAppShutDown() As Boolean
  Dim ShutDownOk As Boolean
  On Error Resume Next
  
  'user app specific here
  ShutDownOk = True
  If Not p11d32.CurrentEmployer Is Nothing And (app.StartMode = vbSModeStandalone) Then
    ShutDownOk = p11d32.MoveMenuUpdateEmployeeCheckEmployer
  End If
  If ShutDownOk Then
    If (Not p11d32 Is Nothing) Then Call p11d32.Kill
    Set p11d32 = Nothing
    If app.StartMode = vbSModeStandalone Then
      Call SetLastListItemSelected(Nothing)
      Call CloseAllForms(Forms, True)
    End If
    UserAppShutDown = True
  End If
End Function

