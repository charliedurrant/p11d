Attribute VB_Name = "Start"
Option Explicit
'************************************************************************
'* This is the main module for your file.  Do not add code here, only   *
'* call functions in other modules.                                     *
'* Always use exitapp to terminate your application not END.            *
'************************************************************************
Public Sub Main()
  Dim AppName As String
  
  On Error GoTo main_nocore_err
  AppName = App.EXEName
  Set gTBCtrl = New ToolBarControl
  Splash.Message = "Initialising - please wait"
  Splash.HideProgressBar = True
  Splash.InitProgressBar
  Splash.Show
  Splash.Refresh
  If Not CoreSetup(Command$(), VB.Global) Then Err.Raise ERR_CORESETUP, "Main", "Unable to initialise Core library"
  Call Splash.IncrementProgressBar
  
  On Error GoTo main_err
  FormattedErrorStrings = True
  AppName = GetStatic("ApplicationName")
  frmMain.Caption = AppName & " Version " & GetStatic("Version")
  Call Splash.IncrementProgressBar
  Call gTBCtrl.Initialise(frmMain.TB_Main, frmMain.TB_ImageList_Normal, frmMain.TB_ImageList_Hot, Nothing)
  Call Splash.IncrementProgressBar
  Call gTBCtrl.RefreshToolbar
  Call Splash.IncrementProgressBar(True)
  frmMain.Show
  Call UnLoadSplash
  
  '** The rest of your standard initialisation here if error goto main_end
  ' Call xmkdir("C:\TEST\TEST1\TEST2")
  ' Call xmkdir("C:\TEST1\TEST2\TEST3\")
  'CAD
  'Call InitSQLExplorer(New SQLDebug)
  Call RegisterDB(Nothing)
    
main_clean_end:
  ErrorOtherButtonCaption = "Hello world"
  ' Call ErrorMessage(ERR_ERROR + ERR_ALLOWOTHER, Nothing, "Main", "No title", "This is a test <B>please do not delete</B><I> hello </I>this is a very long error message and is written to test the wrapping on certain types of error display")
  ErrorOtherButtonCaption = ""
  Exit Sub

main_end:
  Call UnLoadSplash
  Call ExitApp(True)
  Exit Sub
  
main_err:
  Call ErrorMessage(ERR_ERROR, Err, "Main", "Fatal Error in " & AppName, "There was an Error in the main module of this Application")
  Call UnLoadSplash
  Resume main_end
  
main_nocore_err:
  MsgBox "There was an Error in the main module of this Application" & vbCrLf & "Error(" & Err.Number & "): " & Err.Description, vbCritical + vbOKOnly, "Fatal Error in " & AppName
  Resume main_end
  Resume
End Sub

Private Sub UnLoadSplash()
  On Error Resume Next
  Unload Splash
  Set Splash = Nothing
End Sub

Public Sub ExitApp(Optional bForce As Boolean = False)
  Static inexitapp As Boolean
  
  On Error Resume Next
  If (Not inexitapp) Then
    inexitapp = True
    gbAllowAppExit = True
    gbForceExit = bForce
    Call UnLoadSplash
    If gbMDILoaded Then
      Unload frmMain     ' takes account of FatalError
    Else
      Call UserAppShutDown
    End If
    If gbAllowAppExit Then
      Call CoreShutDown
      Set frmMain = Nothing
    Else
      gbForceExit = False
      inexitapp = False
    End If
  End If
End Sub

' Use this function to clean up App specific items
Public Function UserAppShutDown() As Boolean
  Dim ShutDownOk As Boolean
  
  On Error Resume Next
  ShutDownOk = True
  'user app specific here set ShutDownOk variable
  If ShutDownOk Then
    Call InitSQLExplorer(Nothing)
    Call CloseAllForms(Forms, True)
    UserAppShutDown = True
  End If
End Function

