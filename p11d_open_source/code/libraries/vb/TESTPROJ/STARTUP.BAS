Attribute VB_Name = "main"
Option Explicit
'This is the main module for your file.  Do not add code here, only
'call functions in other modules.
'Lines should not be inserted or ammended between AUTOINSERTCOMMENT BEGIN
'and AUTOINSERTCOMMENT END
'Always use exitapp to terminate your application not END.
Sub main()
  
  frmSplash.Message = "Initialising - please wait"
  
' AUTOINSERTCOMMENT BEGIN SPASHSHOW
  frmSplash.Show
  DoEvents
' AUTOINSERTCOMMENT END SPLASHSHOW
  
' AUTOINSERTCOMMENT BEGIN CORESETUP
  If Not CoreSetup(Command$(), App) Then
    Call ExitApp(True)
  End If
' AUTOINSERTCOMMENT END CORESETUP
  
' AUTOINSERTCOMMENT BEGIN MDISHOW
  MDImain.Show
' AUTOINSERTCOMMENT END MDISHOW
  
  'the rest of your initialisation here

main_clean_end:
' AUTOINSERTCOMMENT BEGIN UNLOADSPASH
  Unload frmSplash
  Set frmSplash = Nothing
  Exit Sub
' AUTOINSERTCOMMENT END UNLOADSPASH
' AUTOINSERTCOMMENT BEGIN MAINERROR2
main_end:
  Call ExitApp(True)
  Exit Sub
  
main_err:
  Call ErrorMessage(ERR_ERROR, Err, "Main", "ERR_UNDEFINED", "There was an undefined error in this application")
  Resume main_end
  ' AUTOINSERTCOMMENT END MAINERROR2
End Sub

' Use this function to clean up App specific items
Sub UserAppShutDown()
  Dim frm As Form
  Dim i As Long
  On Error Resume Next
  
  For i = 1 To Forms.Count
    Set frm = Forms(i)
    Unload frm
    Set frm = Nothing
  Next i
End Sub

' AUTOINSERTCOMMENT BEGIN EXITAPP
Sub ExitApp(Optional bForce As Boolean = False)
  Static inexitapp As Boolean
  If (Not inexitapp) Then
    inexitapp = True
    gbAllowAppExit = True
    gbForceExit = bForce
    Unload MDImain
    If gbAllowAppExit Then
      Call CoreShutDown
      Set MDImain = Nothing
      '* if all else fails End
    End If
    gbForceExit = False
    inexitapp = False
  End If
End Sub
' AUTOINSERTCOMMENT END EXITAPP
