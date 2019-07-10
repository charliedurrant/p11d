VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6600
   ClientLeft      =   270
   ClientTop       =   840
   ClientWidth     =   9585
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExitBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAboutSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F11}
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'AUTOINSERTCOMMENT BEGIN DECLARE EVENTS
Private WithEvents clsEvent As TCSEventClass
Attribute clsEvent.VB_VarHelpID = -1
'AUTOINSERTCOMMENT END DECLARE EVENTS

Private Sub clsEvent_DebugMenuItem(Name As String, Index As Long, Parent As TCSCOREDLL.TCSMenuItems)
On Error GoTo clsEvent_DebugMenuItem_err
  Select Case Parent
  Case MNU_BREAK
    Call ExitApp(True)
  Case Else
    ECASE " clsEvent_DebugMenuItem"
  End Select
clsEvent_DebugMenuItem_end:
  Exit Sub
clsEvent_DebugMenuItem_err:
  Call ErrorMessage(ERR_ERROR, Err, "DebugMenuItem", "ERR_DEBUGMENU", "Error processing the debug menu event " & Name & ".")
  Resume clsEvent_DebugMenuItem_end
End Sub

Private Sub MDIForm_Load()
'AUTOINSERTCOMMENT BEGIN MDI LOAD
  Set clsEvent = gEvents
'AUTOINSERTCOMMENT END MDI LOAD
End Sub

' AUTOINSERTCOMMENT BEGIN QUERYUNLOAD
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  gbAllowAppExit = True
  
    Select Case UnloadMode
      Case vbFormCode
        If gbForceExit Then
          Call UserAppShutDown
          Call ExitApp
        Else
          If (MsgBox("Are you sure you want to exit " & AppName & "?", vbYesNo, AppName)) = vbYes Then
            Call UserAppShutDown
            Call ExitApp
          Else
            Cancel = True
            gbAllowAppExit = False
          End If
        End If
      Case vbAppTaskManager
        Call UserAppShutDown
        Call ExitApp
      Case Else
        If (MsgBox("Are you sure you want to exit " & AppName & "?", vbYesNo, AppName)) = vbYes Then
          Call UserAppShutDown
          Call ExitApp
        Else
          Cancel = True
          gbAllowAppExit = False
        End If
    End Select
  
End Sub
' AUTOINSERTCOMMENT END QUERYUNLOAD


Private Sub mnuExit_Click()
  Call ExitApp
End Sub

Private Sub mnuHelpAbout_Click()
  Call AppAbout
End Sub

