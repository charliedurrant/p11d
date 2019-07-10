Attribute VB_Name = "TimerCode"
Option Explicit
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private TimerActive As Boolean
Private mTimers As Collection

Public Sub StartTimerEx(ByVal Interval As Long, ByVal tProc As ITimerProc)
  Dim ti As TimerHelp
  
  On Error GoTo StartTimerEx_err
  If mTimers Is Nothing Then Set mTimers = New Collection
  Set ti = New TimerHelp
  Set ti.tProc = tProc
  If ti.tProc Is Nothing Then Call Err.Raise(ERR_STARTTIMER, "StartTimerEx", "ITimer interface invalid")
  If IsRunningInIDEEx Then
    Debug.Print "Call to StartTimer disabled in IDE"
  Else
    ti.TimerID = SetTimer(0&, 0&, Interval, AddressOf xTimerProc)
  End If
  If ti.TimerID = 0 Then Call Err.Raise(ERR_STARTTIMER, "StartTimerEx", "Unable to setup timer")
  TimerActive = True
  Call mTimers.Add(ti)
  
StartTimerEx_end:
  Exit Sub
  
StartTimerEx_err:
  If Not IsRunningInIDEEx Then Call ErrorMessageEx(ERR_ERROR, Err, "StartTimerEx", "Start Timer Function", "Unable to initialise timer", False)
  Resume StartTimerEx_end
End Sub

Public Sub KillTimerEx(ByVal itproc As ITimerProc)
  Dim i As Long
  Dim ti As TimerHelp
  
  On Error Resume Next
  For i = 1 To mTimers.Count
    Set ti = mTimers.Item(i)
    If ti.tProc Is itproc Then
      Call KillTimer(0&, ti.TimerID)
      Call mTimers.Remove(i)
      Set ti = Nothing
      Exit Sub
    End If
  Next i
End Sub

Public Sub KillAllTimers()
  Dim i As Long
  Dim ti As TimerHelp
  
  On Error Resume Next
  TimerActive = True
  If Not mTimers Is Nothing Then
    For i = mTimers.Count To 1
      Set ti = mTimers.Item(1)
      Call KillTimerEx(ti.tProc)
      Set ti = Nothing
    Next i
  End If
  Set mTimers = Nothing
End Sub

Private Sub xTimerProc(ByVal hWnd As Long, ByVal lMsg As Long, ByVal TimerID As Long, ByVal lTimer As Long)
  Dim ti As TimerHelp
  
  On Error Resume Next
  If Not TimerActive Then Exit Sub
  For Each ti In mTimers
    If ti.TimerID = TimerID Then
      If ti.inTimer Then Exit Sub
      ti.inTimer = True
      Call ti.tProc.OnTimer(lTimer)
      ti.inTimer = False
      Exit Sub
    End If
  Next ti
End Sub
