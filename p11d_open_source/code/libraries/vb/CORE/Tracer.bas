Attribute VB_Name = "Tracer"
Option Explicit

'tracer

#If DEBUGVER Then
' Type LARGE_INTEGER
'   lowpart As Long
'   highpart As Long
' End Type

 Private m_TracerCount As Long
 Private Const Increment As Long = 256
 Private Const STACKMAX As Long = 8192
 Private m_CallStackTop As Long
 Private m_CallStackMax  As Long
 Private m_FunctionStack() As FunctionTrace
 
 Private m_SuspendCount As Long
 Private m_SuspendTime As Long
 Private m_SuspendDuration As Long
 
 Private m_FunctionList As New Collection
 
 'Private XSET_TracerOverhead As Long
 'Private XRETURN_TracerOverhead As Long
 

 ' Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
' Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
' Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
' Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

 'Private Perf_Freq As Double
 'Private Perf_MAXLONG  As Long
#End If

Public Sub Tracer_XSet(ByVal sFunctionName As String)
#If DEBUGVER Then
  Dim fntrace As FunctionTrace
    
  If m_SuspendCount > 0 Then Exit Sub
  m_TracerCount = m_TracerCount + 1
  
  sFunctionName = UCase$(sFunctionName)
  Call addfunction(sFunctionName)
  
  Set fntrace = Push(sFunctionName)
  If fntrace.NestCount = 0 Then fntrace.time0 = timeGetTime
  
  If mCoreTrace Then
    Call logfunction3(Now, 0, "DEBUG", "xSet   " & Space$(m_CallStackTop * 2) & sFunctionName, "XSet", ".TR")
  End If
#End If
End Sub

Public Sub Tracer_XReturn(ByVal sFunctionName As String)
#If DEBUGVER Then
  Dim t0 As Long, dtime As Long, TrackBack As Boolean
  Dim fntrace As FunctionTrace
  Dim fnitem As FunctionItem
  
  On Error GoTo Tracer_XReturn_err
  
  If m_SuspendCount > 0 Then Exit Sub
  t0 = timeGetTime
  TrackBack = True
  sFunctionName = UCase$(sFunctionName)
  If mCoreTrace Then
    Call logfunction3(Now, 0, "DEBUG", "xReturn" & Space$(m_CallStackTop * 2) & sFunctionName, "XReturn", ".TR")
  End If
 
Tracer_XReturn_again:
  Set fntrace = Pop
  If fntrace.NestCount >= 0 Then Exit Sub
  
  ' Set without Return
  If StrComp(fntrace.Name, sFunctionName) <> 0 Then
    If m_SuspendDuration > 0 Then GoTo Tracer_XReturn_again
  End If
  dtime = (t0 - fntrace.time0) - m_SuspendDuration
  If dtime < 0 Then
    dtime = (t0 - fntrace.time0)
    TrackBack = False
  End If
  m_SuspendDuration = 0#
  
  Set fnitem = m_FunctionList.Item(fntrace.Name)
  fnitem.DeltaTime = fnitem.DeltaTime + dtime
  
  fnitem.TotalTime = fnitem.TotalTime + dtime
  
  ' Take account of delta time
  If (m_CallStackTop > 0) And TrackBack Then
    Set fntrace = m_FunctionStack(m_CallStackTop)
    Set fnitem = m_FunctionList.Item(fntrace.Name)
    fnitem.DeltaTime = fnitem.DeltaTime - dtime
  End If
  Set fnitem = Nothing
  Set fntrace = Nothing

Tracer_XReturn_end:
  Exit Sub
  
Tracer_XReturn_err:
  Call ErrorMessageEx(ERR_ERROR, Err, "Tracer_XReturn", "ERR_TRACER", "Error in Function XReturn called from  Function " & sFunctionName, False)
  Resume Tracer_XReturn_end
#End If
End Sub


Public Sub Tracer_report()
#If DEBUGVER Then
  Dim iFileNum As Integer, dt As Double, tt As Double, d0 As Double
  Dim sFilePath As String, s As String
  Dim fnitem As FunctionItem

  On Error GoTo Tracer_report_err
  iFileNum = FreeFile
  sFilePath = mAppPath & "\" & mAppExeName & ".T"
  Open sFilePath For Output As iFileNum
    Print #iFileNum, "Program: " & mAppExeName & "  Version: " & mAppVersion & " Date: " & Format$(Now, "dd/mm/yyyy hh:nn")
    Print #iFileNum, "Name" & vbTab & "Count" & vbTab & "Delta Time" & vbTab & "Unit Delta Time" & vbTab & "Total Time"
    For Each fnitem In m_FunctionList
      dt = fnitem.DeltaTime / 1000
      tt = fnitem.TotalTime / 1000
      If fnitem.CallCount > 0 Then
        d0 = dt / fnitem.CallCount
      Else
        d0 = 0#
      End If
      s = fnitem.Name & vbTab & fnitem.CallCount & vbTab & Format$(dt, "#,###0.0000") & vbTab & Format$(d0, "#,###0.0000") & vbTab & Format$(tt, "#,###0.0000")
      Print #iFileNum, s
    Next fnitem
    Print #iFileNum, "Trace Count" & vbTab & m_TracerCount
  
Tracer_report_end:
  If iFileNum > 0 Then Close #iFileNum
  Exit Sub
  
Tracer_report_err:
  Resume Tracer_report_end
#End If
End Sub

Public Sub Tracer_suspend()
#If DEBUGVER Then
  If m_SuspendCount = 0 Then
    m_SuspendTime = timeGetTime
  End If
  m_SuspendCount = m_SuspendCount + 1
#End If
End Sub

Public Sub Tracer_restart()
#If DEBUGVER Then
  m_SuspendCount = m_SuspendCount - 1
  If m_SuspendCount = 0 Then
    m_SuspendDuration = m_SuspendDuration + (timeGetTime - m_SuspendTime)
  End If
#End If
End Sub

Public Function Tracer_Cleanup(ByVal ClearFunctionList As Boolean)
#If DEBUGVER Then
  Dim i As Long
  On Error Resume Next
  
  For i = 1 To m_CallStackMax
    Set m_FunctionStack(i) = Nothing
  Next i
  m_SuspendCount = 0
  m_CallStackTop = 0
  If ClearFunctionList Then Set m_FunctionList = Nothing
#End If
End Function


#If DEBUGVER Then
Private Sub addfunction(fnname As String)
  Dim i As Integer
  Dim fnitem As FunctionItem
    
  On Error Resume Next
  Set fnitem = m_FunctionList.Item(fnname)
  If fnitem Is Nothing Then
    Set fnitem = New FunctionItem
    fnitem.Name = fnname
    Call m_FunctionList.Add(fnitem, fnname)
  End If
  fnitem.CallCount = fnitem.CallCount + 1
End Sub
#End If

#If DEBUGVER Then
Private Function DumpStack()
  Dim fntrace As FunctionTrace
  Dim i As Long
  
  On Error Resume Next
  Call logfunction3(Now, 0, "DEBUG", "Stack top: " & m_CallStackTop & " Stack max: " & m_CallStackMax, "StackDump", ".RIP")
  i = m_CallStackMax - 1024
  If i < 1 Then i = 1
  For i = i To m_CallStackMax
    Set fntrace = m_FunctionStack(i)
    If Not fntrace Is Nothing Then Call logfunction3(Now, 0, "DEBUG", Format$(i, "0000") & ": " & fntrace.Name, "StackDump", ".RIP")
  Next i
  Call Tracer_Cleanup(False)
End Function

Private Function Push(FName As String) As FunctionTrace
  
  If m_CallStackTop > 0 Then
    If StrComp(m_FunctionStack(m_CallStackTop).Name, FName) = 0 Then
      Set Push = m_FunctionStack(m_CallStackTop)
      Push.NestCount = Push.NestCount + 1
      Exit Function
    End If
  End If
  m_CallStackTop = m_CallStackTop + 1
  If m_CallStackTop < 1 Then m_CallStackTop = 1
  If m_CallStackTop > m_CallStackMax Then
    If (m_CallStackMax + Increment) > STACKMAX Then
      Call DumpStack
      Call ErrorMessageEx(ERR_ERROR, Err, "Tracer Push", "Stack Trace", "Call stack too large, > " & CStr(STACKMAX) & " values", False)
      m_CallStackTop = 1
    Else
      m_CallStackMax = m_CallStackMax + Increment
      If IsArrayEx2(m_FunctionStack) Then
        ReDim Preserve m_FunctionStack(1 To m_CallStackMax) As FunctionTrace
      Else
        ReDim m_FunctionStack(1 To m_CallStackMax) As FunctionTrace
      End If
    End If
  End If
  Set Push = New FunctionTrace
  Push.Name = FName
  Push.StackIndex = m_CallStackTop
  Set m_FunctionStack(m_CallStackTop) = Push
End Function

Private Function Pop() As FunctionTrace
  m_CallStackTop = m_CallStackTop - 1
  If m_CallStackTop < 0 Then Err.Raise ERR_CALLSTACKCORRUPT, "Tracer_POP", "Pop without Push, no XSet for XReturn"
  Set Pop = m_FunctionStack(m_CallStackTop + 1)
  Pop.NestCount = Pop.NestCount - 1
  If Pop.NestCount >= 0 Then
    m_CallStackTop = m_CallStackTop + 1
  Else
    Set m_FunctionStack(m_CallStackTop + 1) = Nothing
  End If
End Function

#End If

Public Sub Tracer_FillList(lst As ListBox)
#If DEBUGVER Then
  Dim i As Long
  Dim fntrace As FunctionTrace
  
  For i = m_CallStackTop To 1 Step -1
    Set fntrace = m_FunctionStack(i)
    If Not fntrace Is Nothing Then lst.AddItem fntrace.Name
  Next i
#End If
End Sub
