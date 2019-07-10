Attribute VB_Name = "HOOKFuncs"
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC As Long = (-4)

Private Type HOOKINFO
  hWnd As Long        'Subclassed window
  Ctrl As HOOK
  OldWndProc As Long  'Old window procedure
End Type

Private m_HookArray() As HOOKINFO
Private m_MaxHooks As Long
Private Const INCREMENT As Long = 128

Public Sub Main()
  Call ExpandHooks
End Sub

Private Sub ExpandHooks()
  m_MaxHooks = m_MaxHooks + INCREMENT
  ReDim Preserve m_HookArray(1 To m_MaxHooks)
End Sub

Private Sub ClearHookInfo(ByVal Index As Long)
  m_HookArray(Index).OldWndProc = 0
  Set m_HookArray(Index).Ctrl = Nothing
  m_HookArray(Index).hWnd = 0
End Sub

Public Sub HookWindow(ByVal hWnd As Long, ByVal Ctrl As HOOK)
  Static InHookWindow As Boolean
  Dim i As Long
  
  If InHookWindow Then Err.Raise 380, "HookWindow", "HookWindow is not reentrant"
  On Error Resume Next
  InHookWindow = True
  If hWnd <> 0 Then
    Call UnhookWindow(hWnd)
    
    For i = 1 To m_MaxHooks
      If m_HookArray(i).hWnd = 0 Then Exit For
    Next i
    If i > m_MaxHooks Then Call ExpandHooks
    m_HookArray(i).hWnd = hWnd
    Set m_HookArray(i).Ctrl = Ctrl
    m_HookArray(i).OldWndProc = GetWindowLong(hWnd, GWL_WNDPROC)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
  End If
  InHookWindow = False
End Sub

Public Sub UnhookWindow(ByVal hWnd As Long)
  Dim i As Long, j As Long
  
  For i = 1 To m_MaxHooks
    If m_HookArray(i).hWnd = hWnd Then
      Call SetWindowLong(hWnd, GWL_WNDPROC, m_HookArray(i).OldWndProc)
      Call ClearHookInfo(i)
      Exit For
    End If
  Next i
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim i As Long, mResult As Long
  
  On Error Resume Next   ' We may get an error from app so suppress
  For i = 1 To m_MaxHooks
    If m_HookArray(i).hWnd = hWnd Then
      If m_HookArray(i).Ctrl.Messages(Msg) Then
        If m_HookArray(i).Ctrl.RaiseWndProc(mResult, Msg, wParam, lParam) Then
          WndProc = mResult
          Exit Function
        End If
      End If
      Exit For
    End If
  Next i
  WndProc = CallWindowProc(m_HookArray(i).OldWndProc, hWnd, Msg, wParam, lParam)
End Function

