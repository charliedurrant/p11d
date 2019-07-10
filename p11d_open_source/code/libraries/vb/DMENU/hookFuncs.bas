Attribute VB_Name = "hookFuncs"
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC As Long = (-4)

Private Type HOOKINFO
  hWnd As Long        'Subclassed window
  Ctrl As DMenu
  OldWndProc As Long  'Old window procedure
End Type

Private m_HookArray() As HOOKINFO
Private m_MaxHooks As Long
Private Const INCREMENT As Long = 128

Public Sub Main()
  gMenuIDNext = MENU_INITIAL_ID
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

Public Sub HookWindow(ByVal hWnd As Long, ByVal Ctrl As DMenu)
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
  Dim vbmi As VBMenuItem
  Dim vbm As VBMenu
  Dim c As Collection
  Dim DMenu As DMenu
  Dim ID As Long
  
  For i = 1 To m_MaxHooks
    If m_HookArray(i).hWnd = hWnd Then
      If Msg = WM_COMMAND Then
        If lParam = 0 Then
          ID = LoWord(wParam)
          Set DMenu = m_HookArray(i).Ctrl
          For Each vbm In DMenu
            Set vbmi = vbm.GetVBMenu(ID)
            If Not vbmi Is Nothing Then
              Call DMenu.RaiseMenuClick(vbm, vbmi)
              Exit For
            End If
          Next
          Exit For
        End If
      End If
      Exit For
    End If
  Next i
  WndProc = CallWindowProc(m_HookArray(i).OldWndProc, hWnd, Msg, wParam, lParam)
End Function

Public Function HiWord(ByVal l As Long) As Long
  HiWord = l \ &H10000 And &HFFFF&
End Function

Public Function LoWord(ByVal dw As Long) As Long
  LoWord = (dw And &HFFFF&)
End Function

