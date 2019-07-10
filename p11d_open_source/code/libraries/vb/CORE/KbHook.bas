Attribute VB_Name = "KeyboardHook"
Option Explicit
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public gKeyHook As IKeyboardHook

Private Const WH_KEYBOARD As Long = 2
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_MENU As Long = &H12

Private m_hHook As Long


Public Sub InitKeyboardHook()
  If Not IsRunningInIDEEx() Then
    m_hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf TCSKeyBoardHook, 0&, ghThreadID)
  End If
End Sub

Public Sub KillKeyboardHook()
  If m_hHook <> 0 Then Call UnhookWindowsHookEx(m_hHook)
  m_hHook = 0
End Sub


Public Function TCSKeyBoardHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim Shift As Integer
  Dim iwParam As Integer
  
  If nCode >= 0 Then
    If wParam = vbKeyF12 Then
      Call ShowDebugPopupex
      TCSKeyBoardHook = 1
      Exit Function
    ElseIf Not gKeyHook Is Nothing Then
      If cIsKeyFirstKeyDown(lParam) Then
        Shift = GetControlKeysStateEx()
        iwParam = LOWORD(wParam)
        Debug.Print "Key: " & iwParam & " Shift: " & Shift
        Call gKeyHook.KeyDown(iwParam, Shift)
        If iwParam = 0 Then
          TCSKeyBoardHook = 1
          Exit Function
        End If
      End If
    End If
  End If
  TCSKeyBoardHook = CallNextHookEx(m_hHook, nCode, wParam, lParam)
End Function

Public Function GetControlKeysStateEx() As ShiftConstants
  GetControlKeysStateEx = 0
  If GetKeyState(VK_SHIFT) < 0 Then GetControlKeysStateEx = GetControlKeysStateEx Or vbShiftMask
  If GetKeyState(VK_CONTROL) < 0 Then GetControlKeysStateEx = GetControlKeysStateEx Or vbCtrlMask
  If GetKeyState(VK_MENU) < 0 Then GetControlKeysStateEx = GetControlKeysStateEx Or vbAltMask
End Function

