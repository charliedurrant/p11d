Attribute VB_Name = "Cursor"
Option Explicit

'system
Private CursorStack() As Long
Private StatusStack() As String
Private CurStack As Long
Private MAX_STACK As Long
Private Const Increment As Long = 50

Public Sub SetCursorEx(ByVal lCursor As MousePointerConstants, StatusString As String)
  If CurStack >= MAX_STACK Then
     MAX_STACK = MAX_STACK + Increment
     If IsArrayEx2(CursorStack) Then
       ReDim Preserve CursorStack(0 To MAX_STACK) As Long
     Else
       ReDim CursorStack(0 To MAX_STACK) As Long
     End If
     If IsArrayEx2(StatusStack) Then
       ReDim Preserve StatusStack(0 To MAX_STACK) As String
     Else
       ReDim StatusStack(0 To MAX_STACK) As String
     End If
  End If
  If vbg Is Nothing Then
    CursorStack(CurStack) = VB.Screen.MousePointer
  Else
    CursorStack(CurStack) = vbg.Screen.MousePointer
  End If
  StatusStack(CurStack) = StatusString
  CurStack = CurStack + 1
  If vbg Is Nothing Then
    VB.Screen.MousePointer = lCursor
  Else
    vbg.Screen.MousePointer = lCursor
  End If
  Debug.Print "Cursor " & CurStack
End Sub

'* Clears the current cursor and restores the prior cursor on the cursor stack
'*
'* return value:
'* none
Public Function ClearCursorEx(ByVal ShowError As Boolean, Optional ByVal PopNumber As Long = 1) As String
  Dim i As Long
  ClearCursorEx = ""
  If PopNumber < 0 Then ShowError = False
  i = 0
  Do
    i = i + 1
    If CurStack > 0 Then
      CurStack = CurStack - 1
      If vbg Is Nothing Then
        VB.Screen.MousePointer = CursorStack(CurStack)
      Else
        vbg.Screen.MousePointer = CursorStack(CurStack)
      End If
      If CurStack > 0 Then ClearCursorEx = StatusStack(CurStack - 1)
    Else
      If vbg Is Nothing Then
        VB.Screen.MousePointer = vbDefault
      Else
        vbg.Screen.MousePointer = vbDefault
      End If
      If ShowError Then Call Err.Raise(ERR_POPCURSOR, "ClearCursorEx", "ClearCursor called woth no matching SetCursor")
    End If
  Loop Until (i = PopNumber) Or (CurStack = 0)
  If CurStack = 0 Then
    If vbg Is Nothing Then
      VB.Screen.MousePointer = vbDefault
    Else
      vbg.Screen.MousePointer = vbDefault
    End If
  End If
  Debug.Print "Cursor " & CurStack
End Function

