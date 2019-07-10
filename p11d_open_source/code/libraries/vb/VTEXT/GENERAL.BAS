Attribute VB_Name = "General"
Option Explicit

Private Const LB_FINDSTRING& = &H18F
Private Const LB_ERR& = (-1)
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)

Public Function lSelectListItem&(lst As ListBox, sUserText$)
  Dim lRet&
  
  With lst
    '-1 means search whole box
    lRet = SendMessageByString(.hWnd, LB_FINDSTRING, -1, sUserText)
    If lRet <> LB_ERR Then
      .ListIndex = lRet
    ElseIf .ListCount Then
      .ListIndex = 0
    End If
  End With
  
End Function


