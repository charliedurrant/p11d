VERSION 5.00
Begin VB.UserControl HOOK 
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   975
   ScaleWidth      =   1080
   ToolboxBitmap   =   "TCSHOOK.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "TCSHOOK.ctx":00FA
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "HOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MsgSpy Subclassing Control"
Option Explicit

Private m_hWnd As Long
Private m_Messages() As Long
Private m_NumMessages As Integer

Event WndProc(Discard As Boolean, MsgReturn As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Property Let Messages(ByVal Message As Long, ByVal SubClassTheMessage As Boolean)
  Dim i As Long, j As Long
  
  For i = 1 To m_NumMessages
    If m_Messages(i) = Message Then
      If SubClassTheMessage Then
        Exit Property
      Else
        m_NumMessages = m_NumMessages - 1
        For j = i To m_NumMessages
            m_Messages(j) = m_Messages(j + 1)
        Next j
        ReDim Preserve m_Messages(m_NumMessages)
        Exit Property
      End If
    End If
  Next i
  If SubClassTheMessage Then
    m_NumMessages = m_NumMessages + 1
    ReDim Preserve m_Messages(m_NumMessages)
    m_Messages(m_NumMessages) = Message
  End If
End Property

Public Property Get Messages(ByVal Message As Long) As Boolean
Attribute Messages.VB_Description = "Specifies which messages are passed to the WndProc event"
Attribute Messages.VB_MemberFlags = "400"
  Dim i As Integer
  
  For i = 1 To m_NumMessages
    If m_Messages(i) = Message Then
      Messages = True
      Exit Property
    End If
  Next i
  Messages = False
End Property

Public Property Let hWnd(ByVal hWndNew As Long)
#If DEBUGVER Then
  Call MsgBox("Cannot Hook Window in Debug version of Hook")
  Exit Property
#End If
  If hWndNew <> m_hWnd Then
    If m_hWnd <> 0 Then Call UnhookWindow(m_hWnd)
    m_hWnd = hWndNew
    If m_hWnd <> 0 Then Call HookWindow(m_hWnd, Me)
    PropertyChanged
  End If
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Specifies the handle of the window to be subclassed"
Attribute hWnd.VB_MemberFlags = "400"
  hWnd = m_hWnd
End Property

Friend Function RaiseWndProc(MsgResult As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
  Dim Discard As Boolean
  
  RaiseEvent WndProc(Discard, MsgResult, m_hWnd, Msg, wParam, lParam)
  RaiseWndProc = Discard
End Function

Private Sub UserControl_Resize()
  Call Size(imgIcon.Width, imgIcon.Height)
End Sub

Private Sub UserControl_Terminate()
  If m_hWnd <> 0 Then Call UnhookWindow(m_hWnd)
End Sub

