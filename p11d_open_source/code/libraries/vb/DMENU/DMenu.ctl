VERSION 5.00
Begin VB.UserControl DMenu 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   930
   ScaleWidth      =   900
   ToolboxBitmap   =   "DMenu.ctx":0000
   Begin VB.Image ImgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "DMenu.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "DMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TCSDMENU_ERRORS
  ERR_VBMENU = TCSDMENU_ERROR + 1
  ERR_ACTION_INV
  ERR_SEPARATOR
End Enum

Private m_vbMenus As Collection
Private m_hWnd As Long

Event MenuClick(ByVal vbm As VBMenu, ByVal vbmi As VBMenuItem)

Private Sub UserControl_Initialize()
  Set m_vbMenus = New Collection
End Sub

Private Sub UserControl_Terminate()
  If m_hWnd <> 0 Then
    Call UnhookWindow(m_hWnd)
  End If
  Call KillMenus
  Set m_vbMenus = Nothing
End Sub

Private Sub UserControl_Resize()
  Call SIZE(ImgIcon.Width, ImgIcon.Height)
End Sub

Public Property Let hWnd(ByVal hWndNew As Long)
  Dim vbm As VBMenu
  
#If DEBUGVER Then
  Call MsgBox("Cannot Hook Window in Debug version of Menu")
  Exit Property
#End If
  If hWndNew <> m_hWnd Then
    If m_hWnd <> 0 Then Call UnhookWindow(m_hWnd)
    m_hWnd = hWndNew
    If m_hWnd <> 0 Then Call HookWindow(m_hWnd, Me)
    PropertyChanged
    Call Refresh
  End If
End Property

Public Property Get hWnd() As Long
  hWnd = m_hWnd
End Property

Public Property Get Item(ByVal Key As Variant) As VBMenu
Attribute Item.VB_UserMemId = 0
  Set Item = m_vbMenus.Item(Key)
End Property

Public Property Get Count() As Long
  Count = m_vbMenus.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = m_vbMenus.[_NewEnum]
End Property

Public Sub Refresh()
  Dim vbm As VBMenu
  
  On Error GoTo Refresh_ERR
  If m_hWnd <> 0 Then
    If GetMenu(m_hWnd) = 0 Then Call Err.Raise(380, "Refresh", "Requires one standard vb menu.")
  End If
  For Each vbm In m_vbMenus
    vbm.hWnd = m_hWnd
    vbm.Refresh
  Next
  
Refresh_END:
  Exit Sub
Refresh_ERR:
  Call Err.Raise(Err.Number, "Refresh", Err.Description)
End Sub

Friend Sub RaiseMenuClick(ByVal vbm As VBMenu, ByVal vbmi As VBMenuItem)
  RaiseEvent MenuClick(vbm, vbmi)
End Sub

Private Function KillMenus()
  Dim vbm As VBMenu
  For Each vbm In m_vbMenus
    vbm.KillMenus
  Next
End Function

Public Function Add(Optional ByVal Key As Variant) As VBMenu
  Dim vbm As VBMenu
  
  Set vbm = New VBMenu
  vbm.Key = Key
  vbm.hWnd = m_hWnd
  If Len(Key) > 0 Then
    m_vbMenus.Add vbm, Key
  Else
    m_vbMenus.Add vbm
  End If
  Set Add = vbm
End Function

Public Sub Remove(ByVal Key As Variant)
  Dim vbm As VBMenu
  
  Set vbm = m_vbMenus.Item(Key)
  Call m_vbMenus.Remove(Key)
  Call vbm.KillMenus
End Sub



