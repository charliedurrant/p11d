VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Submenus As Collection
Public Tag As Variant

Private m_Caption As String, m_ParentName As String, m_Name As String
Private m_MenuID As Long, m_hMenuParent As Long, m_hMenuPopUp As Long
Private m_MenuUsed As Boolean, m_Checked As Boolean, m_Enabled As Boolean, m_Visible As Boolean
Private m_DrawnAlready As Boolean
Private m_hBitmap As Long
Private m_VBMenu As VBMenu
Private m_MenuIndex As Long
Private m_CheckPicture As StdPicture

Public Property Let MenuID(ByVal NewValue As Long)
  If m_MenuID <> 0 Then Err.Raise 380, "MenuID", "Cannot set MenuID more than once"
  m_MenuID = NewValue
End Property

Public Property Get MenuID() As Long
  MenuID = m_MenuID
End Property

Public Property Get ParentMenu() As Boolean
  ParentMenu = (Submenus.Count > 0)
End Property

Public Property Get MenuUsed() As Boolean
  MenuUsed = m_MenuUsed
End Property

Friend Property Let MenuUsed(ByVal NewValue As Boolean)
  m_MenuUsed = NewValue
End Property

Public Property Get MenuIndex() As Long
  MenuIndex = m_MenuIndex
End Property

Friend Property Let MenuIndex(ByVal NewValue As Long)
  m_MenuIndex = NewValue
End Property

Public Property Set CheckPicture(ByVal NewValue As StdPicture)
  Set m_CheckPicture = NewValue
  If m_hMenuParent > 0 Then
    If Me.ParentMenu Then Call Err.Raise(380, "Checked", "Cannot set picture as is parent menu.")
    Call SetCheck
  End If
End Property

Public Property Get CheckPicture() As StdPicture
  Set CheckPicture = m_CheckPicture
End Property

Friend Property Get hBitmap() As Long
  hBitmap = m_hBitmap
End Property

Friend Property Let Name(ByVal NewValue As String)
  m_Name = NewValue
End Property

Public Property Get Name() As String
  Name = m_Name
End Property

Friend Property Let ParentName(ByVal NewValue As String)
  m_ParentName = NewValue
End Property

Public Property Get ParentName() As String
  ParentName = m_ParentName
End Property

Friend Property Let DrawnAlready(ByVal NewValue As Boolean)
  m_DrawnAlready = NewValue
End Property

Friend Property Let hMenuPopUp(ByVal NewValue As Long)
  m_hMenuPopUp = NewValue
End Property

Friend Property Get hMenuPopUp() As Long
  hMenuPopUp = m_hMenuPopUp
End Property

Friend Property Get hMenuParent() As Long
  hMenuParent = m_hMenuParent
End Property

Friend Property Let hMenuParent(ByVal NewValue As Long)
  m_hMenuParent = NewValue
End Property

Public Property Let Visible(ByVal NewValue As Boolean)
  m_Visible = NewValue
  
  If m_DrawnAlready Then
    Call m_VBMenu.Refresh
  End If
End Property

Public Property Get Visible() As Boolean
  Visible = m_Visible
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  If m_hMenuParent <> 0 Then
    If NewValue Then
      Call EnableMenuItem(m_hMenuParent, m_MenuID, MF_ENABLED)
    Else
      Call EnableMenuItem(m_hMenuParent, m_MenuID, MF_GRAYED)
    End If
    Call DrawMenuBar(m_VBMenu.hWnd)
  End If
  m_Enabled = NewValue
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Friend Property Get VBMenu() As VBMenu
  Set VBMenu = m_VBMenu
End Property

Friend Property Set VBMenu(ByVal NewValue As VBMenu)
  Set m_VBMenu = NewValue
End Property

Public Property Let Checked(ByVal NewValue As Boolean)
  Dim ret As Long
  
  If m_hMenuParent <> 0 Then
    If Me.ParentMenu Then Call Err.Raise(380, "Checked", "Cannot set checked as is parent menu.")
    If NewValue Then
      ret = CheckMenuItem(m_hMenuParent, m_MenuID, MF_CHECKED)
    Else
      ret = CheckMenuItem(m_hMenuParent, m_MenuID, MF_UNCHECKED)
    End If
    Call DrawMenuBar(m_VBMenu.hWnd)
  End If
  m_Checked = NewValue
End Property

Public Property Get Checked() As Boolean
  Checked = m_Checked
End Property

Public Property Let Caption(ByVal NewValue As String)
  m_Caption = NewValue
End Property

Public Property Get Caption() As String
  Caption = m_Caption
End Property

Friend Property Get UFlags() As Long
  Dim l As Long
  
  If m_Checked Then l = l Or MF_CHECKED
  If m_Enabled Then
    l = l Or MF_ENABLED
  Else
    l = l Or MF_GRAYED
  End If
  If IsSeparator Then l = l Or MF_SEPARATOR
  If Not m_Visible Then l = l Or MF_OWNERDRAW
  UFlags = l
End Property

Friend Property Get IsSeparator()
  IsSeparator = (StrComp(MENU_SEP, m_Caption, vbBinaryCompare) = 0)
End Property

Public Sub PopUp(ByVal x As Single, ByVal y As Single)
  Dim pT As POINTAPI
  
  If Not Me.ParentMenu Then Call Err.Raise(380, "PopUp", "Cannot popup non parent menu.")
    
  pT.x = x / Screen.TwipsPerPixelX
  pT.y = y / Screen.TwipsPerPixelY
  Call ClientToScreen(m_VBMenu.hWnd, pT)
  Call TrackPopupMenu(m_hMenuPopUp, 0, pT.x, pT.y, 0, m_VBMenu.hWnd, 0)
End Sub

Friend Sub SetCheck()
  Dim ID As Long
  
  If Not m_CheckPicture Is Nothing Then
    If hMenuPopUp <> 0 Then
      ID = m_hMenuPopUp
    Else
      ID = m_MenuID
    End If
    Call SetMenuItemBitmaps(m_hMenuParent, ID, MF_BYCOMMAND, 0, CreateBitmap)
  End If
End Sub

Friend Sub KillBitmap()
  If m_hBitmap <> 0 Then Call DeleteObject(m_hBitmap)
End Sub

Friend Function CreateBitmap() As Long
  Dim pTBlit As POINTAPI
  Dim sZCheckPicture As SIZE, sZCheck As SIZE
  Dim hDCMemSrc As Long, hDCMemDst, hDC As Long, hOldBitmapSrc As Long, hOldBitmapDst As Long
  Dim hBitmapNew As Long, hOldBrush As Long, hOldPen As Long, lCheckDims As Long
  
  Call KillBitmap
  
  If m_CheckPicture Is Nothing Then Exit Function
  
  lCheckDims = GetMenuCheckMarkDimensions
  sZCheck.cx = LoWord(lCheckDims) + 1
  sZCheck.cy = HiWord(lCheckDims) + 1
  
  sZCheckPicture = HimetricToPixel(CheckPicture.Width, CheckPicture.Height)
  
  hDC = GetDC(m_VBMenu.hWnd)
  hDCMemSrc = CreateCompatibleDC(hDC)
  hDCMemDst = CreateCompatibleDC(hDC)
  Call ReleaseDC(m_VBMenu.hWnd, hDC)
  
  hOldBitmapSrc = SelectObject(hDCMemSrc, CheckPicture.Handle)
  hOldBitmapDst = SelectObject(hDCMemDst, CreateCompatibleBitmap(hDCMemSrc, sZCheck.cx, sZCheck.cy))
  
  hOldBrush = SelectObject(hDCMemDst, CreateSolidBrush(&HFFFFFF))
  hOldPen = SelectObject(hDCMemDst, CreatePen(PS_NULL, 0, 0))
  Call Rectangle(hDCMemDst, 0, 0, sZCheck.cx, sZCheck.cy)
  
  If sZCheckPicture.cx < sZCheck.cx Then
    pTBlit.x = (sZCheck.cx - sZCheckPicture.cy) / 2
  End If
  
  If sZCheckPicture.cy < sZCheck.cy Then
    pTBlit.y = (sZCheck.cy - sZCheckPicture.cy) / 2
  End If
  
  Call BitBlt(hDCMemDst, pTBlit.x, pTBlit.y, sZCheck.cx, sZCheck.cy, hDCMemSrc, 0, 0, SRCCOPY)
  Call DeleteObject(SelectObject(hDCMemDst, hOldBrush))
  Call DeleteObject(SelectObject(hDCMemDst, hOldPen))
  
  Call SelectObject(hDCMemSrc, hOldBitmapSrc)
  Call DeleteDC(hDCMemSrc)
  hBitmapNew = SelectObject(hDCMemDst, hOldBitmapDst)
  Call DeleteDC(hDCMemDst)
  CreateBitmap = hBitmapNew
End Function

Friend Function GetCheckCheckPicture() As Long
  Dim lCheckDim As Long
End Function

Private Sub Class_Initialize()
  Set Submenus = New Collection

  m_Enabled = True
  m_Visible = True
  m_MenuUsed = False
  m_MenuIndex = -1
End Sub

Friend Sub AddSubmenu(ByVal SubMenu As VBMenuItem)
  SubMenu.MenuUsed = True
  Submenus.Add SubMenu
End Sub

Friend Sub Kill()
  Call KillBitmap
  Set m_VBMenu = Nothing
End Sub

Private Sub Class_Terminate()
  Call Kill
End Sub
