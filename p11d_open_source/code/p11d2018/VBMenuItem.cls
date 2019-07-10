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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Caption As String
Public ParentName As String
Public MenuID As Long
Public mMenuUsed As Boolean
Public UserID As Long
Public ParentMenu As Boolean
Public Submenus As Collection
Public hMenu As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lNull As Long) As Long

Private Sub Class_Initialize()
  Set Submenus = New Collection
End Sub
Public Sub PopUp(ByVal hWnd As Long, ByVal x As Single, y As Single)
  Dim lx As Long, ly As Long
  Dim lRet As Long
  lx = x / Screen.TwipsPerPixelX
  ly = y / Screen.TwipsPerPixelY
  lRet = TrackPopupMenu(hMenu, 0, lx, ly, 0, hWnd, 0)
  
End Sub
Public Property Get MenuUsed() As Boolean
  MenuUsed = mMenuUsed
End Property

Public Property Let MenuUsed(ByVal NewValue As Boolean)
  mMenuUsed = NewValue
End Property

Public Sub AddSubmenu(ByVal NewSubmenuItem As VBMenuItem)
  On Error GoTo AddSubmenu_err
  
  NewSubmenuItem.MenuUsed = True
  Submenus.Add NewSubmenuItem
  Exit Sub

AddSubmenu_err:
  Err.Raise Err.Number, ErrorSource(Err, "AddSubAction"), "Error in AddSubAction " & Err.Description
End Sub
