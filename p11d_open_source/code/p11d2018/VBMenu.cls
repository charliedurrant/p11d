VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_STRING = &H0&
Private Const MF_POPUP = &H10&

Private Const L_MENU_ID_START As Long = 1024

Private mhWnd As Long
Private mMenus As Collection
Private m_MenuIDMax As Long
Private Const MENU_INITIAL_ID As Long = 1024
Private Const ERR_VBMENU  As Long = TCSCLIENT_ERROR + 1224
Private Const ERR_ACTION_INV As Long = TCSCLIENT_ERROR + 1224
Private Sub Class_Initialize()
  Set mMenus = New Collection
  m_MenuIDMax = L_MENU_ID_START
End Sub

Public Sub Initialise(ByVal hMDIWnd As Long)
  mhWnd = hMDIWnd
End Sub
Public Sub PopUp(ByVal x As Single, ByVal y As Single, ByVal MenuName As String)
  Dim vbmi As VBMenuItem
  Set vbmi = mMenus(MenuName)
  Call vbmi.PopUp(mhWnd, x, y)
End Sub
Public Sub RefreshMenus()
  Dim ret As Long
  If mhWnd = 0 Then Err.Raise ERR_VBMENU, "RefreshMenus", "Cannot Refresh menus, MDI window handle is invalid"
  'KillMenus( )
  Call RefreshMenuStructure
  Call CreateSubMenus(0, mMenus)
  ret = DrawMenuBar(mhWnd)
  If ret = 0 Then Err.Raise ERR_VBMENU, "RefreshMenus", "DrawMenuBar failed (" & Err.LastDllError & ")"
End Sub

Private Sub RefreshMenuStructure()
  Dim vbmi As VBMenuItem, pvbmi As VBMenuItem

  On Error GoTo RefreshMenuStructure_err
  For Each pvbmi In mMenus
    For Each vbmi In mMenus
      If StrComp(pvbmi.Name, vbmi.ParentName, vbTextCompare) = 0 Then
        Call pvbmi.AddSubmenu(vbmi)
      End If
    Next vbmi
  Next pvbmi
  ' validation
  For Each vbmi In mMenus
    If (Len(vbmi.ParentName) > 0) And Not vbmi.MenuUsed Then Err.Raise ERR_ACTION_INV, "RefreshMenuStructure", "Menu" & vbmi.Name & " has a parent " & vbmi.ParentName & " that does not exist"
    If (vbmi.Submenus.Count = 0) And Not vbmi.MenuUsed Then Err.Raise ERR_ACTION_INV, "RefreshMenuStructure", "Menu" & vbmi.Name & " is a parent menu, but does not have any children"
  Next vbmi
  
RefreshMenuStructure_end:
  Exit Sub
RefreshMenuStructure_err:
  Call ErrorMessage(ERR_ALLOWIGNORE + ERR_ERROR, Err, "RefreshMenuStructure", "Refresh Menus", "Error refreshing menu structure")
End Sub

Private Sub CreateSubMenus(ByVal hMenu As Long, ByVal Submenus As Collection)
  Dim vbmi As VBMenuItem, TopLevelOnly As Boolean
  Dim hPopup As Long, ret As Long
  
  If hMenu = 0 Then
    TopLevelOnly = True
    hMenu = GetMenu(mhWnd)
  End If
  For Each vbmi In Submenus
    If TopLevelOnly Or (vbmi.Submenus.Count > 0) Then
      If TopLevelOnly And Len(vbmi.ParentName) > 0 Then GoTo next_menu
      hPopup = CreateMenu()
      ret = AppendMenu(hMenu, MF_POPUP, hPopup, vbmi.Caption)
      vbmi.hMenu = hPopup
      If ret = 0 Then Err.Raise ERR_VBMENU, "CreateTopLevelMenus", "Cannot append menu " & vbmi.Name
      Call CreateSubMenus(hPopup, vbmi.Submenus)
    Else
      
      ret = AppendMenu(hMenu, MF_ENABLED, vbmi.MenuID, vbmi.Caption)
      vbmi.hMenu = ret
    End If
    
next_menu:
  Next vbmi
End Sub

Public Sub AddMenu(ByVal MenuName As String, ByVal MenuCaption As String, ByVal ParentName As String, Optional ByVal UserID As Long = -1)
  Dim vbmi As VBMenuItem
  
  Set vbmi = New VBMenuItem
  vbmi.Name = MenuName
  vbmi.Caption = MenuCaption
  vbmi.ParentName = ParentName
  vbmi.UserID = UserID
  m_MenuIDMax = m_MenuIDMax + 1
  
  vbmi.MenuID = m_MenuIDMax
  mMenus.Add vbmi, vbmi.Name
End Sub
Public Property Get MenuIDMin() As Long
  MenuIDMin = L_MENU_ID_START - 1
End Property
Public Property Get MenuIDMax() As Long
  MenuIDMax = m_MenuIDMax
End Property
Public Function GetVBMenuItem(ByVal MenuID As Long) As VBMenuItem
  
  Dim vbmi As VBMenuItem
  
  For Each vbmi In mMenus
    If MenuID = vbmi.MenuID Then
      Set GetVBMenuItem = vbmi
      Exit For
    End If
  Next
  
End Function

Private Sub AddMenus3(ByVal MenuName As String, ByVal MenuCaption As String)
  Dim hPopup  As Long, hMenu As Long, hPopup2  As Long
  Dim ID As Long, ret As Long
    

'  hPopup = CreateMenu()
'  ID = 1065
'  ret = AppendMenu(hPopup, MF_ENABLED, ID, "text 1")
'  ID = ID + 1
'  ret = AppendMenu(hPopup, MF_ENABLED, ID, "text 2")
'
'  hPopup2 = CreateMenu()
'  ID = ID + 1
'  ret = AppendMenu(hPopup2, MF_ENABLED, ID, "sub text 1")
'  ID = ID + 1
'  ret = AppendMenu(hPopup2, MF_ENABLED, ID, "sub text 2")
'   ret = AppendMenu(hPopup, MF_POPUP, hPopup2, "sub menu")
'
'  ' kkk
'  hMenu = GetMenu(hWnd)
'  ret = AppendMenu(hMenu, MF_POPUP, hPopup, "new pop-up")
'  ret = DrawMenuBar(hWnd)
End Sub



