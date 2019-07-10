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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private mMenus As Collection
Private m_Key As Variant

Private Sub Class_Initialize()
  Set mMenus = New Collection
  Debug.Print "Menu init"
End Sub

Private Sub Class_Terminate()
  Call KillMenus
  Debug.Print "Menu terminate"
End Sub

Friend Property Let Key(ByVal NewValue As Variant)
  m_Key = NewValue
End Property

Friend Function GetVBMenu(ByVal ID As Long) As VBMenuItem
  Dim vbmi As VBMenuItem
  
  For Each vbmi In mMenus
    If vbmi.MenuID = ID Then
      Set GetVBMenu = vbmi
      Exit For
    End If
  Next
End Function

Friend Property Let hWnd(ByVal NewValue As Long)
  m_hWnd = NewValue
End Property

Friend Property Get hWnd() As Long
  hWnd = m_hWnd
End Property

Public Sub PopUp(ByVal Key As Variant, ByVal x As Single, y As Single)
  Dim vbmi As VBMenuItem
  
  On Error GoTo PopUp_ERR
  Set vbmi = mMenus.Item(Key)
  Call vbmi.PopUp(x, y)
  
  Exit Sub
PopUp_ERR:
  Call Err.Raise(Err.Number, "PopUp", Err.Description)
End Sub

Public Sub Refresh()
  Dim ret As Long
  
  If m_hWnd = 0 Then Err.Raise ERR_VBMENU, "RefreshMenus", "Cannot Refresh menus, MDI window handle is invalid"
  Call KillMenus
  Call RefreshMenuStructure
  Call CreateSubMenus(0, mMenus)
  
  ret = DrawMenuBar(m_hWnd)
  If ret = 0 Then Err.Raise ERR_VBMENU, "RefreshMenus", "DrawMenuBar failed (" & Err.LastDllError & ")"
End Sub

Friend Sub KillMenus()
  Dim pvbmi As VBMenuItem
  Dim ret As Long
  Dim hMasterMenu As Long
  
  If m_hWnd = 0 Then Exit Sub
  Debug.Print "Menu Kill"
  hMasterMenu = GetMenu(m_hWnd)
  
  For Each pvbmi In mMenus
    Call pvbmi.KillBitmap
    If pvbmi.hMenuPopUp = 0 Then
      Call DeleteMenu(pvbmi.hMenuParent, pvbmi.MenuID, MF_BYCOMMAND)
    End If
  Next
  
  For Each pvbmi In mMenus
    If pvbmi.hMenuPopUp <> 0 Then
      Call DeleteMenu(pvbmi.hMenuParent, pvbmi.hMenuPopUp, MF_BYCOMMAND)
    End If
    pvbmi.hMenuParent = 0
    pvbmi.hMenuPopUp = 0
  Next
End Sub

Private Sub RefreshMenuStructure()
  Dim vbmi As VBMenuItem, pvbmi As VBMenuItem

  On Error GoTo RefreshMenuStructure_err
  For Each pvbmi In mMenus
    For Each vbmi In mMenus
      If Not vbmi.MenuUsed Then
        If StrComp(pvbmi.Name, vbmi.ParentName, vbTextCompare) = 0 Then Call pvbmi.AddSubmenu(vbmi)
      End If
    Next vbmi
  Next pvbmi
  
  ' validation
  For Each vbmi In mMenus
    If (Len(vbmi.ParentName) > 0) And Not vbmi.MenuUsed Then Err.Raise ERR_ACTION_INV, "RefreshMenuStructure", "Menu" & vbmi.Name & " has a parent " & vbmi.ParentName & " that does not exist"
    If (vbmi.Submenus.Count = 0) And Not vbmi.MenuUsed Then Err.Raise ERR_ACTION_INV, "RefreshMenuStructure", "Menu" & vbmi.Name & " is a parent menu, but does not have any children"
    If (vbmi.Checked And vbmi.ParentMenu) Then Call Err.Raise(ERR_ACTION_INV, "RefreshMenuStructure", "Menu " & vbmi.Name & " is a parent and is set to be checked.")
    If (vbmi.IsSeparator) And (vbmi.ParentMenu Or vbmi.Checked Or (Not vbmi.CheckPicture Is Nothing)) Then Call Err.Raise(ERR_ACTION_INV, "RefreshMenuStructure", "Menu " & vbmi.Name & " is a separator and cannot be (checked or parent or has CheckPicture)")
  Next vbmi
  
RefreshMenuStructure_end:
  Exit Sub
RefreshMenuStructure_err:
  Call Err.Raise(Err.Number, "RefreshMenuStructure", Err.Description)
End Sub

Private Sub CreateSubMenus(ByVal hMenu As Long, ByVal Submenus As Collection)
  Dim vbmi As VBMenuItem, TopLevelOnly As Boolean
  Dim hPopup As Long, ret As Long
  Dim UFlags As Long
  
  On Error GoTo CreateSubMenus_ERR
  If hMenu = 0 Then
    TopLevelOnly = True
    hMenu = GetMenu(m_hWnd)
    If hMenu = 0 Then Call Err.Raise(380, "CreateSubMenus", "Form needs a standard vb menu, add invisible item.")
  End If
  For Each vbmi In Submenus
    vbmi.DrawnAlready = True
    If Not vbmi.Visible Then GoTo next_menu
    vbmi.hMenuParent = hMenu
    If TopLevelOnly Or (vbmi.Submenus.Count > 0) Then
      If TopLevelOnly And Len(vbmi.ParentName) > 0 Then GoTo next_menu
      
      hPopup = CreateMenu()
      Debug.Print "Create Menu: " & vbmi.Caption & " " & vbmi.MenuIndex
      ret = InsertMenu(hMenu, vbmi.MenuIndex, MF_BYPOSITION + MF_POPUP, hPopup, vbmi.Caption)
      vbmi.hMenuPopUp = hPopup
      If ret = 0 Then Err.Raise ERR_VBMENU, "CreateTopLevelMenus", "Cannot append menu " & vbmi.Name
      Call CreateSubMenus(hPopup, vbmi.Submenus)
    Else
      UFlags = vbmi.Checked
      vbmi.hMenuParent = hMenu
      ret = AppendMenu(vbmi.hMenuParent, vbmi.UFlags, vbmi.MenuID, vbmi.Caption)
    End If
    Call vbmi.SetCheck
next_menu:
  Next vbmi
  
CreateSubMenus_END:
  Exit Sub
CreateSubMenus_ERR:
  Call Err.Raise(Err.Number, "CreateSubMenus", Err.Description)
End Sub

Public Function Add(ByVal Name As String, ByVal Caption As String, ByVal ParentName As String, Optional ByVal MenuIndex As Long = -1) As VBMenuItem
  Dim vbmi As VBMenuItem
  
  Set vbmi = New VBMenuItem
  gMenuIDNext = gMenuIDNext + 1
  
  vbmi.Name = Name
  vbmi.Caption = Caption
  vbmi.ParentName = ParentName
  If (Len(ParentName) > 0) And (MenuIndex <> -1) Then Err.Raise ERR_VBMENU, "Add", "MenuIndex can only be specified for a root menu"
  vbmi.MenuIndex = MenuIndex
  
  vbmi.MenuID = gMenuIDNext
  Set vbmi.VBMenu = Me
  mMenus.Add vbmi, vbmi.Name
  Debug.Print "Add Menu: [" & vbmi.Name & "] Caption='" & Caption & "' Parent=[" & ParentName & "] Index=" & vbmi.MenuIndex
  Set Add = vbmi
End Function

Public Property Get Item(ByVal Index As Variant) As VBMenuItem
  Set Item = mMenus.Item(Index)
End Property

Public Property Get Count() As Long
  Count = mMenus.Count
End Property

Public Sub Remove(ByVal Index As Variant)
  Dim vbmi As VBMenuItem
  
  Set vbmi = mMenus.Item(Index)
  Call mMenus.Remove(Index)
  Call vbmi.Kill
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mMenus.[_NewEnum]
End Property

