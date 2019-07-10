VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCSEventClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Name As String
Public Event DebugMenuItem(Name As String, Index As Long, Parent As TCSMenuItems)

'*NOT VISIBLE
Public Function DebugMenu(Item As Long, menu As TCSMenuItems, Name As String) As Boolean
  Dim DoRaise As Boolean
  On Error GoTo DebugMenu_err
  
  DebugMenu = True
  Select Case menu
    Case MNU_BREAK
      If MsgBox("Quit application without saving?", vbOKCancel, UCase$(mAppName)) = vbOK Then
        Unload frmAbout
        Set frmAbout = Nothing
        Unload frmErr
        Set frmErr = Nothing
        Unload frmDebugMenu
        Set frmDebugMenu = Nothing
        m_FatalError = True
        RaiseEvent DebugMenuItem(Name, Item, menu)
      End If
    Case MNU_SYSTEM
      Select Case Item
        Case 0
          Call ShowSystemInfo
        Case 1
          Call frmVerComp.Initialise
        Case 2
          Call ShowAppInfo
      Case Else
        RaiseEvent DebugMenuItem(Name, Item, menu)
      End Select
    Case MNU_DATABASE
      DoRaise = True
      If Not m_IDebugMenu Is Nothing Then
        DoRaise = Not m_IDebugMenu.DatabaseMenuItem(Name)
      End If
      If DoRaise Then RaiseEvent DebugMenuItem(Name, Item, menu)
    Case MNU_APPLICATION
      RaiseEvent DebugMenuItem(Name, Item, menu)
    Case Else
      DebugMenu = False
      Call ECASE_SYS("Unknown top level menu item")
  End Select
    
DebugMenu_exit:
  Exit Function
  
DebugMenu_err:
  DebugMenu = False
  Call ErrorMessageEx(ERR_ERROR, Err, "DebugMenu", "ERR_MENUITEM", "Error obtaining the menu debug item", False)
  Resume DebugMenu_exit
End Function

Public Sub FillDebugMenu(frm As Object)
  Dim sMenuItem As String
  Dim m As MenuClass
  Dim l As Long
  Static InFillDebugMenu As Boolean
  On Error GoTo FillDebugMenu_err
  If InFillDebugMenu Then Exit Sub
  InFillDebugMenu = True
  frm.mnuSystemItem(0).Tag = "SYSTEM"
  frm.mnuSystemItem(1).Tag = "COMPONENTS"
  frm.mnuSystemItem(2).Tag = "APPLICATION"
  
  frm.mnuDatabaseItem(0).Tag = "REPAIR_COMPACT"
  frm.mnuDatabaseItem(1).Tag = "SQL"
  
  'frm.mnuOtherItem(0).Tag =
  'frm.mnuUserDebugItem.Enabled = True
  If Not m_IDebugMenu Is Nothing Then
    frm.mnuDatabaseItem(0).Enabled = m_IDebugMenu.RepairCompactAvailable
    frm.mnuDatabaseItem(1).Enabled = m_IDebugMenu.SQLAvailable
  Else
    frm.mnuDatabaseItem(0).Enabled = False
    frm.mnuDatabaseItem(1).Enabled = False
  End If
  
  If gMenusCollection.Count > 0 Then
    For Each m In gMenusCollection
      l = -1
      sMenuItem = m.Caption
      Select Case m.Parent
        Case MNU_DATABASE
          If m.MenuIndex <> -1 Then
            l = m.MenuIndex
          Else
            l = frm.mnuDatabaseItem.Count
            Load frm.mnuDatabaseItem(l)
          End If
          frm.mnuDatabaseItem(l).Caption = m.Caption
          frm.mnuDatabaseItem(l).Enabled = m.Enabled
          frm.mnuDatabaseItem(l).Tag = m.Name
        Case MNU_SYSTEM
          If m.MenuIndex <> -1 Then
            l = m.MenuIndex
          Else
            l = frm.mnuEnvirItem.Count
            Load frm.mnuEnvirItem(l)
          End If
          frm.mnuEnvirItem(l).Caption = m.Caption
          frm.mnuEnvirItem(l).Enabled = m.Enabled
          frm.mnuEnvirItem(l).Tag = m.Name
        Case MNU_APPLICATION
          If m.MenuIndex <> -1 Then
            l = m.MenuIndex
          Else
            l = frm.mnuOtherItem.Count
            Load frm.mnuOtherItem(l)
          End If
          frm.mnuOtherItem(l).Caption = m.Caption
          frm.mnuOtherItem(l).Enabled = m.Enabled
          frm.mnuOtherItem(l).Tag = m.Name
        Case Else
          Call ECASE_SYS("CREATE_MENUS - INVALID PARENT")
      End Select
      m.MenuIndex = l
    Next m
    sMenuItem = ""
    'special case - no default items in menu
    If frm.mnuOtherItem.Count > 1 Then frm.mnuOtherItem(0).Visible = False
  End If
  
FillDebugMenu_end:
  InFillDebugMenu = False
  Exit Sub
  
FillDebugMenu_err:
  Call ErrorMessageEx(ERR_ERROR, Err, "FillDebugMenu", "Fill Debug Menu", "Error filling debug menu." & vbCrLf & "Current menu item: " & sMenuItem, False)
  Resume FillDebugMenu_end
End Sub

Public Sub AttemptEndApplication()
  RaiseEvent DebugMenuItem("BREAK", 0, MNU_BREAK)
End Sub

