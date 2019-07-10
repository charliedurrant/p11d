Attribute VB_Name = "MenuTest"
Option Explicit
Public mMenu As VBMenu

Public Sub AddMenus(ByVal hWnd As Long)
  Set mMenu = New VBMenu
  
  Call mMenu.AddMenu("TopLevel", "This is a Menu", "")
  Call mMenu.AddMenu("TopLevel2", "Root menu 2", "")
  Call mMenu.AddMenu("SubLevel1a", "This is a SubMenu L1", "TopLevel")
  Call mMenu.AddMenu("SubLevel1b", "This is a SubMenu L1", "TopLevel")
  Call mMenu.AddMenu("SubLevel1c", "This is a SubMenu L1", "TopLevel")
  Call mMenu.AddMenu("SubLevel2a", "This is a SubMenu L2", "SubLevel1a")
  Call mMenu.AddMenu("SubLevel3", "This is a SubMenu L1", "TopLevel2")
  Call mMenu.Initialise(frmMain.hWnd)
  Call mMenu.RefreshMenus
End Sub
