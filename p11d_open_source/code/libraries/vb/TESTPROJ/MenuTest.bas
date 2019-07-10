Attribute VB_Name = "MenuTest"
Option Explicit
Public mMenu As VBMenu

Public Sub AddMenus(ByVal hwnd As Long)
  Set mMenu = New VBMenu
  
  Call mMenu.AddMenu("TopLevel", "This is a Menu", "")
  Call mMenu.AddMenu("TopLevel2", "Root menu 2", "")
  Call mMenu.AddMenu("SubLevel1a", "This is a SubMenu L1", "TopLevel")
  Call mMenu.AddMenu("SubLevel1b", "This is a SubMenu L1", "TopLevel")
  Call mMenu.AddMenu("SubLevel1c", "This is a SubMenu L1", "TopLevel")
  Call mMenu.AddMenu("SubLevel2a", "This is a SubMenu L2", "SubLevel1a")
  Call mMenu.AddMenu("SubLevel3", "This is a SubMenu L1", "TopLevel2")
  Call mMenu.Initialise(frmMain.hwnd)
  'Call mMenu.RefreshMenus
End Sub


Public Sub TestLDAP()
  Dim UserName As String, Password As String
  Dim lh As LDAPHelper, i As Long
  Dim s As String
  Dim maxattempts As Long, t0 As Long
  
  
  t0 = GetTicks
  maxattempts = 10
  UserName = "albert.fleming"
  Password = ""
  Set lh = New LDAPHelper
  lh.ServerContext = "<tcsau5001.arthurandersen.com;tcsau5002.arthurandersen.com;tcsau5003.arthurandersen.com;tcsau5004.arthurandersen.com;tcsau5005.arthurandersen.com;tcsau5006.arthurandersen.com;tcsau5007.arthurandersen.com;tcsau5008.arthurandersen.com;tcsau5099.arthurandersen.com>"
  For i = 1 To maxattempts
    Call lh.Authenticate(UserName, Password, ADS_STANDARD)
  Next i
  s = lh.ServerContext
  
  Debug.Print lh.ServerContext
  lh.ServerContext = s
  For i = 1 To maxattempts
    Call lh.Authenticate(UserName, Password, ADS_STANDARD)
  Next i
  'Debug.Print lh.ServerContext
  t0 = GetTicks - t0
  s = lh.DebugServerContext & "Total time " & Format$(t0 / 1000, "#,###.00") & " seconds" & vbCrLf
  MsgBox s
'
'  lh.ServerContext = "<tcsau5003.arthurandersen.com;tcsau5001.arthurandersen.com>"
'  Call lh.Authenticate(username, password, ADS_USE_ENCRYPTION)
'  Call lh.Authenticate(username, password, ADS_USE_SSL)
'  Debug.Print lh.ServerContext
'
'  lh.ServerContext = "<tcsau50XX.arthurandersen.com;tcsau5003.arthurandersen.com;tcsau5001.arthurandersen.com>"
'  Call lh.Authenticate(username, password, ADS_STANDARD)
'  Call lh.Authenticate(username, password, ADS_USE_ENCRYPTION)
'  Call lh.Authenticate(username, password, ADS_USE_SSL)
'
'
'  lh.ServerContext = "<tcsau5001.arthurandersen.com>"
'  Call lh.Authenticate(username, password & "D", ADS_STANDARD)
'  Call lh.Authenticate(username & "D", password & "D", ADS_STANDARD)
'  Call lh.Authenticate(username & "D", password, ADS_STANDARD)
End Sub
