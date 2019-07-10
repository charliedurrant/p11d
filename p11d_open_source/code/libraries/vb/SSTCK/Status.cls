VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const L_DEFAULT_PANEL_SIZE As Long = 30

Public MouseCursor As MousePointerConstants
Public p1 As PanelStatus
Public p2 As PanelStatus
Public IsDefault As Boolean
Private Sub Class_Initialize()
  Set p1 = New PanelStatus
  Set p2 = New PanelStatus
End Sub

Public Function Copy() As Status
  Dim NewCopy As Status
  Set NewCopy = New Status
  Call Me.p1.Copy(NewCopy.p1)
  Call Me.p2.Copy(NewCopy.p2)
  Set Copy = NewCopy
End Function

Public Sub Flash(ByVal iml As ImageList)
  'if we ever convert these to store multiple PanelStatus then I will go crazy
  Call p1.Flash(iml)
  Call p2.Flash(iml)
End Sub
Public Sub HourGlass(ByVal iml As ImageList)
  'if we ever convert these to store multiple PanelStatus then I will go crazy
  Call p1.HourGlass(iml)
  Call p2.HourGlass(iml)
End Sub

Private Sub EnableTimers(tmrFlash As Timer, tmrHourGlass As Timer)
  tmrFlash.Enabled = p1.EnableTimerFlash Or p2.EnableTimerFlash
  tmrHourGlass.Enabled = p1.EnableTimerHourGlasss Or p2.EnableTimerHourGlasss
End Sub

Friend Sub ToStatusBar(ByVal stat As TCSStatus, ByVal Statuss As ObjectList, ByVal tmrFlash As Timer, ByVal tmrHourGlass As Timer, ByVal iml As ImageList, ByVal stdefault As Status)
  Dim p As StdPicture
  Dim i As Long
  Dim st As Status
  'pass the current status to the status bar if one available
  If stat Is Nothing Then Exit Sub
  

  
  If stat.PanelCount = 0 Then
    'add the 2 panels in
    
    Set p1.Panel = stat.AddPanel(L_DEFAULT_PANEL_SIZE, p1.Message, Down3D, "P1", p)
    Call p1.SetPanelImage(iml)
    
    Set p2.Panel = stat.AddPanel(L_DEFAULT_PANEL_SIZE, p1.Message, Down3D, "P2", p)
    Call p2.SetPanelImage(iml)
    'link up all status to have links to panels
    For i = 1 To Statuss.Count
      Set st = Statuss(i)
      If Not st Is Nothing Then
        If Not st Is Me Then
          Set st.p1.Panel = p1.Panel
          Set st.p2.Panel = p2.Panel
        End If
      End If
    Next
    Set stdefault.p1.Panel = p1.Panel
    Set stdefault.p2.Panel = p2.Panel
    
  Else
    
    p1.Panel.Caption = p1.Message
    Call p1.SetPanelImage(iml)
    p2.Panel.Caption = p2.Message
    Call p1.SetPanelImage(iml)
    
  End If
      
  Call EnableTimers(tmrFlash, tmrHourGlass)
  

End Sub

