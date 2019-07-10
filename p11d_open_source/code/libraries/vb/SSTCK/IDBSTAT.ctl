VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl SSTCK 
   Alignable       =   -1  'True
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   6285
   ToolboxBitmap   =   "IDBSTAT.ctx":0000
   Begin MSComctlLib.ImageList imlHourGlass 
      Left            =   1230
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":0664
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":0D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":105A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":13AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":16FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":1A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":1DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":20F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":2446
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrHourGlass 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2520
      Top             =   1620
   End
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1890
      Top             =   1620
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   1200
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":2798
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":28AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IDBSTAT.ctx":29BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image ImgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "IDBSTAT.ctx":2ACE
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "SSTCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Enum PANEL_IMAGE
  PI_BLANK = 1
  PI_LIGHTENING
  PI_INFO
  PI_HOUR_GLASS
  PI_NONE = 128
  PI_FLASH = 256
End Enum

Public Enum STATUS_ID
  STATUSID_1 = 1
  STATUSID_2
End Enum

Private m_Panel1 As TCSPANEL
Private m_Panel2 As TCSPANEL

Private m_DefaultStatus As Status
Private m_Statuss As ObjectList
Private m_CurrentStatus As Status
Private m_Stat As atc2Stat.TCSStatus
Private m_SetDefaults As Boolean


Public Property Set stat(ByVal NewValue As Object)
  If Not m_Stat Is Nothing Then Call Err.Raise(380, "stat", "The stat property can only be set once")
  If Not m_SetDefaults Then Call Err.Raise(380, "stat", "The stat property can only be set if called DefaultStatus")
  
  Set m_Stat = NewValue
  
  If NewValue.PanelCount <> 0 Then
    Set m_Stat = Nothing
    Call Err.Raise(380, "stat", "The stat must have no panels")
  End If
  Call m_CurrentStatus.ToStatusBar(NewValue, m_Statuss, tmrFlash, tmrHourGlass, iml, m_DefaultStatus)
  
End Property

Private Sub tmrFlash_Timer()
  Call m_CurrentStatus.Flash(iml)
End Sub

Private Sub tmrHourGlass_Timer()
  Call m_CurrentStatus.HourGlass(imlHourGlass)
End Sub

Private Sub UserControl_Initialize()
  'stores all the status
  Set m_Statuss = New ObjectList
  Set m_DefaultStatus = New Status
  Set m_CurrentStatus = m_DefaultStatus
  m_DefaultStatus.IsDefault = True
End Sub

Public Sub DefaultStatus(ByVal SID As STATUS_ID, ByVal Caption As String, Optional ByVal PanelImage As PANEL_IMAGE = PI_NONE)

  m_SetDefaults = True

  Select Case SID
    Case STATUSID_1
      m_DefaultStatus.p1.Message = Caption
      m_DefaultStatus.p1.MessageImage = PanelImage
      
    Case STATUSID_2
      m_DefaultStatus.p2.Message = Caption
      m_DefaultStatus.p2.MessageImage = PanelImage
    Case Else
      Call Err.Raise(380, "DefaultStatus", "Invalid status id.")
  End Select
End Sub

' copy the prvious one and amend for new one
Public Sub PushStatus(ByVal SID As STATUS_ID, ByVal Caption As String, Optional ByVal PanelImage As PANEL_IMAGE = PI_NONE, Optional ByVal MouseCursor As MousePointerConstants = vbDefault)
  Dim st As Status, ps As PanelStatus
  
  Set st = m_CurrentStatus.Copy
  Select Case SID
    Case STATUSID_1
      Set ps = st.p1
    Case STATUSID_2
      Set ps = st.p2
    Case Else
      Call Err.Raise(380, "PushStatus", "Invalid Status ID")
  End Select
  
  st.MouseCursor = MouseCursor
  ps.Message = Caption
  ps.MessageImage = PanelImage
    
  Call m_Statuss.Add(st)
  Set m_CurrentStatus = st
  If m_CurrentStatus.MouseCursor <> vbDefault Then Call SetCursor(m_CurrentStatus.MouseCursor)
  Call m_CurrentStatus.ToStatusBar(m_Stat, m_Statuss, tmrFlash, tmrHourGlass, iml, m_DefaultStatus)
  
End Sub

Public Sub PopStatus()
  Dim i As Long, lastcursor As MousePointerConstants
    
  If m_CurrentStatus Is m_DefaultStatus Then Call Err.Raise(380, "PopStatus", "Pop without push.")
  lastcursor = m_CurrentStatus.MouseCursor
  i = m_Statuss.ItemIndex(m_CurrentStatus)
  Call m_Statuss.Remove(i)
  Call m_Statuss.CompactTop
  If i > 1 Then
    Set m_CurrentStatus = m_Statuss.Item(i - 1)
  Else
    Set m_CurrentStatus = m_DefaultStatus
  End If
  If lastcursor <> vbDefault Then Call ClearCursor
  Call m_CurrentStatus.ToStatusBar(m_Stat, m_Statuss, tmrFlash, tmrHourGlass, iml, m_DefaultStatus)
End Sub

Public Sub PopStatusAll()
  Call m_Statuss.RemoveAll
  Call ClearAllCursors
  Set m_CurrentStatus = m_DefaultStatus
  Call m_CurrentStatus.ToStatusBar(m_Stat, m_Statuss, tmrFlash, tmrHourGlass, iml, m_DefaultStatus)
End Sub


Public Sub Clear()
  Call m_Statuss.RemoveAll
  Set m_CurrentStatus = m_DefaultStatus
  Call ClearAllCursors
  Call m_CurrentStatus.ToStatusBar(m_Stat, m_Statuss, tmrFlash, tmrHourGlass, iml, m_DefaultStatus)
End Sub


Private Sub UserControl_Resize()
  Call Size(ImgIcon.Width, ImgIcon.Height)
End Sub
