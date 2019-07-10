VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Message As String
Public MessageImage As PANEL_IMAGE
Public MessageImageOn As Boolean
Public Panel As TCSPANEL
Private m_HourGlassImage As Long
Private Sub Class_Initialize()
  MessageImage = PI_NONE
End Sub

Public Sub Copy(ByVal CopyTo As PanelStatus)
  CopyTo.Message = Me.Message
  CopyTo.MessageImage = Me.MessageImage
  CopyTo.MessageImageOn = Me.MessageImageOn
  Set CopyTo.Panel = Panel
End Sub
Public Function EnableTimerFlash()
  EnableTimerFlash = ((MessageImage And PI_FLASH) = PI_FLASH)
End Function
Public Function EnableTimerHourGlasss()
  EnableTimerHourGlasss = ((MessageImage And PI_HOUR_GLASS) = PI_HOUR_GLASS)
End Function

Public Sub SetPanelImage(iml As ImageList)

  Dim p As StdPicture
  Dim PI As PANEL_IMAGE
  
  PI = MessageImage
  
  If PI And PI_FLASH Then PI = PI Xor PI_FLASH
  
  Select Case PI
    Case PI_NONE
      Set p = Nothing
    Case PI_BLANK
      Set p = iml.ListImages(1).Picture
    Case PI_LIGHTENING
      Set p = iml.ListImages(2).Picture
    Case PI_INFO
      Set p = iml.ListImages(3).Picture
    Case PI_HOUR_GLASS
      
      'setup timer
  End Select
  
  Set Panel.Picture = p
  
End Sub

Public Sub Flash(ByVal iml As ImageList)
  If (MessageImage And PI_FLASH) = PI_FLASH Then
    If MessageImageOn Then
      Call SetPanelImage(iml)
      MessageImageOn = False
    Else
      Set Panel.Picture = iml.ListImages(PI_BLANK).Picture
      MessageImageOn = True
    End If
    Exit Sub
  End If
End Sub
Public Sub HourGlass(ByVal iml As ImageList)
  If (MessageImage And PI_HOUR_GLASS) = PI_HOUR_GLASS Then
    If m_HourGlassImage = iml.ListImages.Count Then m_HourGlassImage = 0
    m_HourGlassImage = m_HourGlassImage + 1
    Set Panel.Picture = iml.ListImages(m_HourGlassImage).Picture
  End If
End Sub

