VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.UserControl TCSStatus 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00808080&
   ScaleHeight     =   660
   ScaleWidth      =   7260
   ToolboxBitmap   =   "TCSStat.ctx":0000
   Begin TCSPROG.TCSProgressBar prg 
      Height          =   360
      Left            =   930
      TabIndex        =   0
      Top             =   150
      Width           =   2055
      _cx             =   4197929
      _cy             =   4194939
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   50
      BarBackColor    =   12632256
      BarForeColor    =   8388608
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   0
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   2
   End
End
Attribute VB_Name = "TCSStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Panels As Collection
Private m_NoRedraw As Boolean
Private m_LastPanelMouseMove As TCSPANEL
'property constants
Private Const S_FONT As String = "Font"

'Events
Public Event PanelMouseDown(ByVal p As TCSPANEL, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event PictureMouseDown(ByVal p As TCSPANEL, Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Get hdc() As Long
  hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get prg() As Object
  Set prg = UserControl.prg
End Property

Public Property Get Font() As StdFont
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewValue As StdFont)
  Dim p As TCSPANEL
  
  Set UserControl.Font = FontCopy(NewValue)
  Set UserControl.prg.Font = FontCopy(NewValue)
  For Each p In m_Panels
    Set p.Font = FontCopy(NewValue)
  Next
  PropertyChanged ("Font")
End Property

Public Function StepCaption(Caption As String)
  StepCaption = UserControl.prg.StepCaption(Caption)
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim p As TCSPANEL
  Dim SHT As STAT_HIT_TEST
  
  SHT = StatHitTest(p, x, y)
  Select Case SHT
    Case SHT_BITMAP
      RaiseEvent PictureMouseDown(p, Button, Shift, x, y)
      RaiseEvent PanelMouseDown(p, Button, Shift, x, y)
    Case SHT_PANEL
      RaiseEvent PanelMouseDown(p, Button, Shift, x, y)
  End Select
  
End Sub

Private Function StatHitTest(ByRef p As TCSPANEL, ByVal x As Single, ByVal y As Single) As STAT_HIT_TEST
  Dim pT As POINTAPI
  
  pT.x = x / Screen.TwipsPerPixelX
  pT.y = y / Screen.TwipsPerPixelY
  StatHitTest = HitTest(p, pT)
End Function

Private Function HitTest(ByRef pHit As TCSPANEL, pT As POINTAPI) As STAT_HIT_TEST
  Dim p As TCSPANEL
  Dim SHT As STAT_HIT_TEST
    
  SHT = SHT_NO_HIT
  For Each p In m_Panels
    SHT = p.HitTest(pT)
    If SHT <> SHT_NO_HIT Then
      Set pHit = p
      HitTest = SHT
      Exit For
    End If
  Next p
End Function

Private Sub UserControl_Initialize()
  m_NoRedraw = False
  
  Set m_Panels = New Collection
  Set UserControl.prg.Font = UserControl.Font
  Call StopPrg
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim SHT As STAT_HIT_TEST
  Dim p As TCSPANEL
  
  SHT = StatHitTest(p, x, y)
  If SHT <> SHT_NO_HIT Then
    If Not (p Is m_LastPanelMouseMove) Then
      Set m_LastPanelMouseMove = p
      UserControl.Extender.ToolTipText = p.ToolTipText
    End If
  Else
    Set m_LastPanelMouseMove = Nothing
    UserControl.Extender.ToolTipText = ""
  End If
End Sub

Private Sub UserControl_Terminate()
  Set m_Panels = Nothing
End Sub

Public Sub ClearCaptions()
  Dim p As TCSPANEL
  
  On Error Resume Next
  For Each p In m_Panels
    p.Caption = ""
  Next
  Set p = Nothing
  Call StopPrg
End Sub

Private Sub UserControl_InitProperties()
  On Error Resume Next
  UserControl.Extender.Align = vbAlignBottom
End Sub

Private Sub PaintMe()
  Dim p As TCSPANEL
  Dim hMemDC As Long, hBmpOld As Long, hBrushOld As Long, hPenOld As Long
  Dim sZ As SIZE
  Dim r As RECT
  
  hMemDC = CreateCompatibleDC(Me.hdc)
  If hMemDC = 0 Then Exit Sub
  'open draw
  hBmpOld = SelectObject(hMemDC, CreateCompatibleBitmap(Me.hdc, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY))
  hBrushOld = SelectObject(hMemDC, CreateSolidBrush(UserControl.BackColor))
  hPenOld = SelectObject(hMemDC, CreatePen(PS_SOLID, 1, UserControl.BackColor))
  
  Call GetClientRect(UserControl.hWnd, r)
  sZ = GetRectDimensions(r)
  Call Rectangle(hMemDC, r.Left, r.Top, r.Right, r.Bottom)
  
  Call DeleteObject(SelectObject(hMemDC, hPenOld))
  For Each p In m_Panels
    Call p.Draw(hMemDC)
  Next p
  
  Call BitBlt(UserControl.hdc, 0, 0, sZ.cx, sZ.cy, hMemDC, 0, 0, SRCCOPY)
  
  'close draw
  Call DeleteObject(SelectObject(hMemDC, hBrushOld))
  Call DeleteObject(SelectObject(hMemDC, hBmpOld))
  Call DeleteDC(hMemDC)

End Sub
Private Sub UserControl_Paint()
  Call PaintMe
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set UserControl.Font = PropBag.ReadProperty(S_FONT, Ambient.Font)
  Set prg.Font = UserControl.Font
End Sub

Private Sub UserControl_Resize()
  Call RecalcPanelsRects
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty(S_FONT, UserControl.Font, Ambient.Font)
  End With
End Sub
Public Function Step() As Long
  Step = UserControl.prg.Step
End Function

Public Sub StartPrg(ByVal Max As Long, Optional ByVal Caption As String = "", Optional ByVal Indicator As Indicator = Percentage)
  With UserControl.prg
    .Caption = Caption
    .Indicator = Indicator
    .Value = UserControl.prg.Min
    .Max = Max
  End With
End Sub

Public Sub GotoEndPrg()
  UserControl.prg.Value = UserControl.prg.Max
End Sub

Public Sub StopPrg()
On Error Resume Next

  With UserControl.prg
    .Indicator = None
    .Caption = ""
    .Value = .Min
    .Max = .Min + 1
  End With

End Sub


Public Sub SetPB(ByVal PBMax As Long)
  UserControl.prg.Max = PBMax
End Sub

Public Sub SetStatus(ByVal PBValue As Long, ParamArray Captions() As Variant)
  Dim i As Long, lb As Long
  Dim p As TCSPANEL
  Dim pCount As Long
  Static bInUse As Boolean
      
  With UserControl
    If PBValue = .prg.Min Or PBValue = .prg.Max Then
      Call StopPrg
      bInUse = False
    Else
      If Not bInUse Then
        .prg.Indicator = Percentage
        bInUse = True
      End If
      .prg.Value = PBValue
    End If
       
    If IsEmpty(Captions) Then Exit Sub
    lb = LBound(Captions)
    pCount = UBound(Captions) - lb + 1
    If pCount > m_Panels.Count Then Call Err.Raise(ERR_SETSTATUS, "SetStatus", "Expected captions for " & CStr(m_Panels.Count) & " panels")
    For i = 1 To m_Panels.Count
      Set p = m_Panels(i)
      If i <= pCount Then p.Caption = Captions(lb + i - 1)
    Next i
  End With
End Sub

Public Property Let PanelCount(ByVal NewValue As Long)
  Dim pWidth As Long, j As Long
  Dim p As TCSPANEL
  
  On Error Resume Next
  m_NoRedraw = True
  Set m_Panels = Nothing
  Set m_Panels = New Collection
  If NewValue < 1 Then Exit Property
  pWidth = 100 / (NewValue + 1)
  
  For j = 1 To NewValue
    Set p = Me.AddPanel(pWidth)
  Next
  m_NoRedraw = False
  Call RecalcPanelsRects
End Property

Public Sub SetPanelWidths(ParamArray Widths() As Variant)
  Dim i As Long, lb As Long
  Dim p As TCSPANEL
  Dim pCount As Long
  
  On Error GoTo SetPanelWidths_err
  If IsEmpty(Widths) Then Exit Sub
  m_NoRedraw = True
  lb = LBound(Widths)
  pCount = UBound(Widths) - lb + 1
  If pCount <> m_Panels.Count Then Call Err.Raise(ERR_SETWIDTHS, "SetPanelWidths", "Expected widths for " & CStr(m_Panels.Count) & " panels")
  For i = 1 To m_Panels.Count
    Set p = m_Panels(i)
    p.PercentageWidth = Widths(lb + i - 1)
  Next i
  m_NoRedraw = False
  Call RecalcPanelsRects
SetPanelWidths_end:
  Exit Sub
SetPanelWidths_err:
  m_NoRedraw = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'START COLLECTION ***************************************************
Public Function AddPanel(ByVal PercentageWidth As Long, Optional ByVal Caption As String = "", Optional ByVal Style As Appearance = Down3D, Optional ByVal sKey As String, Optional ByVal Picture As StdPicture) As TCSPANEL
  Dim Panel As TCSPANEL
  
  Set Panel = New TCSPANEL
  Panel.hWnd = Me.hWnd
  Panel.PercentageWidth = PercentageWidth
  Panel.Style = Style
  Set Panel.Font = FontCopy(UserControl.Font)
  Panel.BackColor = UserControl.BackColor
  Panel.ForeColor = UserControl.ForeColor
  Panel.Caption = Caption
  Set Panel.Picture = Picture
  If Len(sKey) = 0 Then
    m_Panels.Add Panel
  Else
    m_Panels.Add Panel, sKey
  End If
  Set AddPanel = Panel
  Call RecalcPanelsRects
  Set Panel = Nothing
End Function

Public Property Get Panels(ByVal vIndexKey As Variant) As TCSPANEL
  Set Panels = m_Panels(vIndexKey)
End Property

Public Property Get PanelCount() As Long
  PanelCount = m_Panels.Count
End Property

Public Sub RemovePanel(ByVal vIndexKey As Variant)
  Dim p As TCSPANEL
  
  Set m_LastPanelMouseMove = Nothing
  Set p = m_Panels(vIndexKey)
  m_Panels.Remove vIndexKey
  Call RecalcPanelsRects
End Sub

'END COLLECTION ***************************************************
Private Sub RecalcPanelsRects()
  Dim p As TCSPANEL
  Dim StatusBarRect As RECT
  Dim lCurrentLeft As Long
  Dim bNoPBar As Boolean
  Dim sZ As SIZE
  
  If m_NoRedraw Then Exit Sub
  Call GetClientRect(Me.hWnd, StatusBarRect)
  sZ = GetRectDimensions(StatusBarRect)
  Me.prg.Top = (StatusBarRect.Top + 2) * Screen.TwipsPerPixelX
  If StatusBarRect.Bottom < 2 Then StatusBarRect.Bottom = 2
  Me.prg.Height = (StatusBarRect.Bottom - 2) * Screen.TwipsPerPixelY
  
  For Each p In m_Panels
    With p
      .Top = StatusBarRect.Top + 2
      .Bottom = StatusBarRect.Bottom
      .Left = lCurrentLeft
      .Right = lCurrentLeft + (CDbl(sZ.cx) * (CDbl(p.PercentageWidth) / 100))
      lCurrentLeft = p.Right + L_PANELGAP
      If .Right >= sZ.cx Then
        .Right = sZ.cx
        bNoPBar = True
        Exit For
      End If
    End With
  Next
  
  If bNoPBar Then
    Me.prg.Visible = False
  Else
    Me.prg.Visible = True
    Me.prg.Left = lCurrentLeft * Screen.TwipsPerPixelX
    Me.prg.Width = (sZ.cx - lCurrentLeft) * Screen.TwipsPerPixelX
  End If
  Call PaintMe
End Sub
