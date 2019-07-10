VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCSPANEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Style As Appearance

Private m_PercentageWidth As Long
Private m_ToolTipText
Private m_Caption As String
Private m_Rect As RECT
Private m_PictureRect As RECT
Private m_hWnd As Long
Private m_Picture As StdPicture
Public m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Private m_BackColor As Long, m_ForeColor As Long
Public Property Let ToolTipText(ByVal NewValue As String)
  m_ToolTipText = NewValue
End Property
Public Property Get ToolTipText() As String
  ToolTipText = m_ToolTipText
End Property
Public Property Set Font(ByVal NewValue As StdFont)
  Set m_Font = NewValue
  Call DrawMe
End Property
Public Property Get Font() As StdFont
  Set Font = m_Font
End Property
Public Property Let BackColor(ByVal NewValue As Long)
  If m_BackColor <> NewValue Then
    m_BackColor = NewValue
    Call DrawMe
  End If
End Property
Public Property Get BackColor() As Long
  BackColor = m_BackColor
End Property
Public Property Let ForeColor(ByVal NewValue As Long)
  If m_ForeColor <> NewValue Then
    m_ForeColor = NewValue
    Call DrawMe
  End If
End Property
Public Property Get ForeColor() As Long
  m_ForeColor = m_ForeColor
End Property
Public Property Set Picture(ByVal NewValue As StdPicture)
  Set m_Picture = NewValue
  Call DrawMe
End Property
Public Property Get Picture() As StdPicture
  Set Picture = m_Picture
End Property
Friend Property Let hWnd(ByVal NewValue As Long)
  m_hWnd = NewValue
End Property

Private Sub DrawMe()
  Dim hdc As Long
  
  hdc = GetDC(m_hWnd)
  If (hdc <> 0) Then
    Call Draw(hdc)
    Call ReleaseDC(m_hWnd, hdc)
  End If
End Sub

Public Property Let Caption(ByVal NewValue As String)
  If StrComp(m_Caption, NewValue, vbTextCompare) <> 0 Then
    m_Caption = NewValue
    Call DrawMe
  End If
End Property
Public Property Get Caption() As String
  Caption = m_Caption
End Property
Friend Property Get Left() As Long
  Left = m_Rect.Left
End Property
Friend Property Let Left(NewValue As Long)
  m_Rect.Left = NewValue
End Property
Friend Property Get Right() As Long
  Right = m_Rect.Right
End Property
Friend Property Let Right(NewValue As Long)
  m_Rect.Right = NewValue
End Property
Friend Property Get Top() As Long
  Top = m_Rect.Top
End Property
Friend Property Let Top(NewValue As Long)
  m_Rect.Top = NewValue
End Property
Friend Property Get Bottom() As Long
  Bottom = m_Rect.Bottom
End Property
Friend Property Let Bottom(NewValue As Long)
  m_Rect.Bottom = NewValue
End Property

Property Let PercentageWidth(ByVal NewValue As Long)
  Dim lRet As Long
  Dim r As RECT
  Dim lParam As Long
  
  m_PercentageWidth = NewValue
  If m_hWnd = 0 Then Call Err.Raise(380, "PercentageWidth", "No Window handle set")
' apf cd this does not work
'  Call GetClientRect(m_hWnd, r)
'  Call MoveWindow(m_hWnd, 0, 10, r.Right - r.Left, r.Bottom - r.Top, 1)
End Property

Public Property Get PercentageWidth() As Long
  PercentageWidth = m_PercentageWidth
End Property
Private Function HimetricToPixel(ByVal sngHimetricX As Single, ByVal sngHimetricY As Single) As SIZE
  Const sngRatio As Single = 567 / 1000
  Dim sngTwips As Single
  
  sngTwips = sngHimetricX * (sngRatio)
  HimetricToPixel.cx = sngTwips / Screen.TwipsPerPixelX
  sngTwips = sngHimetricY * (sngRatio)
  HimetricToPixel.cy = sngTwips / Screen.TwipsPerPixelY
End Function

Private Sub DrawPicture(BoundingRect As RECT, ByVal hdc As Long)
  Dim sngScreenX As Single, sngScreenY As Single
  Dim sZPicture As SIZE, sZBounding As SIZE
  Dim hMemDC As Long
  Dim lDiff As Long
  
  Dim hOldBitmap As Long
  
  If m_Picture Is Nothing Then Exit Sub
  
  sZPicture = HimetricToPixel(m_Picture.Width, m_Picture.Height)
  
  hMemDC = CreateCompatibleDC(hdc)
  hOldBitmap = SelectObject(hMemDC, m_Picture.Handle)
  'centre picture
  sZBounding = GetRectDimensions(BoundingRect)
  m_PictureRect = BoundingRect
  m_PictureRect.Right = m_PictureRect.Left + sZPicture.cx
  If sZPicture.cy < sZBounding.cy Then
    If m_PictureRect.Right > BoundingRect.Right Then m_PictureRect.Right = BoundingRect.Right
    lDiff = sZBounding.cy - sZPicture.cy
    If lDiff > 0 Then
      lDiff = lDiff / 2
      m_PictureRect.Top = m_PictureRect.Top + lDiff
      m_PictureRect.Bottom = m_PictureRect.Bottom - lDiff
    End If
    sZBounding = GetRectDimensions(m_PictureRect)
  End If
  Call BitBlt(hdc, m_PictureRect.Left, m_PictureRect.Top, sZPicture.cx, sZBounding.cy, hMemDC, 0, 0, SRCCOPY)
  Call SelectObject(hMemDC, hOldBitmap)
  Call DeleteDC(hMemDC)
  BoundingRect.Left = m_PictureRect.Right
  'adjust picture rect for overall control for hittest
  'm_Rect is the rect for the panel
  m_PictureRect.Left = m_Rect.Left + m_PictureRect.Left
  m_PictureRect.Right = m_Rect.Left + sZPicture.cx
  m_PictureRect.Top = m_Rect.Top + m_PictureRect.Top
  m_PictureRect.Bottom = m_Rect.Bottom + sZPicture.cy
End Sub

Friend Function HitTest(pT As POINTAPI) As STAT_HIT_TEST
   HitTest = SHT_NO_HIT
   If Not m_Picture Is Nothing Then
    If PtInRect(m_PictureRect, pT.x, pT.y) <> 0 Then
      HitTest = SHT_BITMAP
      Exit Function
    End If
    If PtInRect(m_Rect, pT.x, pT.y) <> 0 Then HitTest = SHT_PANEL
   Else
    If PtInRect(m_Rect, pT.x, pT.y) <> 0 Then HitTest = SHT_PANEL
   End If
End Function

Friend Sub Draw(ByVal hdc As Long)
  Dim r As RECT
  Dim pT As POINTAPI
  Dim hMemDC As Long, hBmpOld As Long, hBrushOld As Long, hPenOld As Long, hFontOld As Long
  Dim sZPanelRect As SIZE
  Dim sZ As SIZE
      
  On Error GoTo draw_Err
  'set up
  hMemDC = CreateCompatibleDC(hdc)
  If hMemDC = 0 Then Exit Sub
  sZPanelRect = GetRectDimensions(m_Rect)
  
  'open draw
  hBmpOld = SelectObject(hMemDC, CreateCompatibleBitmap(hdc, sZPanelRect.cx, sZPanelRect.cy))
  hBrushOld = SelectObject(hMemDC, CreateSolidBrush(m_BackColor))
  hPenOld = SelectObject(hMemDC, CreatePen(PS_SOLID, 1, m_BackColor))
  
  ' Windows expects the font size to be in pixels and to
  ' be negative if you are specifying the character height
  ' you want.
  hFontOld = SelectObject(hMemDC, CreateFont((m_Font.SIZE * -20) / Screen.TwipsPerPixelY, 0, 0, 0, IIf(m_Font.Bold, FW_BOLD, FW_NORMAL), m_Font.Italic, m_Font.Underline, m_Font.Strikethrough, 0, 0, 0, 0, 0, m_Font.Name))
  
  Call SetBkColor(hMemDC, m_BackColor)
  Call SetTextColor(hMemDC, m_ForeColor)
  Call Rectangle(hMemDC, 0, 0, sZPanelRect.cx, sZPanelRect.cy)
  Call DeleteObject(SelectObject(hMemDC, hPenOld))
  r.Right = sZPanelRect.cx
  r.Bottom = sZPanelRect.cy
  'draw panel
  Call Draw3DRect(hMemDC, r, Style)
  r.Left = r.Left + 1
  r.Right = r.Right - 1
  r.Bottom = r.Bottom - 1
  r.Top = r.Top + 1
  Call DrawPicture(r, hMemDC)
  pT = GetCenterTextPosition(hMemDC, m_Caption, r)
  Call ExtTextOut(hMemDC, r.Left + 2, pT.y, ETO_CLIPPED, r, m_Caption, Len(m_Caption), 0)
  'end draw panel
  
  Call BitBlt(hdc, m_Rect.Left, m_Rect.Top, sZPanelRect.cx, sZPanelRect.cy, hMemDC, 0, 0, SRCCOPY)
  
  'close draw
  Call DeleteObject(SelectObject(hMemDC, hBrushOld))
  Call DeleteObject(SelectObject(hMemDC, hFontOld))
  Call DeleteObject(SelectObject(hMemDC, hBmpOld))
  
  Call DeleteDC(hMemDC)
  
draw_Err:
  Exit Sub
End Sub

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
  Call DrawMe
End Sub
