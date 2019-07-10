VERSION 5.00
Begin VB.UserControl SPLIT 
   Alignable       =   -1  'True
   BackColor       =   &H000000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   EditAtDesignTime=   -1  'True
   MousePointer    =   9  'Size W E
   ScaleHeight     =   3600
   ScaleWidth      =   1035
   ToolboxBitmap   =   "TCSSPLIT.ctx":0000
End
Attribute VB_Name = "SPLIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'events
Event FinishedSplit(ByVal MovementInTwips As Single)
'properties
Private m_lMinBorderPixels As Long
Private m_bMDIForm As Boolean

Private Const S_MIN_BORDER_PIXELS As String = "MinBorderPixels"
Private Const L_MIN_BORDER_PIXELS_DEF As Long = 10
'other
Private SI As SPLIT_INFO
Private Const L_SPLIT_WIDTH_PIXELS As Long = 5
Public Sub Initialise(hWndClient As Long, bMDIForm As Boolean)
  m_bMDIForm = bMDIForm
  If m_bMDIForm Then
    SI.hWndClient = GetWindow(hWndClient, GW_CHILD)
  Else
    SI.hWndClient = hWndClient
  End If
  SI.hWndParent = GetParent(UserControl.hWnd)
  
  Call AddControl(Me)
  If Not m_bMDIForm Then Call SizeMeToParent
End Sub
Private Sub DrawBar(r As RECT)
  Dim hDCMem As Long
  Dim hBitmapMemOld As Long
  Dim hBrushMemOld As Long
  Dim hPenOld As Long
  Dim sZ As SIZEL
  
  sZ = RectDimensions(r)
  'create memory bitmap
  hDCMem = CreateCompatibleDC(SI.hDCScreen)
  hBitmapMemOld = SelectObject(hDCMem, CreateCompatibleBitmap(SI.hDCScreen, sZ.cx, sZ.cy))
  hPenOld = SelectObject(hDCMem, CreatePen(PS_NULL, 0, 0))
  'draw to memory bitmap , first select brush
  hBrushMemOld = SelectObject(hDCMem, CreateHatchBrush(HS_DIAGCROSS, 0))
  Call Rectangle(hDCMem, 0, 0, sZ.cx + 1, sZ.cy + 1)
  Call BitBlt(SI.hDCScreen, r.Left, r.Top, sZ.cx, sZ.cy, hDCMem, 0, 0, SRCINVERT)
  Call DeleteObject(SelectObject(hDCMem, hBrushMemOld))
  Call DeleteObject(SelectObject(hDCMem, hPenOld))
  Call DeleteObject(SelectObject(hDCMem, hBitmapMemOld))
  Call DeleteDC(hDCMem)
End Sub
Private Sub DrawSplit(SMM As SPLIT_MOUSE_MODE)
  Select Case SMM
    Case SPLIT_MOUSE_MODE.SMM_DOWN
      Call DrawBar(SI.SplitRect)
    Case SPLIT_MOUSE_MODE.SMM_MOVE
      Call DrawBar(SI.SplitRect)
      Call DrawBar(SI.LastSplitRect)
    Case SPLIT_MOUSE_MODE.SMM_UP
      Call DrawBar(SI.SplitRect)
  End Select
End Sub
Public Property Get hWndParent() As Long
  hWndParent = SI.hWndParent
End Property
Public Property Get hWndClient() As Long
  hWndClient = SI.hWndClient
End Property
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim r As RECT
  Dim pT As POINTAPI
  
  
  Call SetCapture(UserControl.hWnd)
  
  If Button = vbLeftButton Then
    
    With SI
      Call GetWindowRect(UserControl.hWnd, r)
      .SplitRect = r
      .StartSplitRect = r
      .SplitDimensions = RectDimensions(.SplitRect)
      .MousePosition = TranslateCoords(x, y)
      pT = ConvertCoords(x, y)
      .XMouseMaxUser = (.XMouseMax) - (pT.x)
      .XMouseMinUser = .XMouseMin + (pT.x)
      .hDCScreen = CreateDC("DISPLAY", 0, 0, 0)
    End With
    
    Call DrawSplit(SMM_DOWN)
  End If
  
End Sub
Private Function TranslateCoords(x As Single, y As Single) As POINTAPI
  TranslateCoords = TranslateCoordsEx(x, y, UserControl.hWnd, 0)
End Function
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim pT As POINTAPI
  Dim lDiff As Long
  Dim r As RECT
  
  
  If Button = vbLeftButton Then
    pT = TranslateCoords(x, y)
    With SI
      .LastSplitRect = .SplitRect
      If pT.x > .XMouseMaxUser Then
        pT.x = .XMouseMaxUser
        r = .SplitRect
        r.Left = .XMouseMax - .SplitDimensions.cx
        r.Right = .XMouseMax
        If r.Left <> .SplitRect.Left Then
          .SplitRect = r
          Call DrawSplit(SMM_MOVE)
        End If
      ElseIf pT.x < .XMouseMinUser Then
        pT.x = .XMouseMinUser
        .MousePosition = pT
        r = .SplitRect
        r.Left = .XMouseMin
        r.Right = r.Left + .SplitDimensions.cx
        If r.Left <> .SplitRect.Left Then
          .SplitRect = r
          Call DrawSplit(SMM_MOVE)
        End If
      Else
        lDiff = pT.x - .MousePosition.x
        .SplitRect.Left = .SplitRect.Left + lDiff
        .SplitRect.Right = .SplitRect.Right + lDiff
        Call DrawSplit(SMM_MOVE)
      End If
      .MousePosition = pT
    End With
  ElseIf Button = vbRightButton Then
    pT = TranslateCoords(x, y)
    If PtInRect(SI.StartSplitRect, pT.x, pT.y) = 0 Then
      Call ReleaseCapture
    End If
  End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lDiff As Long
  Dim rParent As RECT, rClient As RECT
  Dim sZ As SIZEL
  
  If Button = vbLeftButton Then
    Call DrawSplit(SMM_UP)
    Call ReleaseCapture
    Call DeleteDC(SI.hDCScreen)
    
    lDiff = SI.SplitRect.Left - SI.StartSplitRect.Left
    
    If m_bMDIForm Then
      'resize the parent window
      Call GetClientRect(SI.hWndParent, rParent)
      sZ = RectDimensions(rParent)
      Call MoveWindow(SI.hWndParent, rParent.Left + lDiff, rParent.Top, sZ.cx + lDiff, sZ.cy, 1)
    Else
      Call GetClientRect(UserControl.hWnd, rClient)
      Call MapWindowPoints(UserControl.hWnd, SI.hWndParent, rClient, 2)
      sZ = RectDimensions(rClient)
      Call MoveWindow(UserControl.hWnd, rClient.Left + lDiff, rClient.Top, sZ.cx, sZ.cy, 1)
    End If
    
    RaiseEvent FinishedSplit(lDiff * Screen.TwipsPerPixelX)
  End If
End Sub
Private Sub UserControl_Terminate()
  Call RemoveControl(Me)
End Sub
Public Sub SizeMeToParent()
  Dim r As RECT
  Dim sZ As SIZEL
  Dim pT As POINTAPI
  
  If m_bMDIForm Then
    Call GetWindowRect(SI.hWndParent, r)
    sZ = RectDimensions(r)
    Call MoveWindow(UserControl.hWnd, sZ.cx - L_SPLIT_WIDTH_PIXELS, 0, L_SPLIT_WIDTH_PIXELS, sZ.cy, 1)
    'left, in screen coordinates as NULL passed as hWndTo
  End If
  
  pT.x = 0
  Call MapWindowPoints(SI.hWndParent, 0, pT.x, 1)
  SI.XMouseMin = pT.x + m_lMinBorderPixels
  'right
  Call GetWindowRect(SI.hWndClient, r)
  SI.XMouseMax = r.Right - m_lMinBorderPixels

End Sub
Public Property Get MinBorderPixels() As Long
  MinBorderPixels = m_lMinBorderPixels
End Property
Public Property Let MinBorderPixels(NewValue As Long)
  m_lMinBorderPixels = NewValue
  Call PropertyChanged(S_MIN_BORDER_PIXELS)
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty(S_MIN_BORDER_PIXELS, m_lMinBorderPixels, L_MIN_BORDER_PIXELS_DEF)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
  m_lMinBorderPixels = PropBag.ReadProperty(S_MIN_BORDER_PIXELS, L_MIN_BORDER_PIXELS_DEF)
  If Ambient.UserMode = True Then UserControl.BackColor = &HC0C0C0 'light grey
End Sub
Private Sub UserControl_InitProperties()
  m_lMinBorderPixels = L_MIN_BORDER_PIXELS_DEF
End Sub
Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property
