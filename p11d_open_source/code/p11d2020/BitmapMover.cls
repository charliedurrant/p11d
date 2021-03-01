VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Private Type SIZE
  Width As Long
  Height As Long
End Type

Private Type POINTL
  X As Long
  Y As Long
End Type

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long



Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT = &HEE0086



Private Type BLIT_INFO
  hDCMem As Long
  hBitmapOld As Long
  sZ As SIZE
  rMemory As RECT
  rScreen As RECT
End Type

Private Type MOVE_INFO
  hWnd As Long
  bStart As Boolean
  rToBlitTo As RECT
  sZToBlitTo As SIZE
  BMITemplate As BLIT_INFO
  BMIPicture As BLIT_INFO
  BMIBackground As BLIT_INFO
  BMIScreen As BLIT_INFO
End Type

Private MI As MOVE_INFO

Public Sub Class_Terminate()
  Call Terminate
End Sub
Public Sub Terminate()
  With MI
    If .hWnd <> 0 Then
      Call TeminateBitmapInfo(.BMIBackground)
      Call TeminateBitmapInfo(.BMIPicture)
      Call TeminateBitmapInfo(.BMIScreen)
      Call TeminateBitmapInfo(.BMITemplate)
      .hWnd = 0
    End If
  End With
End Sub
Private Function ConvertCords(X As Single, Y As Single) As POINTL
  ConvertCords.X = X / Screen.TwipsPerPixelX
  ConvertCords.Y = Y / Screen.TwipsPerPixelY
End Function
Private Function GetPicDimensions(IPD As IPictureDisp) As SIZE
  GetPicDimensions.Width = HimetricToTwips(IPD.Width) / Screen.TwipsPerPixelX
  GetPicDimensions.Height = HimetricToTwips(IPD.Height) / Screen.TwipsPerPixelY
End Function
Private Function HimetricToTwips(lHimetric) As Single
  'see scaleX in vb help
  HimetricToTwips = (567 / 1000) * lHimetric
End Function

Public Sub Initialise(Template As IPictureDisp, Picture As IPictureDisp, ByVal LeftInTwips As Single, ByVal TopInTwips As Single, ByVal hWndToDrawTo As Long)
  If (Template Is Nothing) Or (Picture Is Nothing) Or (hWndToDrawTo = 0) Then Exit Sub
  
  'clean up previous
  Call Terminate
  
  MI.hWnd = hWndToDrawTo
  
  Call CreateMemBitmapFromIPictureDisp(Picture, MI.BMIPicture, MI.hWnd)
  Call CreateMemBitmapFromIPictureDisp(Template, MI.BMITemplate, MI.hWnd)
  'create background with same dimensions
  Call CreateMemBitmapFromIPictureDisp(Template, MI.BMIBackground, MI.hWnd)
  
  MI.bStart = True
  Call Move(LeftInTwips, TopInTwips)
  
End Sub

Public Sub Move(ByVal LeftInTwips As Single, ByVal TopInTwips As Single)
  Dim hDCDest As Long
  Dim pT As POINTL
  
  
  With MI
    hDCDest = GetDC(MI.hWnd)
  
    pT = ConvertCords(LeftInTwips, TopInTwips)
    
    If Not MI.bStart Then
      .BMIScreen.rScreen = MoveRectToPoint(.BMIScreen.rScreen, pT)
      If Not RectInRect(.BMIScreen.rScreen, .BMIBackground.rScreen) Then
        GoTo NORMAL_BLIT
      Else
        'create mem bitmap size of previous rect and new rect if overlap - stops flicker
        .rToBlitTo = RectFrom2Rects(.BMIScreen.rScreen, .BMIBackground.rScreen)
        .sZToBlitTo = RectDimensions(.rToBlitTo)
        Call CreateMemBitmapFromRect(.BMIScreen, .rToBlitTo)
        'copy current screen in to memory from r.top, r.left
        Call BitBlt(.BMIScreen.hDCMem, 0, 0, .sZToBlitTo.Width, .sZToBlitTo.Height, hDCDest, .rToBlitTo.left, .rToBlitTo.top, SRCCOPY)
        'get the background rect and screen rect relative to new larger rect
        Call RebaseRects(.BMIScreen.rMemory, .BMIBackground.rMemory, .BMIScreen.rScreen, .BMIBackground.rScreen, .rToBlitTo)
        'copy previous background in at .BMIBackbrounf.rMemory
        Call BitBlt(.BMIScreen.hDCMem, .BMIBackground.rMemory.left, .BMIBackground.rMemory.top, .BMIBackground.sZ.Width, .BMIBackground.sZ.Height, .BMIBackground.hDCMem, 0, 0, SRCCOPY)
        'copy new background out at memory screen
        Call BitBlt(.BMIBackground.hDCMem, 0, 0, .BMIBackground.sZ.Width, .BMIBackground.sZ.Height, .BMIScreen.hDCMem, .BMIScreen.rMemory.left, .BMIScreen.rMemory.top, SRCCOPY)
        'copy template to screen
        Call BitBlt(.BMIScreen.hDCMem, .BMIScreen.rMemory.left, .BMIScreen.rMemory.top, .BMITemplate.sZ.Width, .BMITemplate.sZ.Height, .BMITemplate.hDCMem, 0, 0, SRCPAINT)
        'copy pict to screem
        Call BitBlt(.BMIScreen.hDCMem, .BMIScreen.rMemory.left, .BMIScreen.rMemory.top, .BMIPicture.sZ.Width, .BMIPicture.sZ.Height, .BMIPicture.hDCMem, 0, 0, SRCAND)
        .BMIBackground.rScreen = .BMIScreen.rScreen
      End If
      
    Else
NORMAL_BLIT:
      'where do we blit back to
      MI.rToBlitTo = MoveRectToPoint(MI.BMIBackground.rMemory, pT)
      MI.sZToBlitTo = RectDimensions(MI.rToBlitTo)
    
      'set up the rects
      If MI.bStart Then
        MI.BMIBackground.rScreen = MoveRectToPoint(MI.BMIBackground.rMemory, pT)
      Else
        'COPY background back
        Call BitBlt(hDCDest, .BMIBackground.rScreen.left, .BMIBackground.rScreen.top, .BMIBackground.sZ.Width, .BMIBackground.sZ.Height, .BMIBackground.hDCMem, 0, 0, SRCCOPY)
        MI.BMIBackground.rScreen = MoveRectToPoint(MI.BMIBackground.rScreen, pT)
      End If
      MI.BMIScreen.rScreen = MI.BMIBackground.rScreen
      'create mem bitmap the size of the picture
      Call CreateMemBitmapFromRect(MI.BMIScreen, MI.BMIBackground.rMemory)
      'copy background out
      Call BitBlt(MI.BMIBackground.hDCMem, 0, 0, MI.BMIBackground.sZ.Width, .BMIBackground.sZ.Height, hDCDest, pT.X, pT.Y, SRCCOPY)
      'copy bakground to screenmem dc
      Call BitBlt(.BMIScreen.hDCMem, 0, 0, .BMIBackground.sZ.Width, .BMIBackground.sZ.Height, .BMIBackground.hDCMem, 0, 0, SRCCOPY)
      'copy template to screen
      Call BitBlt(.BMIScreen.hDCMem, 0, 0, .BMIBackground.sZ.Width, .BMIBackground.sZ.Height, .BMITemplate.hDCMem, 0, 0, SRCPAINT)
      'copy pict to screem
      Call BitBlt(.BMIScreen.hDCMem, 0, 0, .BMIBackground.sZ.Width, .BMIBackground.sZ.Height, .BMIPicture.hDCMem, 0, 0, SRCAND)
      
      MI.bStart = False
    End If
  
    Call BitBlt(hDCDest, MI.rToBlitTo.left, MI.rToBlitTo.top, MI.sZToBlitTo.Width, MI.sZToBlitTo.Height, MI.BMIScreen.hDCMem, 0, 0, SRCCOPY)
    Call TeminateBitmapInfo(.BMIScreen)
    Call ReleaseDC(.hWnd, hDCDest)
  End With
End Sub

Private Function RectInRect(r1 As RECT, r2 As RECT) As Boolean
  Dim hRgn As Long
  
  hRgn = CreateRectRgn(r2.left, r2.top, r2.right, r2.bottom)
  RectInRect = RectInRegion(hRgn, r1)
  DeleteObject (hRgn)
End Function

Private Function RectDimensions(r As RECT) As SIZE
  RectDimensions.Height = r.bottom - r.top
  RectDimensions.Width = r.right - r.left
End Function

Private Sub CreateMemBitmapFromRect(BI As BLIT_INFO, r As RECT)
  Dim hDC As Long
  
  hDC = GetDC(MI.hWnd)
  BI.hDCMem = CreateCompatibleDC(hDC)
  BI.sZ = RectDimensions(r)
  BI.rMemory = r
  BI.hBitmapOld = SelectObject(BI.hDCMem, CreateCompatibleBitmap(hDC, BI.sZ.Width, BI.sZ.Height))
  Call ReleaseDC(MI.hWnd, hDC)
End Sub

Private Sub CreateMemBitmapFromIPictureDisp(IPD As IPictureDisp, BI As BLIT_INFO, hWnd As Long)
  Dim hDC As Long
  Dim hBitmapOldTemp As Long
  Dim hDCMemTemp As Long
  
  hDC = GetDC(hWnd)
  BI.sZ = GetPicDimensions(IPD)
  BI.rMemory.top = 0
  BI.rMemory.left = 0
  BI.rMemory.right = BI.sZ.Width
  BI.rMemory.bottom = BI.sZ.Height
  
  hDCMemTemp = CreateCompatibleDC(hDC)
  hBitmapOldTemp = SelectObject(hDCMemTemp, IPD.Handle)
  
  BI.hDCMem = CreateCompatibleDC(hDC)
  BI.hBitmapOld = SelectObject(BI.hDCMem, CreateCompatibleBitmap(hDC, BI.sZ.Width, BI.sZ.Height))
  Call BitBlt(BI.hDCMem, 0, 0, BI.sZ.Width, BI.sZ.Height, hDCMemTemp, 0, 0, SRCCOPY)
  Call DeleteObject(SelectObject(hDCMemTemp, hBitmapOldTemp))
  
  'this should work but flips bitmap???
  'BI.hBitmapOld = SelectObject(BI.hDCMem, CreateCompatibleBitmap(hDC, BI.sZ.Width, BI.sZ.Height))
  'BI.hDCMem = CreateCompatibleDC(hDC)
  'copy into memoryI
  'Call IPD.Render(CLng(BI.hDCMem), 0&, 0&, CLng(BI.sZ.Width), CLng(BI.sZ.Height), 0&, 0&, IPD.Width, IPD.Height, 0&)
End Sub

Private Sub TeminateBitmapInfo(BI As BLIT_INFO)
  If BI.hDCMem <> 0 Then
    DeleteObject (SelectObject(BI.hDCMem, BI.hBitmapOld))
    DeleteDC (BI.hDCMem)
    BI.hDCMem = 0
    BI.hBitmapOld = 0
  End If
End Sub

Private Function RectFrom2Rects(r1 As RECT, r2 As RECT) As RECT
  If (r2.left > r1.left) Then
    If r2.top < r1.top Then
      'type A
      RectFrom2Rects.left = r1.left
      RectFrom2Rects.top = r2.top
      RectFrom2Rects.right = r2.right
      RectFrom2Rects.bottom = r1.bottom
    Else
      'type b
      RectFrom2Rects.left = r1.left
      RectFrom2Rects.top = r1.top
      RectFrom2Rects.right = r2.right
      RectFrom2Rects.bottom = r2.bottom
    End If
  Else
    If r2.top < r1.top Then
      'type c
      RectFrom2Rects.left = r2.left
      RectFrom2Rects.top = r2.top
      RectFrom2Rects.right = r1.right
      RectFrom2Rects.bottom = r1.bottom
    Else
      'type d
      RectFrom2Rects.left = r2.left
      RectFrom2Rects.top = r1.top
      RectFrom2Rects.right = r1.right
      RectFrom2Rects.bottom = r2.bottom
    End If
  End If
End Function
Private Function ReduceRectBySize(r As RECT, sZ As SIZE) As RECT
  ReduceRectBySize.top = r.top - sZ.Height
  ReduceRectBySize.bottom = r.bottom - sZ.Height
  
  ReduceRectBySize.left = r.left - sZ.Width
  ReduceRectBySize.right = r.right - sZ.Width
End Function
Private Sub RebaseRects(rChild1New As RECT, rChild2New As RECT, rChildOld1 As RECT, rChildOld2 As RECT, rParentOf2Olds As RECT)
  Dim sZ As SIZE
  
  sZ.Height = rParentOf2Olds.top
  sZ.Width = rParentOf2Olds.left
  
  rChild1New = ReduceRectBySize(rChildOld1, sZ)
  rChild2New = ReduceRectBySize(rChildOld2, sZ)
  
End Sub
Private Function RebaseRect(r As RECT) As RECT
  Dim sZ As SIZE
  
  sZ = RectDimensions(r)
  RebaseRect = ReduceRectBySize(r, sZ)
End Function
Private Function MoveRectToPoint(r As RECT, pT As POINTL) As RECT
  Dim sZ As SIZE
  
  sZ = RectDimensions(r)
  MoveRectToPoint.left = pT.X
  MoveRectToPoint.top = pT.Y
  MoveRectToPoint.right = pT.X + sZ.Width
  MoveRectToPoint.bottom = pT.Y + sZ.Height
End Function
Private Function RectToMemoryRect(r As RECT) As RECT
  Dim pT As POINTL
  
  RectToMemoryRect = MoveRectToPoint(r, pT)
End Function
