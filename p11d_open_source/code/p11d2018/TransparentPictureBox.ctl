VERSION 5.00
Begin VB.UserControl TransparentPictureBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Height          =   1515
      Left            =   3525
      ScaleHeight     =   1455
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   1650
      Width           =   915
   End
End
Attribute VB_Name = "TransparentPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" _
    (ByVal OLE_COLOR As Long, ByVal hPalette As Long, _
    pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Const MASK_COLOR As Long = 16776960
Private m_MaskColor As OLE_COLOR
Public Property Get Picture() As StdPicture
  Set Picture = Picture1.Picture
End Property
Public Property Let MaskColor(ByVal NewValue As OLE_COLOR)
  m_MaskColor = NewValue
  PropertyChanged "MaskColor"
End Property
Public Property Get MaskColor() As OLE_COLOR
  MaskColor = m_MaskColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
  UserControl.BackColor = NewValue
  PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Set Picture(ByVal NewValue As StdPicture)
  Set Picture1.Picture = NewValue
  Picture1.width = Picture1.Picture.width
  Picture1.height = Picture1.Picture.height
  
  Call PaintImage
End Property
Private Sub UserControl_Initialize()
  Picture1.Visible = False
  m_MaskColor = 16776960
  Call PaintImage
End Sub
Private Sub UserControl_Paint()
  Call PaintImage
End Sub
Private Sub PaintImage()
  'UserControl.Line (0, 0)-(UserControl.width, UserControl.height), UserControl.BackColor, BF
  Call TransparentBlt(Picture1, TranslateColor(m_MaskColor))
End Sub
Sub TransparentBlt(pic As StdPicture, TransColor As Long)
    Dim maskDC As Long      'DC for the mask
    Dim tempDC As Long      'DC for temporary data
    Dim hMaskBmp As Long    'Bitmap for mask
    Dim hTempBmp As Long    'Bitmap for temporary data
    Dim srchDC As Long
    Dim srcBitmap As Long
    Dim width As Long, height As Long
    Dim dsthDC As Long
    
    width = UserControl.ScaleX(pic.width, vbHimetric, vbPixels)
    height = UserControl.ScaleY(pic.height, vbHimetric, vbPixels)
    
    srchDC = CreateCompatibleDC(dsthDC)
    
    Call SelectObject(srchDC, pic.Handle)
    'First, create some DC's. These are our gateways to associated
    dsthDC = UserControl.hdc
    'bitmaps in RAM
    maskDC = CreateCompatibleDC(dsthDC)
    tempDC = CreateCompatibleDC(dsthDC)
    
    
    'Then, we need the bitmaps. Note that we create a monochrome
    'bitmap here!
    'This is a trick we use for creating a mask fast enough.
    hMaskBmp = CreateBitmap(width, height, 1, 1, ByVal 0&)
    hTempBmp = CreateCompatibleBitmap(dsthDC, width, height)
    
    'Then we can assign the bitmaps to the DCs
    hMaskBmp = SelectObject(maskDC, hMaskBmp)
    hTempBmp = SelectObject(tempDC, hTempBmp)

    'Now we can create a mask. First, we set the background color
    'to the transparent color; then we copy the image into the
    'monochrome bitmap.
    'When we are done, we reset the background color of the
    'original source.
    TransColor = SetBkColor(srchDC, TransColor)
    BitBlt maskDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
    TransColor = SetBkColor(srchDC, TransColor)

    'The first we do with the mask is to MergePaint it into the
    'destination.
    'This will punch a WHITE hole in the background exactly were
    'we want the graphics to be painted in.
    BitBlt tempDC, 0, 0, width, height, maskDC, 0, 0, vbSrcCopy
    BitBlt dsthDC, 0, 0, width, height, tempDC, 0, 0, vbMergePaint
      
    'Now we delete the transparent part of our source image. To do
    'this, we must invert the mask and MergePaint it into the
    'source image. The transparent area will now appear as WHITE.
    BitBlt maskDC, 0, 0, width, height, maskDC, 0, 0, vbNotSrcCopy
    BitBlt tempDC, 0, 0, width, height, srchDC, 0, 0, vbSrcCopy
    BitBlt tempDC, 0, 0, width, height, maskDC, 0, 0, vbMergePaint

    'Both target and source are clean. All we have to do is to AND
    'them together!
    BitBlt dsthDC, 0, 0, width, height, tempDC, 0, 0, vbSrcAnd
    'Now all we have to do is to clean up after us and free system
    'resources..
    DeleteObject (hMaskBmp)
    DeleteObject (hTempBmp)
    
    DeleteDC (srchDC)
    
    DeleteDC (maskDC)
    DeleteDC (tempDC)
End Sub


Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_MaskColor = PropBag.ReadProperty("MaskColor", MASK_COLOR)
  Set Picture1.Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("MaskColor", m_MaskColor, MASK_COLOR)
  Call PropBag.WriteProperty("Picture", Picture1.Picture, Nothing)
End Sub
