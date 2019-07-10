Attribute VB_Name = "ColorFunctions"
Option Explicit
Public Const LINEWIDTH_CONST As Single = 10
Private Const LINESPACE_CONST As Single = 30
Public LINESPACE As Single
Public LINEWIDTH As Single

Public COLOR_WHITE As Long ' initailised in ReporterClass Initialise
Private m_STACKTOP As Long
Private m_STACKMAX As Long
Private m_StyleStack() As ColorType

Public Sub ResetColors()
  ReportControl.FColor.ColorValid = False
  ReportControl.FColor.ColorSet = False
  ReportControl.FColor.ForeColor = 0 'Black
  ReportControl.FColor.LineColor = 0 'Black
  ReportControl.FColor.ColorFill = False
  ReportControl.FColor.FillColor = COLOR_WHITE
  ReportControl.FColor.FillStyle = vbFSTransparent
End Sub

Public Sub SetColor()
  Dim InError As Boolean
  Dim fh As Single, dw As Single
  
  Call xSet("SetColor")
  On Error GoTo SetColor_err
SetColor_retry:
  If NoGraphics Then GoTo SetColor_end
  If Not ReportControl.FColor.ColorValid Then Call ResetColors
  If ReportControl.FColor.ColorSet Then GoTo SetColor_end
  
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    Call SetPrinterMode(FONT_TRANSPARANT)
    Printer.FillColor = ReportControl.FColor.FillColor
    Printer.FillStyle = ReportControl.FColor.FillStyle
    Printer.ForeColor = ReportControl.FColor.ForeColor
    dw = Printer.ScaleY(LINEWIDTH, Printer.ScaleMode, vbPixels)
    If dw < 1 Then dw = 1
    Printer.DrawWidth = dw
    LINESPACE = LINESPACE_CONST
    Printer.DrawStyle = vbSolid
    Printer.DrawMode = vbCopyPen
  ElseIf ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE Then
    ReportControl.Preview.FontTransparent = True
    ReportControl.Preview.FillColor = ReportControl.FColor.FillColor
    ReportControl.Preview.FillStyle = ReportControl.FColor.FillStyle
    ReportControl.Preview.ForeColor = ReportControl.FColor.ForeColor
    dw = ReportControl.Zoom * ReportControl.Preview.ScaleY(LINEWIDTH, ReportControl.Preview.ScaleMode, vbPixels) / 100!
    If dw < 1 Then dw = 1
    ReportControl.Preview.DrawWidth = dw
    dw = ReportControl.Zoom * LINESPACE_CONST / 100!
    If dw < 1 Then dw = 1
    LINESPACE = dw
    ReportControl.Preview.DrawStyle = vbSolid
    ReportControl.Preview.DrawMode = vbCopyPen
  End If
  ReportControl.FColor.ColorValid = True
  ReportControl.FColor.ColorSet = True
  
SetColor_end:
  Call xReturn("SetColor")
  Exit Sub
  
SetColor_err:
  If InError Then Resume SetColor_end
  InError = True
  ReportControl.FColor.ColorValid = False
  ReportControl.FColor.ColorSet = False
  Resume SetColor_err
End Sub

Public Sub PushColor()
  Call xSet("PushColor")
  
  m_STACKTOP = m_STACKTOP + 1
  If m_STACKTOP > m_STACKMAX Then
    m_STACKMAX = m_STACKMAX + 1
    ReDim Preserve m_StyleStack(1 To m_STACKMAX) As ColorType
  End If
  m_StyleStack(m_STACKTOP).ColorValid = ReportControl.FColor.ColorValid
  m_StyleStack(m_STACKTOP).ColorSet = ReportControl.FColor.ColorSet
  m_StyleStack(m_STACKTOP).ColorFill = ReportControl.FColor.ColorFill
  m_StyleStack(m_STACKTOP).LineColor = ReportControl.FColor.LineColor
  m_StyleStack(m_STACKTOP).ForeColor = ReportControl.FColor.ForeColor
  m_StyleStack(m_STACKTOP).FillColor = ReportControl.FColor.FillColor
  m_StyleStack(m_STACKTOP).FillStyle = ReportControl.FColor.FillStyle
  Call xReturn("PushColor")
End Sub

Public Sub PopColor(Optional ByVal bSetColor As Boolean = True)
  Call xSet("PopColor")
  If m_STACKTOP > 0 Then
    ReportControl.FColor.ForeColor = m_StyleStack(m_STACKTOP).ForeColor
    ReportControl.FColor.FillColor = m_StyleStack(m_STACKTOP).FillColor
    ReportControl.FColor.FillStyle = m_StyleStack(m_STACKTOP).FillStyle
    ReportControl.FColor.ColorFill = m_StyleStack(m_STACKTOP).ColorFill
    ReportControl.FColor.LineColor = m_StyleStack(m_STACKTOP).LineColor
    ReportControl.FColor.ColorValid = True
    ReportControl.FColor.ColorSet = bSetColor
    Call SetColor
    m_STACKTOP = m_STACKTOP - 1
  Else
    Call Err.Raise(ERR_POPCOLORSTACK, "PopColor", "Pop without Push!")
  End If
  Call xReturn("PopColor")
End Sub

Public Function ProcessFillColorRGB(rgb As Long) As Long
  If NoGraphics() Then Exit Function
  ProcessFillColorRGB = ReportControl.FColor.FillColor
  If rgb = COLOR_WHITE Then
    ReportControl.FColor.ColorFill = False
    ReportControl.FColor.FillStyle = 1  'vbFSTransparent
    ReportControl.FColor.LineColor = 0& 'Black
  Else
    ReportControl.FColor.ColorFill = True
    ReportControl.FColor.FillStyle = 0  'vbFSSolid
    ReportControl.FColor.LineColor = rgb
  End If
  ReportControl.FColor.FillColor = rgb
  ReportControl.FColor.ColorValid = True
  ReportControl.FColor.ColorSet = False
End Function

Public Function ProcessForeColorRGB(rgb As Long) As Long
  If NoGraphics() Then Exit Function
  ProcessForeColorRGB = ReportControl.FColor.ForeColor
  ReportControl.FColor.ForeColor = rgb
  ReportControl.FColor.LineColor = rgb
  ReportControl.FColor.ColorValid = True
  ReportControl.FColor.ColorSet = False
End Function

