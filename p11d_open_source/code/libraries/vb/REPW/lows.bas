Attribute VB_Name = "lows"
Option Explicit

'apf 2008
Public Const NOPRINT_AREA_BL As Single = 10  ' leave last 10 twips in X,Y dir blank

Public Enum BKMODE_TYPE
  FONT_TRANSPARANT = 1
  FONT_OPAQUE = 1
End Enum
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Public Function GetPercent(ByVal xPercent As Single, ByVal limit As Single) As Single
  If Not IsExportHTML Then
    If (xPercent < 0) Or (xPercent > 100) Then Call Err.Raise(ERR_PERCENT, "GetPercent", "Expected percent found " & CStr(xPercent))
  End If
  GetPercent = (xPercent * limit) / 100!
  If GetPercent < 1! Then GetPercent = 0!
End Function
  
Public Function GetRelativePercent(ByVal x As Single, ByVal xPercent As Single, ByVal limit As Single) As Single
  'apf limit - 1 used since coord system 1 .. maxwidth, => GetRelativePercent(1,100) = MaxWidth
  'RK handle/log errors
  If xPercent < 0 Then
    x = Int(x + -1! * GetPercent(-xPercent, limit - 1))
  Else
    x = Int(x + GetPercent(xPercent, limit - 1))
  End If
  If (x < 1) Or (x > limit) Then
    'Call Err.Raise(ERR_PRINTRANGE, "GetRelativePercent", "Attempting to print outside the current page at " & CStr(x) & " where limit is 1 to " & CStr(limit))
    Call HandleRangeErrors(x, limit)
  End If
  GetRelativePercent = x
End Function

Public Function GetRelative(ByVal x As Single, ByVal inc As Single, ByVal limit As Single) As Single
  'RK handle/log errors
  x = Int(x + inc)
  If Not IsExportHTML Then
    If (x < 1) Or (x > limit) Then
       Call HandleRangeErrors(x, limit)
'      Call Err.Raise(ERR_PRINTRANGE, "GetRelative", "Attempting to print outside the current page at " & CStr(x) & " where limit is 1 to " & CStr(limit))
    End If
  End If
  GetRelative = x
End Function
  
Public Function GetAbsolute(ByVal x As Single, ByVal limit As Single) As Single
  'RK handle/log errors
  x = Int(x)
  If (x < 1) Or (x > limit) Then
    Call HandleRangeErrors(x, limit)
    'Call Err.Raise(ERR_PRINTRANGE, "GetAbsolute", "Attempting to print outside the current page at " & CStr(x) & " where limit is 1 to " & CStr(limit))
  End If
  GetAbsolute = x
End Function
Public Sub HandleRangeErrors(ByRef x As Single, ByRef limit As Single)
  'RK reset x to inside range and log as silent error
  If (x < 1) Then
   Call ErrorMessage(ERR_ERRORSILENT, Err, "HandleRangeErrors", "ERR_PRINTRANGE", "GetRelativePercent (x < 1): Attempting to print outside the current page at " & CStr(x) & " where limit is 1 to " & CStr(limit))
   x = 1
  ElseIf (x > limit) Then
   Call ErrorMessage(ERR_ERRORSILENT, Err, "HandleRangeErrors", "ERR_PRINTRANGE", "GetRelativePercent (x > limit): Attempting to print outside the current page at " & CStr(x) & " where limit is 1 to " & CStr(limit))
   x = limit
  End If
End Sub
  
Public Sub SetZoomLimit()
  If ReportControl.Zoom > 110 Then
    ReportControl.ZoomLimit = 105&
  ElseIf ReportControl.Zoom < 90 Then
    ReportControl.ZoomLimit = 95&
  Else
    ReportControl.ZoomLimit = 100&
  End If
End Sub

Public Sub SetPrinterMode(mode As BKMODE_TYPE)
  Call SetBkMode(Printer.hdc, mode)
End Sub

Public Function GetTextWidth(s As String) As Single
  Dim printToScreenRatio As Single, pfInfo As FontInfo
  Dim InError As Boolean, cfgWidth As Single
  Dim cTarget As REPORT_TARGET, pWidth As Single
    
  On Error GoTo GetTextWidth_err
  cTarget = ReportControl.rTarget
  'apf2008 changed - only scale printing on printer
  If IsPrinterAvail(False) And ((ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT)) Then
    ReportControl.rTarget = RPT_PRINTER
    Call SetFont
    GetTextWidth = Printer.ScaleX(Printer.TextWidth(s), Printer.ScaleMode, vbTwips)
  Else
    ReportControl.rTarget = RPT_CONFIG
    Call SetFont
    GetTextWidth = ReportControl.Preview.ScaleX(ReportControl.Preview.TextWidth(s), ReportControl.Preview.ScaleMode, vbTwips)
  
    'apf2008 get scale ratio from printer to screen font
    If Not InCollection(PrinterFonts, FontKey(ReportControl.Preview.Font.Name, ReportControl.fStyle.Size, ReportControl.Preview.Font.bold, ReportControl.Preview.Font.Italic, ReportControl.Preview.Font.UnderLine, ReportControl.Preview.Font.StrikeThrough)) Then
      'Set printer font to gather size information
      ReportControl.rTarget = RPT_PRINTER
      Call SetFont
    End If
    Set pfInfo = PrinterFonts.Item(FontKey(ReportControl.Preview.Font.Name, ReportControl.fStyle.Size, ReportControl.Preview.Font.bold, ReportControl.Preview.Font.Italic, ReportControl.Preview.Font.UnderLine, ReportControl.Preview.Font.StrikeThrough))
    printToScreenRatio = (pfInfo.Size / ReportControl.Preview.FontSize)
    GetTextWidth = printToScreenRatio * GetTextWidth
  End If
GetTextWidth_end:
  ReportControl.rTarget = cTarget
  Exit Function
  
GetTextWidth_err:
  InError = True
  Resume GetTextWidth_end
End Function

Function GetTextHeight(s As String) As Single
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    GetTextHeight = Printer.TextHeight(s)
  ElseIf IsExportHTML Then
    GetTextHeight = ReportControl.fStyle.FontHeight
  Else
    GetTextHeight = ReportControl.Preview.ScaleY(ReportControl.Preview.TextHeight(s), vbUser, vbTwips)
  End If
End Function

Public Function NoGraphics() As Boolean
  NoGraphics = (ReportControl.rTarget = RPT_PREPARE) Or (ReportControl.rTarget = RPT_CONFIG) Or (ReportControl.rTarget > RPT_EXPORT)
End Function

Public Sub TrimToWidth(s As String, TrimX As Single)
  Dim tw As Single, tw3 As Single
  
  If TrimX > 0 Then
    If Len(s) > 3 Then
      tw = GetTextWidth(s)
      tw3 = GetTextWidth("...")
      If (tw > (tw3 * 2)) And (tw > TrimX) Then
        Do While (tw > TrimX) And (Len(s) > 1)
          s = Left$(s, Len(s) - 1)
          tw = GetTextWidth(s) + tw3
        Loop
        s = s & "..."
      End If
    End If
    TrimX = 0!
  End If
End Sub

Public Function CenterToWidth(s As String, CenterX As Single) As Single
  If CenterX > 0 Then
    CenterToWidth = (CenterX - GetTextWidth(s)) / 2
    If CenterToWidth < 0 Then CenterToWidth = 0!
    CenterX = 0!
  End If
End Function


Public Function toupperbyte(ByVal ch As Byte) As Byte
  If (ch >= 97) And (ch <= 122) Then
    ch = ch - 97 + 65 ' - a + A
  End If
  toupperbyte = ch
End Function
