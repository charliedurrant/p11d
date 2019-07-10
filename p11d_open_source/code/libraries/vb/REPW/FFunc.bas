Attribute VB_Name = "FontFunctions"
Option Explicit

Public Declare Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, ByVal lpxBuffer As Long) As Long
Public Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, ByVal lpBuffer As Long) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long

Private m_STACKTOP As Long
Private m_STACKMAX As Long
Private m_StyleStack() As FontStyle

Private Sub SetFontEx(CFont As StdFont, ByVal fh As Single, ByVal FontTypeMode As FONT_TYPE)
  Dim fi As FontInfo, pfi As FontInfo
  Dim fs As Single
  
  'apf2008 do not set font unless necessary
  With ReportControl.fStyle
    If Abs(CFont.Size - fh) > 0.2 Then
      If (fh < 8) Then
        CFont.Size = fh
      End If
      'may need to set font name check
      CFont.Name = .Name
      CFont.Size = fh
    End If
    If (CFont.Name <> .Name) Then CFont.Name = .Name
    If (CFont.bold <> .bold) Then CFont.bold = .bold
    If (CFont.Italic <> .Italic) Then CFont.Italic = .Italic
    If (CFont.UnderLine <> .UnderLine) Then CFont.UnderLine = .UnderLine
    If (CFont.StrikeThrough <> .StrikeThrough) Then CFont.StrikeThrough = .StrikeThrough
    If FontTypeMode = PRINTER_FONT Then
      If Not InCollection(PrinterFonts, FontKey(.Name, fh, .bold, .Italic, .UnderLine, .StrikeThrough)) Then
        ' add in mapped font
        Set fi = New FontInfo
        fi.Name = .Name
        fi.SizeRequested = fh
        fi.Size = .Size
        fi.bold = .bold
        fi.Italic = .Italic
        fi.UnderLine = .UnderLine
        fi.StrikeThrough = .StrikeThrough
        Call PrinterFonts.Add(fi, fi.Key)
      End If
    End If
  End With
End Sub


Public Sub SetFont()
  Dim InError As Boolean
  Dim fh As Single
  Call xSet("SetFont")
  
  On Error GoTo SetFont_err
SetFont_retry:
  If ReportControl.rTarget > RPT_EXPORT And Not IsExportHTML Then GoTo SetFont_end
  If ReportControl.fStyle.FontType = INVALID_FONT_TYPE Then
    Call ResetFonts
  End If
    
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    If ReportControl.fStyle.FontType <> PRINTER_FONT Then
      Call SetFontEx(Printer.Font, ReportControl.fStyle.Size, PRINTER_FONT)
      ReportControl.fStyle.FontType = PRINTER_FONT
    End If
  ElseIf ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE Then
    If ReportControl.fStyle.FontType <> PREVIEW_SCREEN_FONT Then
      Call SetFontEx(ReportControl.Preview.Font, (ReportControl.fStyle.Size * ReportControl.Zoom) / ReportControl.ZoomLimit, PREVIEW_SCREEN_FONT)
      ReportControl.fStyle.FontType = PREVIEW_SCREEN_FONT
    End If
  ElseIf (ReportControl.rTarget = RPT_CONFIG) Or IsExportHTML Then
    If ReportControl.fStyle.FontType <> SCREEN_FONT Then
      Call SetFontEx(ReportControl.Preview.Font, ReportControl.fStyle.Size, SCREEN_FONT)
      ReportControl.fStyle.FontType = SCREEN_FONT
    End If
  End If
  
SetFont_end:
  Call xReturn("SetFont")
  Exit Sub
  
SetFont_err:
  If InError Then Resume SetFont_end
  InError = True
  ReportControl.fStyle.FontType = INVALID_FONT_TYPE
  Resume SetFont_retry
End Sub

Public Sub PushFontStyle()
  Call xSet("PushFontStyle")
  m_STACKTOP = m_STACKTOP + 1
  If m_STACKTOP > m_STACKMAX Then
    m_STACKMAX = m_STACKMAX + 1
    ReDim Preserve m_StyleStack(1 To m_STACKMAX) As FontStyle
  End If
  Call CopyFontStyle(m_StyleStack(m_STACKTOP), ReportControl.fStyle)
  Call xReturn("PushFontStyle")
End Sub

Public Sub PopFontStyle(Optional ByVal bSetFont As Boolean = False)
  Call xSet("PopFontStyle")
  
  If m_STACKTOP > 0 Then
    Call CopyFontStyle(ReportControl.fStyle, m_StyleStack(m_STACKTOP))
    ReportControl.fStyle.FontType = VALID_FONT_TYPE
    Call SetFont
    m_STACKTOP = m_STACKTOP - 1
  Else
    Call Err.Raise(ERR_POPFONTSTACK, "PopFontStyle", "Pop without Push!")
  End If
  Call xReturn("PopFontStyle")
End Sub

Public Sub ResetFonts()
  Call CopyFontStyle(ReportControl.fStyle, DefaultFontStyle)
  DefaultFontStyle.FontType = VALID_FONT_TYPE
End Sub

Public Sub ProcessFontFormat(FontName As String, ByVal paramcount As Long, ByVal startparam As Long, params() As String)
  Dim fmt() As Byte, i As Long, tlen As Long
  Dim fmtstr As String
  
  Call xSet("ProcessFontFormat")
  If ReportControl.rTarget > RPT_EXPORT And Not IsExportHTML Then GoTo ProcessFontFormat_end
  ReportControl.fStyle.Name = FontName
  ReportControl.fStyle.FontType = VALID_FONT_TYPE
  If paramcount >= 1 Then
    ReportControl.fStyle.Size = CSng(params(startparam))
    ReportControl.fStyle.FontHeight = FontHeights(ReportControl.fStyle.Size)
  End If
  If paramcount = 2 Then
    fmtstr = UCase$(params(startparam + 1))
    If InStr(1, fmtstr, "N", vbBinaryCompare) > 0 Then
      ReportControl.fStyle.bold = False
      ReportControl.fStyle.Italic = False
      ReportControl.fStyle.UnderLine = False
      ReportControl.fStyle.StrikeThrough = False
      ReportControl.fStyle.Align = ALIGN_LEFT
      If Len(fmtstr) = 1 Then GoTo ProcessFontFormat_setfont
    End If
    fmt = fmtstr
    tlen = LenB(fmtstr) - 1
    For i = 0 To tlen Step 2
      Select Case fmt(i)
        Case vbKeyB ' B
            ReportControl.fStyle.bold = True
        Case vbKeyU ' U
            ReportControl.fStyle.UnderLine = True
        Case vbKeyI  'I
            ReportControl.fStyle.Italic = True
        Case vbKeyS  'S
            ReportControl.fStyle.StrikeThrough = True
        Case vbKeyL  'L
            ReportControl.fStyle.Align = ALIGN_LEFT
        Case vbKeyC  'C
            ReportControl.fStyle.Align = ALIGN_CENTER
        Case vbKeyR  'R
            ReportControl.fStyle.Align = ALIGN_RIGHT
        Case vbKeyN
        Case Else
            Err.Raise ERR_INVALIDFONT, "ProcessFontFormat", "Invalid font format flags: " & UCase$(params(startparam + 1))
      End Select
    Next i
  End If
ProcessFontFormat_setfont:
  Call SetFont
  
ProcessFontFormat_end:
  Call xReturn("ProcessFontFormat")
End Sub

Public Sub CopyFontStyle(FDest As FontStyle, FSource As FontStyle)
  FDest.Name = FSource.Name
  FDest.Size = FSource.Size
  FDest.bold = FSource.bold
  FDest.Italic = FSource.Italic
  FDest.UnderLine = FSource.UnderLine
  FDest.StrikeThrough = FSource.StrikeThrough
  FDest.FontHeight = FSource.FontHeight
  FDest.Align = FSource.Align
  FDest.FontType = FSource.FontType
End Sub

Public Function FontKey(ByVal Name As String, ByVal Size As Single, ByVal bold As Boolean, ByVal Italic As Boolean, ByVal UnderLine As Boolean, ByVal StrikeThrough As Boolean) As String
  FontKey = Name & "::" & Size & "::" & bold & "::" & Italic & "::" & UnderLine & "::" & StrikeThrough
End Function

