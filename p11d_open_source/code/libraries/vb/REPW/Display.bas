Attribute VB_Name = "Display"
Option Explicit

Private Const END_PAGE_INC As Single = 199 ' 1 twip short of 10 Pt Font

Public Function DisplayText(Text As String) As Boolean
  Call xSet("DisplayText")
  
  On Error GoTo DisplayText_err
  If Len(Text) > 0 Then
    ReportControl.PageTextPrinted = True
    Call SetFont
    Call SetColor
    Call PrintText(Text)
  End If
DisplayText_end:
  Call xReturn("DisplayText")
  Exit Function
 
DisplayText_err:
  If Err.Number = 482 Then
    Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "DisplayText", "Display text", "Unable to complete display operation" & vbCrLf & "Aborting current report - " & ReportControl.Name)
    ReportControl.NoOutput = True
    ReportControl.AbortReport = True
  Else
    Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "DisplayText", "Display text", "Unable to complete display operation")
  End If
  Resume DisplayText_end
  Resume
End Function

Public Function CheckForNewPage(ByVal fh As Single, ByVal CheckOnReturn As Boolean, ByVal CheckOnly As Boolean) As Boolean
  Static inCheckForNewPage As Boolean
  Dim cy As Single, SaveBeginSection As Single
  'check for new page END_PAGE
    
  inCheckForNewPage = False
  CheckForNewPage = False
  If ReportControl.SuppressNewPageCalc Or IsExportHTML Then GoTo CheckForNewPage_end
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREPARE) Then
    If inCheckForNewPage Then Call Err.Raise(ERR_PAGERECURSIVE, "CheckForNewPage", "Attempted to print another new page while processing a new page, reports cannot be recursive")
    inCheckForNewPage = True

    If ReportControl.rTarget = RPT_PRINTER Then
      cy = Printer.CurrentY
    Else
      cy = ReportControl.Preview.CurrentY
    End If
    If CheckOnReturn Then fh = fh + END_PAGE_INC
    If ReportControl.PageHeight < (cy + fh + rpt(ReportControl.CurReport).PFooterH) Then
      If Not CheckOnly Then
        SaveBeginSection = ReportControl.BeginSectionY
        ReportControl.BeginSectionY = -1
        Call PushFontStyle
        Call PushColor
        Call InitNewPage 'suppress pagecalc done in Init
        Call PopColor
        Call PopFontStyle
        ReportControl.SuppressNewPageCalc = True
        Call PreOut
        ReportControl.SuppressNewPageCalc = False
        ReportControl.OnNewPage = True
        If ReportControl.rTarget = RPT_PRINTER Then
          cy = Printer.CurrentY
        Else
          cy = ReportControl.Preview.CurrentY
        End If
        If SaveBeginSection <> -1 Then
         'ReportControl.OutputSection = True
          ReportControl.BeginSectionY = cy
        End If
        Call SetStackX(0)
        Call SetStackY(cy)
      End If
      CheckForNewPage = True
    End If
    inCheckForNewPage = False
  End If
CheckForNewPage_end:
End Function

Private Sub PrintText(s As String)
  Dim DoPrint As Boolean, Returns As Long
  Dim tmp0 As String, tmp1 As String
  Dim p0 As Long, p1 As Long
  Dim pos As Integer, lastpos As Integer
  Dim fh As Single, cury As Single, curx As Single, OffsetX As Single

  ReportControl.CurPageValid = True
  fh = ReportControl.fStyle.FontHeight
  'check for new page END_PAGE
  Call CheckForNewPage(fh, False, False)
  pos = 1: lastpos = 1
  Do While (pos <> 0)
    If ReportControl.OnNewPage And (ReportControl.BeginSectionY <> -1) Then Exit Sub
    pos = InStr(pos, s, vbCrLf)
    If pos = 0 Then
      tmp0 = Mid$(s, lastpos)
    Else
      tmp0 = Mid$(s, lastpos, (pos - lastpos))
      pos = pos + 2
    End If
    lastpos = pos

    If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
      DoPrint = True
      cury = Printer.CurrentY
      curx = Printer.CurrentX
    ElseIf (ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE) Or (ReportControl.rTarget = RPT_CONFIG) Or (ReportControl.rTarget = RPT_PREPARE) Then
      DoPrint = False
      cury = ReportControl.Preview.CurrentY
      curx = ReportControl.Preview.CurrentX
      If ReportControl.rTarget = RPT_PREPARE Then
        If (ReportControl.TrimX > 0!) Or (ReportControl.CenterX > 0!) Then
          ReportControl.DelimitOut = True
          ReportControl.TrimX = 0!
          ReportControl.CenterX = 0!
        End If
        GoTo next_line_process
      End If
    ElseIf ReportControl.rTarget > RPT_EXPORT Then
      If ReportControl.fStyle.Align = ALIGN_RIGHT Then
        ReportControl.HTML.CurrentX = (ReportControl.HTML.CurrentX - GetTextWidth(tmp0) - 10)
      End If
      Call ExportOut(OUT_TEXT, tmp0)
      GoTo next_line_process
    End If
    
    If Len(tmp0) > 0 Then
      If ReportControl.fStyle.Align = ALIGN_RIGHT Then
        curx = (curx - GetTextWidth(tmp0) - 10)
      End If
      If ReportControl.FColor.ColorFill Then
        If DoPrint Then
          Call TextBoxP(tmp0, fh, curx, cury)
        Else
          Call TextBoxS(tmp0, fh, curx, cury)
        End If
      End If
      Call CheckXCoord(curx, tmp0)
      Call TrimToWidth(tmp0, ReportControl.TrimX)
      OffsetX = CenterToWidth(tmp0, ReportControl.CenterX)
      
      If DoPrint Then
        Printer.CurrentX = curx + OffsetX
        Printer.CurrentY = cury
        If ReportControl.BeginSectionY = -1 Then Printer.Print tmp0;
      Else
        ReportControl.Preview.CurrentX = curx + OffsetX
        ReportControl.Preview.CurrentY = cury
        ' ReportControl.BeginSectionY cannot be -1 on preview
        ReportControl.Preview.Print tmp0;
      End If
    End If
    
next_line_process:
    If pos <> 0 Then
      ReportControl.TrimX = 0!
      ReportControl.CenterX = 0!
      If DoPrint Then
        Printer.CurrentX = ReportControl.LeftMargin + ReportControl.FirstX
        If Not CheckForNewPage(fh, True, False) Then Printer.CurrentY = cury + fh
      Else
        If ReportControl.rTarget > RPT_EXPORT Then
          If Not ReportControl.IgnoreExportCR Then
            Call ExportOut(OUT_CR, "")
          Else
            If NotInStr(s, vbCrLf, pos) > 0 Then Call ExportOut(OUT_TEXT, " ")
          End If
        Else
          ReportControl.Preview.CurrentX = ReportControl.LeftMargin + ReportControl.FirstX
          If Not CheckForNewPage(fh, True, False) Then ReportControl.Preview.CurrentY = cury + fh
        End If
      End If
    Else
      If DoPrint Then
        Printer.CurrentY = cury
      Else
        ReportControl.Preview.CurrentY = cury
      End If
    End If
next_line:
  Loop
End Sub

' negative line percent = right aligned
Public Sub xBox(ByVal PercentWidth As Single, ByVal PercentHeight As Single, ByVal bFilled As Boolean)
  Dim x0 As Single, x1 As Single, y0 As Single, y1 As Single

  Call xSet("xBox")
  ReportControl.CurPageValid = True
  If NoGraphics() Then GoTo xBox_end
  Call SetFont
  Call SetColor

  Call PushCoord
    x0 = StackTopX
    y0 = StackTopY
    If PercentWidth >= 100 Then x0 = 1
    If PercentWidth <= -100 Then x0 = ReportControl.PageWidth
    x1 = GetRelativePercent(x0, PercentWidth, ReportControl.PageWidth)
    y1 = GetRelativePercent(y0, PercentHeight, ReportControl.PageHeight)
    If ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE Then
      Call LineBoxPrintS(x0, y0, x1, y1, bFilled)
    Else ' must be printer
      Call LineBoxPrintP(x0, y0, x1, y1, bFilled)
    End If
  Call PopCoord
xBox_end:
  Call xReturn("xBox")
End Sub

' negative line percent = right aligned
Public Sub xLine(ByVal PercentWidth As Single, ByVal bDouble As Boolean)
  Dim x0 As Single, x1 As Single, y0 As Single

  Call xSet("xLine")
  ReportControl.CurPageValid = True
  If IsExportHTML Then
    Call SetHTMLLine(PercentWidth, bDouble, True)
  End If
  If NoGraphics() Then GoTo xLine_end
  Call SetFont
  Call SetColor

  Call PushCoord
    x0 = StackTopX
    y0 = StackTopY
    If PercentWidth >= 100 Then x0 = 1
    If PercentWidth <= -100 Then x0 = ReportControl.PageWidth
    x1 = GetRelativePercent(x0, PercentWidth, ReportControl.PageWidth)
    If ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE Then
      Call LineBoxPrintS(x0, y0, x1, y0, True)
      If bDouble Then Call LineBoxPrintS(x0, y0 + LINESPACE, x1, y0 + LINESPACE, True)
    Else ' must be printer
      Call LineBoxPrintP(x0, y0, x1, y0, True)
      If bDouble Then Call LineBoxPrintP(x0, y0 + LINESPACE, x1, y0 + LINESPACE, True)
    End If
  Call PopCoord
  
xLine_end:
  Call xReturn("xLine")
End Sub

' negative line percent = right aligned
Public Sub xLineAbs(ByVal Width As Single, ByVal bDouble As Boolean)
  Dim x0 As Single, x1 As Single, y0 As Single

  Call xSet("xLineAbs")
  ReportControl.CurPageValid = True
  If NoGraphics() Then GoTo xLineAbs_end
  Call SetFont
  Call SetColor

  Call PushCoord
    x0 = StackTopX
    y0 = StackTopY
    x1 = GetRelative(x0, Width, ReportControl.PageWidth)
    If ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE Then
      Call LineBoxPrintS(x0, y0, x1, y0, True)
      If bDouble Then Call LineBoxPrintS(x0, y0 + LINESPACE, x1, y0 + LINESPACE, True)
    Else ' must be printer
      Call LineBoxPrintP(x0, y0, x1, y0, True)
      If bDouble Then Call LineBoxPrintP(x0, y0 + LINESPACE, x1, y0 + LINESPACE, True)
    End If
  Call PopCoord
  
xLineAbs_end:
  Call xReturn("xLineAbs")
End Sub

' negative line percent = line up
Public Sub yLine(ByVal PercentHeight As Single, ByVal bDouble As Boolean)
  Dim y0 As Single, y1 As Single, x0 As Single

  Call xSet("yLine")
  ReportControl.CurPageValid = True
  If NoGraphics() Then GoTo yLine_end
  Call SetFont
  Call SetColor

  Call PushCoord
    y0 = StackTopY
    x0 = StackTopX
    If PercentHeight >= 100 Then y0 = 1
    If PercentHeight <= -100 Then y0 = ReportControl.PageHeight
    y1 = GetRelativePercent(y0, PercentHeight, ReportControl.PageHeight)
    If ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE Then
      Call LineBoxPrintS(x0, y0, x0, y1, True)
      If bDouble Then Call LineBoxPrintS(x0 + LINESPACE, y0, x0 + LINESPACE, y1, True)
    Else ' must be printer
      Call LineBoxPrintP(x0, y0, x0, y1, True)
      If bDouble Then Call LineBoxPrintS(x0 + LINESPACE, y0, x0 + LINESPACE, y1, True)
    End If
  Call PopCoord
  
yLine_end:
  Call xReturn("yLine")
End Sub

Private Sub LineBoxPrintS(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal bFill As Boolean)
  If bFill Then
    ReportControl.Preview.Line (x0, y0)-(x1, y1), ReportControl.FColor.LineColor, BF
  Else
    ReportControl.Preview.Line (x0, y0)-(x1, y1), ReportControl.FColor.LineColor, B
  End If
End Sub

Private Sub LineBoxPrintP(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal bFill As Boolean)
  If ReportControl.BeginSectionY <> -1 Then Exit Sub
  If bFill Then
    Printer.Line (x0, y0)-(x1, y1), ReportControl.FColor.LineColor, BF
  Else
    Printer.Line (x0, y0)-(x1, y1), ReportControl.FColor.LineColor, B
  End If
End Sub

Private Sub TextBoxS(s As String, ByVal fh As Single, ByVal x0 As Single, ByVal y0 As Single)
  Dim x1 As Single, y1 As Single
  
  x1 = x0 + ReportControl.Preview.TextWidth(s)
  y1 = y0 + fh + 20
  Call LineBoxPrintS(x0, y0, x1, y1, True)
End Sub

Private Sub TextBoxP(s As String, ByVal fh As Single, ByVal x0 As Single, y0 As Single)
  Dim x1 As Single, y1 As Single
  
  x1 = x0 + Printer.TextWidth(s) - 10
  y1 = y0 + fh + 5
  y0 = y0 - 30
  If y0 < 0 Then y0 = 1
  x0 = x0 - 10
  If x0 < 0 Then x0 = 1
  Call LineBoxPrintP(x0, y0, x1, y1, True)
End Sub

Private Sub CheckXCoord(ByVal x0 As Single, PrintString As String)
  If (x0 < 0) Or (x0 > ReportControl.PageWidth) Then Call Err.Raise(ERR_NOPRINT, "CheckXCoord", "Attempt to print text """ & PrintString & """ outside the bounds of the current Page" & vbCr & "Cannot print at X=" & CStr(x0))
End Sub

Public Sub DisplayTextBox(s As String, ByVal xPercent As Single, ByVal yPercent As Single, ByVal bFill As Boolean, ByVal Xalign As ALIGNMENT_TYPE)
  Dim ForeColorOld As Long, FillColorOld As Long
  Dim XWidth As Single, TWidth As Single, x0 As Single
  Dim YHeight  As Single, THeight  As Single, y0 As Single
  
  Call PushCoord
  If Not NoGraphics() Then
    Call PushFontStyle
    Call xBox(xPercent, yPercent, bFill)
    
    ReportControl.FColor.ColorFill = False  ' turn off box on text
    ReportControl.fStyle.Align = ALIGN_LEFT ' no alignment on text
    
    XWidth = GetPercent(xPercent, ReportControl.PageWidth)
    TWidth = GetTextWidth(s)
    If Xalign = ALIGN_CENTER Then
      x0 = (XWidth - TWidth) / 2
    ElseIf Xalign = ALIGN_RIGHT Then
      x0 = XWidth - TWidth - 10
    End If
    If x0 < 0 Then x0 = 0
    
    YHeight = GetPercent(yPercent, ReportControl.PageHeight)
    THeight = GetTextHeight(s)
    y0 = (YHeight - THeight) / 2
    If y0 < 0 Then y0 = 0
    
    If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
      Printer.CurrentX = Printer.CurrentX + x0
      Printer.CurrentY = Printer.CurrentY + y0
    Else
      ReportControl.Preview.CurrentX = ReportControl.Preview.CurrentX + x0
      ReportControl.Preview.CurrentY = ReportControl.Preview.CurrentY + y0
    End If
    Call DisplayText(s)
    Call PopFontStyle(True)
  Else
    Call DisplayText(s)
  End If
  Call PopCoord
End Sub
   
Public Function PreOut() As Boolean

  On Error GoTo PreOut_Err
  Call xSet("PreOut")
  If ReportControl.NoRecord Then GoTo PreOut_End
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREPARE) Then
    If ReportControl.PrinterNewPage Then
      Printer.NewPage
      Printer.Print "";
      Printer.CurrentX = ReportControl.LeftMargin
      Printer.CurrentY = 1
      ReportControl.FColor.ColorSet = False
      ReportControl.fStyle.FontType = VALID_FONT_TYPE
      ReportControl.PrinterNewPage = False
    End If
    If Not ReportControl.ReportHeaderPrinted Then
      ReportControl.ReportHeaderPrinted = True
      If Len(rpt(ReportControl.CurReport).RHeader) > 0 Then Call bOut("{Arial=10,N}" & rpt(ReportControl.CurReport).RHeader)
    End If
    If Not ReportControl.PageHeaderPrinted Then
      ReportControl.PageHeaderPrinted = True
      ReportControl.PageTextPrinted = False
      If Len(rpt(ReportControl.CurReport).PHeader) > 0 Then Call bOut("{Arial=10,N}{STARTSKIPEXPORT}" & rpt(ReportControl.CurReport).PHeader & "{ENDSKIPEXPORT}")
    End If
  End If

PreOut_End:
  Call xReturn("PreOut")
  Exit Function

PreOut_Err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "PreOut", "Pre Output", "Unable to output report " & rpt(ReportControl.CurReport).ReportName & " error in Pre output stage")
  Resume PreOut_End
End Function


