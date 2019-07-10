Attribute VB_Name = "Setup"
Option Explicit
Public Sub Main()
  ReDim PDF_DRIVER_NAMES(1 To 3) As String
  
  PDF_DRIVER_NAMES(1) = PDF_DRIVER_ONE_SOURCE
  PDF_DRIVER_NAMES(2) = PDF_DRIVER_SAGE
  PDF_DRIVER_NAMES(3) = PDF_DRIVER_ABACUS
  
End Sub
Public Sub InitReportEx(rptName As String, ByVal rpttarget As PRIVATE_REPORT_TARGET, ByVal orient As REPORT_ORIENTATION, ByVal bThrowPage As Boolean, ByVal OverrideHeader As Boolean, ByVal OverrideFooter As Boolean)
  Dim i As Integer
  
  Call xSet("InitreportEx")
  If Not ((rpttarget = RPT_PREPARE) Or (rpttarget = RPT_PRINTER)) Then
    Call Err.Raise(ERR_INVALID_TARGET, "InitReportEx", "The report target choosen is invalid, please choose PREPARE or PRINT")
  End If
  ReportControl.CurReport = ReportControl.CurReport + 1
  rpt(ReportControl.CurReport).ReportName = rptName
  If ReportControl.CurReport = 1 Then
    If rpttarget = RPT_PRINTER Then
      If PDFDriverInstall Then
        g_cdi.FileNameOptionsEx = 0 ' NoPrompt + UseFileName
        'g_cdi.DefaultFileName = exportName
      End If
      If Not IsPrinterAvail(True) Then Err.Raise ERR_NOPRINTER, "InitReportEx", "No printer defined or current printer invalid"
    End If
    Call SetCursor
    Call SetOrient(orient, rpttarget)
    Call resetdefaults
    ReportControl.ThrowPageOnSubReport = bThrowPage
    ReportControl.Name = rptName
    ReportControl.ReportDateTime = Now
    ReportControl.rTarget = rpttarget
    rpt(ReportControl.CurReport).OverrideHeader = OverrideHeader
    rpt(ReportControl.CurReport).OverrideFooter = OverrideFooter
    
    ReportControl.Preview.FontTransparent = True
    ReportControl.InReport = True
    Call InitNewPage
  Else
    If ReportControl.CurReport > MAX_REPORT_LEVELS Then Call Err.Raise(ERR_REPORT_STACK, "InitReportEx", "Report nesting to great. Max report Levels = " & CStr(MAX_REPORT_LEVELS))
    ReportControl.ReportHeaderPrinted = False
    ReportControl.ReportFooterPrinted = False

    rpt(ReportControl.CurReport).OverrideHeader = rpt(ReportControl.CurReport - 1).OverrideHeader
    rpt(ReportControl.CurReport).OverrideFooter = rpt(ReportControl.CurReport - 1).OverrideFooter
    
    'Call setrptheader(rpt(iReport - 1).RHeader)
    'Call SetRptFooter(rpt(iReport - 1).RFooter)
    
    ' page headers/footers
    If rpt(ReportControl.CurReport).OverrideHeader Then
      Call SetHeaderFooter(rpt(ReportControl.CurReport - 1).PHeader, PAGE_HEADER, False, True)
    Else
      Call SetHeaderFooter("", PAGE_HEADER, False, True)
    End If
    If rpt(ReportControl.CurReport).OverrideFooter Then
      Call SetHeaderFooter(rpt(ReportControl.CurReport - 1).PFooter, PAGE_FOOTER, False, True)
    Else
      Call SetHeaderFooter("", PAGE_FOOTER, False, True)
    End If
  End If
  Call SetHeaderFooter("", RECALC_ONLY)
  
  Call xReturn("InitReportEx")
End Sub

Public Sub resetdefaults()
  Rem track current style for report
  Call xSet("ResetDefaults")
  ReportControl.InReport = False
  ReportControl.PreviewOK = False
  ReportControl.IgnoreExportCR = False

  Set ReportControl.Preview = ReportControl.PreviewForm.picPaper
  Set ReportControl.PreviewTest = frmConfig.picPaper

  Call SetPageHeightWidth
  Call ResetFonts
  Call ClearPages
  ReportControl.Name = ""
  ReportControl.ThrowPageOnSubReport = False
  ReportControl.Preview.Cls
  ReportControl.NoOutput = False
  ReportControl.AbortReport = False
  ReportControl.CurPage = 0
  ReportControl.CurPageValid = False
  ReportControl.Pages_N = 0
  ReportControl.SuppressNewPageCalc = False
  ReportControl.PageHeaderPrinted = False
  ReportControl.PageTextPrinted = False
  ReportControl.PageFooterPrinted = False
  ReportControl.ReportHeaderPrinted = False
  ReportControl.ReportFooterPrinted = False
  ReportControl.Zoom = 100&
  ReportControl.PrinterNewPage = False
  ReportControl.OnNewPage = False
  
  ReportControl.FirstX = 0
  ReportControl.LeftMargin = 1
  ReportControl.RightMargin = 0
  
  ReportControl.TrimX = 0!
  ReportControl.CenterX = 0!
  ReportControl.BeginSectionX = -1!
  ReportControl.BeginSectionY = -1!
  ReportControl.OutputSection = False
  ReportControl.Section = ""
  ReportControl.FixedWidth = 0
  ReportControl.FW_PadLeft = True
  ReportControl.DelimitOut = False
  ReportControl.NoRecord = False
  ReportControl.PrintDlgOpt = PAGES_ALL
  ReportControl.PageTo = -1
  ReportControl.PageFrom = -1
  ReportControl.PageCopies = 1
  ReportControl.StartTime = GetTicks()
  ReportControl.EndTime = 0
  LINEWIDTH = LINEWIDTH_CONST
  Call ResetColors
  Call SetZoomLimit
  Call SetHeaderFooter("", PAGE_HEADER, False, True)
  Call SetHeaderFooter("", PAGE_FOOTER, False, True)
  Call SetHeaderFooter("", REPORT_FOOTER, False, True)
  Call SetHeaderFooter("", REPORT_HEADER, False, True)
  With ReportControl.HTML
    .OpenDiv = False
    .CloseDiv = False
    .CurrentX = 0
    .CurrentY = 0
    .HTMLFontSet = False
    .Position = False
    .XSetHTML = False
    .YSetHTML = False
    .TopString = "0mm"
    .BottomString = "0mm"
    .LeftString = "9mm"
    .RightString = "9mm"
    .ReportPages = 0
  End With
  NAV_HEADER = 0 '20 * Screen.TwipsPerPixelY
  Call xReturn("ResetDefaults")
End Sub
    
Public Sub EndReportEx(ByVal bForce As Boolean, ByVal bEndPreview As Boolean)
  Dim rlimit As Long, i As Long
  Dim endDocced As Boolean
  
  Call xSet("EndReportEx")
  
  If (ReportControl.rTarget = RPT_PRINTER) Then
    Call A4CheckPrinterSize
  End If
  
  
  i = ReportControl.CurReport
  rlimit = ReportControl.CurReport
  If bForce Then rlimit = 1
  Do While (i >= rlimit)
    Call EndCurrentReport
    ReportControl.CurReport = ReportControl.CurReport - 1
    i = i - 1
  Loop
  If ReportControl.CurReport < 1 Then
    If (ReportControl.rTarget = RPT_PRINTER) And ReportControl.InReport Then
      If ReportControl.AbortReport Or (ReportControl.Pages_N = 0) Then
        Printer.KillDoc
      Else
        Printer.EndDoc
        If (Not g_cdi Is Nothing) Then
          Call PDFDriverUninstall
        End If
        m_LastReportControl = ReportControl
      End If
    Else
      If bEndPreview Then
        ReportControl.PreviewOK = False
      ElseIf (ReportControl.rTarget = RPT_PREPARE) And ReportControl.InReport Then
        ReportControl.PreviewOK = True
      End If
    End If
    ReportControl.EndTime = GetTicks
    ReportControl.InReport = False
  End If
  ReportControl.AbortReport = False
  Call xReturn("EndReportEx")
End Sub

Private Sub EndCurrentReport()
  Dim btmp As Boolean
  
  If Not ReportControl.InReport Then Exit Sub
  Call xSet("EndCurrentReport")
  
  ' Print ReportFooter - make room for it first
  If Not ReportControl.AbortReport Then
    If ReportControl.rTarget = RPT_PRINTER Then
      If ReportControl.PageHeight <= (Printer.CurrentY + rpt(ReportControl.CurReport).RFooterH + rpt(ReportControl.CurReport).PFooterH) Then Call InitNewPage
    ElseIf ReportControl.rTarget = RPT_PREPARE Then
      If ReportControl.PageHeight <= (ReportControl.Preview.CurrentY + rpt(ReportControl.CurReport).RFooterH + rpt(ReportControl.CurReport).PFooterH) Then Call InitNewPage
    End If
    If ReportControl.CurPageValid Or (Len(rpt(ReportControl.CurReport).RFooter) > 0) Then
      Call PreOut
      ReportControl.SuppressNewPageCalc = True
      Call bOut("{Arial=10,N}" & rpt(ReportControl.CurReport).RFooter)
      ReportControl.SuppressNewPageCalc = False
    End If
    ' reset margins
    ReportControl.LeftMargin = 1
    ReportControl.RightMargin = 0
  End If
  If ReportControl.CurReport = 1 Then
    Call ClearCursor
    btmp = ReportControl.CurPageValid
    Call InitNewPage
    If Not btmp Then ReportControl.Pages_N = ReportControl.Pages_N - 1
  ElseIf ReportControl.CurPageValid And ReportControl.ThrowPageOnSubReport Then
    Call InitNewPage
  End If
  Call xReturn("EndCurrentReport")
End Sub

Private Sub SetPageHeightWidth()
  Dim sngHeight As Single, sngWidth As Single
  On Error GoTo SetPageHeight_err
  ' Page Height in twips
  
  If (A4Force) Then
    sngWidth = 11905.511811024 - (2 * A4NonPrintableMArginTwips())
    sngHeight = 16837.79527559 - (2 * A4NonPrintableMArginTwips())
  Else
    sngHeight = Printer.ScaleHeight - NOPRINT_AREA_BL
    sngWidth = Printer.ScaleWidth - NOPRINT_AREA_BL
  End If
  'ReportControl.PageHeight = Printer.ScaleHeight - NOPRINT_AREA_BL   'EAW landscape.. changed to 2, still wrong name
  'ReportControl.PageWidth = Printer.ScaleWidth - NOPRINT_AREA_BL

  If ReportControl.Orientation = PORTRAIT Then
    If (sngHeight > sngWidth) Then
      ReportControl.PageHeight = sngHeight
      ReportControl.PageWidth = sngWidth
    Else
      ReportControl.PageHeight = sngWidth
      ReportControl.PageWidth = sngHeight
    End If
  Else
    If (sngHeight > sngWidth) Then
      ReportControl.PageHeight = sngWidth
      ReportControl.PageWidth = sngHeight
    Else
      ReportControl.PageHeight = sngHeight
      ReportControl.PageWidth = sngWidth
    End If
  End If

SetPageHeight_end:
  'apf2008 set preview control to operate in twips
  With ReportControl.Preview
    .ScaleTop = 0
    .ScaleLeft = 0
    .ScaleHeight = ReportControl.PageHeight
    .ScaleWidth = ReportControl.PageWidth
    .ScaleMode = vbTwips
    
  End With
  With ReportControl.PreviewTest
    .ScaleTop = 0
    .ScaleLeft = 0
    .ScaleHeight = ReportControl.PageHeight
    .ScaleWidth = ReportControl.PageWidth
    .ScaleMode = vbTwips
  End With
  Exit Sub
  
SetPageHeight_err:
  If ReportControl.Orientation = PORTRAIT Then
    ReportControl.PageHeight = ((297! * 1440) / 25.4) - NOPRINT_AREA_BL
    ReportControl.PageWidth = ((210! * 1440) / 25.4) - NOPRINT_AREA_BL
  Else
    ReportControl.PageHeight = ((210! * 1440) / 25.4) - NOPRINT_AREA_BL
    ReportControl.PageWidth = ((297! * 1440) / 25.4) - NOPRINT_AREA_BL
  End If
  Resume SetPageHeight_end
End Sub
