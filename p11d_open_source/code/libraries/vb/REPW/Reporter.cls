VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ILibraryVersion

Public Enum REPORT_TARGET
  PREPARE_REPORT = 255
  PRINT_REPORT
End Enum

Public Enum REPORT_EXPORTS
  [_REPORT_EXPORTS_FIRST] = 4097
  EXPORT_CSV = [_REPORT_EXPORTS_FIRST]
  EXPORT_EXCEL
  EXPORT_WORD
  EXPORT_FIXEDWIDTH
  EXPORT_HTML_IE
  EXPORT_HTML_NETSCAPE
  [_REPORT_EXPORTS_LAST] = EXPORT_HTML_NETSCAPE
  EXPORT_HTML_INTEXP5 'update const L_LAST_EXPORT if add value
End Enum



Public Enum REPORT_CONSTANTS
  REPORT_VAR0 = 0
  REPORT_VAR1
  REPORT_VAR2
  REPORT_VAR3
  REPORT_VAR4
  REPORT_VAR5
  REPORT_VAR6
  REPORT_VAR7
  REPORT_VAR8
  REPORT_VAR9
  REPORT_CONSTANTS_N
End Enum

Public Enum REPORT_ORIENTATION
  PORTRAIT = vbPRORPortrait
  LANDSCAPE = vbPRORLandscape
End Enum

Public Property Get ReportHeader() As String
  Call xSet("ReportHeaderGet")
  ReportHeader = SetHeaderFooter("", REPORT_HEADER, True)
  Call xReturn("ReportHeaderGet")
End Property

Public Property Let ReportHeader(NewValue As String)
  Call xSet("ReportHeaderLet")
  Call SetHeaderFooter(NewValue, REPORT_HEADER)
  Call xReturn("ReportHeaderLet")
End Property

Public Property Get HTMLContactString() As String
  HTMLContactString = ReportControl.HTML.ContactString
End Property

Public Property Let HTMLContactString(NewValue As String)
  ReportControl.HTML.ContactString = NewValue
End Property

Public Property Get HTMLOrientationString() As String
  HTMLOrientationString = ReportControl.HTML.OrientationString
End Property

Public Property Let HTMLOrientationString(NewValue As String)
  ReportControl.HTML.OrientationString = NewValue
End Property

Public Property Get HTMLTopString() As String
  HTMLTopString = ReportControl.HTML.TopString
End Property

Public Property Let HTMLTopString(NewValue As String)
  ReportControl.HTML.TopString = NewValue
End Property

Public Property Get HTMLBottomString() As String
  HTMLBottomString = ReportControl.HTML.BottomString
End Property

Public Property Let HTMLBottomString(NewValue As String)
  ReportControl.HTML.BottomString = NewValue
End Property

Public Property Get HTMLLeftString() As String
  HTMLLeftString = ReportControl.HTML.LeftString
End Property

Public Property Let HTMLLeftString(NewValue As String)
  ReportControl.HTML.LeftString = NewValue
End Property

Public Property Get HTMLRightString() As String
  HTMLRightString = ReportControl.HTML.RightString
End Property

Public Property Let HTMLRightString(NewValue As String)
  ReportControl.HTML.RightString = NewValue
End Property

Public Property Get ReportFooter() As String
  Call xSet("ReportFooterGet")
  ReportFooter = SetHeaderFooter("", REPORT_FOOTER, True)
  Call xReturn("ReportFooterGet")
End Property

Public Property Let ReportFooter(NewValue As String)
  Call xSet("ReportFooterLet")
  Call SetHeaderFooter(NewValue, REPORT_FOOTER)
  Call xReturn("ReportFooterLet")
End Property

Public Property Get ExportHeader() As String
  Call xSet("ExportHeaderGet")
  ExportHeader = SetHeaderFooter("", EXPORT_HEADER, True)
  Call xReturn("ExportHeaderGet")
End Property

Public Property Let ExportHeader(NewValue As String)
  Call xSet("ExportHeaderLet")
  Call SetHeaderFooter(NewValue, EXPORT_HEADER)
  Call xReturn("ExportHeaderLet")
End Property

Public Property Get PageHeader() As String
  Call xSet("PageHeaderGet")
  PageHeader = SetHeaderFooter("", PAGE_HEADER, True)
  Call xReturn("PageHeaderGet")
End Property

Public Property Let PageHeader(NewValue As String)
  Call xSet("PageHeaderLet")
  Call SetHeaderFooter(NewValue, PAGE_HEADER)
  Call xReturn("PageHeaderLet")
End Property

Public Property Get PageFooter() As String
  Call xSet("PageFooterGet")
  PageFooter = SetHeaderFooter("", PAGE_FOOTER, True)
  Call xReturn("PageFooterGet")
End Property

Public Property Let PageFooter(NewValue As String)
  Call xSet("PageFooterLet")
  Call SetHeaderFooter(NewValue, PAGE_FOOTER)
  Call xReturn("PageFooterLet")
End Property

Public Property Let StartPageNumber(ByVal NewValue As Long)
  If Not ReportControl.InReport Then Err.Raise ERR_NO_INIT_REPORT, "StartPageNumber", "No report has been initialised."
  Pages(ReportControl.CurPage).PageNumber = NewValue
End Property

Public Property Let PrePageNumber(ByVal NewValue As String)
  If Not ReportControl.InReport Then Err.Raise ERR_NO_INIT_REPORT, "PrePageNumber", "No report has been initialised."
  Pages(ReportControl.CurPage).PrePageNumber = NewValue
End Property

Public Property Let PostPageNumber(ByVal NewValue As String)
  If Not ReportControl.InReport Then Err.Raise ERR_NO_INIT_REPORT, "PostPageNumber", "No report has been initialised."
  Pages(ReportControl.CurPage).PostPageNumber = NewValue
End Property

Public Property Get LastReportTime() As Long
  If Not ReportControl.InReport Then LastReportTime = ReportControl.EndTime = ReportControl.StartTime
End Property

'* Initialise a report or a sub report
'* Notes:  You are only allowed to nest reports up to 10 deep
'* Throwpage and inherit are ignored for the top level report
'* Orientation cannot be changed mid report
'*
'* ReportName               Name of report
'* Destination              Select printer or screen (only used on top level reports)
'* [Orientation]            Orientation of paper - cannot be changed within a report at the moment
'* [ThrowPage]              Throw a new page at the start of this Report if it is a
'*                          Sub report
'* [OverrideHeader]         Forces the current Page Header onto Sub reports
'* [OverrideFooter]         Forces the current Page footer onto Sub reports
'*
'* return value:
'* true if report is initialised correctly
Public Function InitReport(ByVal ReportName As String, ByVal Destination As REPORT_TARGET, Optional ByVal Orientation As REPORT_ORIENTATION = PORTRAIT, Optional ByVal ThrowPageOnSubReport As Boolean = True, Optional ByVal OverrideHeader As Boolean = False, Optional ByVal OverrideFooter As Boolean = False) As Boolean
  On Error GoTo InitReport_err
  Call xSet("InitReport")
  
  Call InitReportEx(ReportName, Destination, Orientation, ThrowPageOnSubReport, OverrideHeader, OverrideFooter)
  InitReport = ReportControl.InReport
  
InitReport_end:
  Call xReturn("InitReport")
  Exit Function
  
InitReport_err:
  InitReport = False
  Call ErrorMessage(ERR_ERROR, Err, "InitReport", "ERR_INITREPORT", "Error initialising report " & ReportName)
  Resume InitReport_end
End Function

'* End the most recent report initialised.
'*
'* return value
'* true
Public Function EndReport(Optional ByVal Force As Boolean) As Boolean
  On Error GoTo EndReport_err
  
  Call xSet("EndReport")
  Call EndReportEx(Force, False)
  EndReport = True

EndReport_end:
  Call xReturn("EndReport")
  Exit Function
  
EndReport_err:
  EndReport = False
  Call ErrorMessage(ERR_ERROR, Err, "EndReport", "ERR_ENDREPORT", "Error terminating report.")
  Resume EndReport_end
  Resume
End Function
Public Sub SetPrevExportFooter(ByVal s As String)
  Pages(ReportControl.CurPage - 1).ExportOnlyFooter = s
End Sub

'* Print text, boxes, lines etc to the printer
'* Uses the TCS control codes
'*
'* FormattedText                   The text and control codes
'*
'* return value:
'* True if NewPage has been thrown
Public Function Out(FormattedText As String) As Boolean
  On Error GoTo Out_err
  
  Call xSet("Out")
  If Not ReportControl.InReport Then Err.Raise ERR_NO_INIT_REPORT, "Out", "No report has been initialised"
  If ReportControl.AbortReport Then Exit Function
  Call PreOut
  Out = bOut(FormattedText)
  
Out_end:
  Call xReturn("Out")
  Exit Function
  
Out_err:
  Resume Out_end
End Function

'* Returns True if PageBreak would be thrown if nReturns were output in current font
Public Function TestPageBreak(ByVal nReturns As Long) As Boolean
 
  Call xSet("TestPageBreak")
  If Not ReportControl.InReport Then Err.Raise ERR_NO_INIT_REPORT, "TestPageBreak", "No report has been initialised."
  If ReportControl.AbortReport Then Exit Function
  TestPageBreak = CheckForNewPage(ReportControl.fStyle.FontHeight * nReturns, True, True)

TestPageBreak_end:
  Call xReturn("TestPageBreak")
  Exit Function
  
TestPageBreak_err:
  Resume TestPageBreak_end
End Function


'* Which report are we printing at the moment
Public Property Get ReportName() As String
  If (Not ReportControl.InReport) Or (ReportControl.CurReport < 1) Then Err.Raise ERR_NOACTIVEREPORT, "ReportName", "Cannot Get current ReportName. No active Report"
  ReportName = rpt(ReportControl.CurReport).ReportName
End Property

'* gives Total Pages of Current report or Last active report (if no current report)
Public Property Get ReportPages() As Long
  ReportPages = ReportControl.Pages_N
End Property

Public Sub AbortReport()
  ReportControl.AbortReport = True
End Sub

Public Property Get Zoom() As Long
  If Not ReportControl.InReport Then Err.Raise ERR_NOACTIVEREPORT, "Zoom", "Cannot Get zoom value. No active Report"
  Zoom = ReportControl.Zoom
End Property

Public Property Let Zoom(ByVal NewValue As Long)
  If NewValue > 1000 Or NewValue < 10 Then Call Err.Raise(ERR_INVALID_ZOOM_VALUE, "Zoom", "The zoom value for the reporter preview is invalid")
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "Zoom", "Cannot set zoom value. No active Report")
  ReportControl.Zoom = NewValue
  Call SetZoomLimit
End Property

Public Property Let IgnoreExportCR(ByVal NewValue As Boolean)
  If Not ReportControl.InReport Then Err.Raise ERR_NOACTIVEREPORT, "IgnoreExportCR", "Cannot set IgnoreExportCR as there is no active Report"
  ReportControl.IgnoreExportCR = NewValue
End Property

Public Property Let PreviewOnly(ByVal NewValue As Boolean)
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "PreviewOnly", "Cannot make text PreviewOnly as there is no active Report")
  ReportControl.NoRecord = NewValue
End Property

Public Function PreviewReport()
  Dim initok As Boolean, rTarget As PRIVATE_REPORT_TARGET
  
  On Error GoTo PreviewReport_err
  If ReportControl.InReport Then Exit Function
  If ReportControl.rTarget <> RPT_PREPARE Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "PreviewReport", "Cannot Preview a non Preview Report")
  If Not ReportControl.PreviewOK Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "PreviewReport", "Cannot Preview as Report not prepared for Preview")
  If ReportControl.Pages_N <= 0 Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "PreviewReport", "Cannot Preview as no Pages prepared for Preview")
  rTarget = ReportControl.rTarget
  initok = True
  ReportControl.InReport = True
  ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE
  ReportControl.CurPage = 1
  Call SetCursor(vbArrow)
  ReportControl.PreviewForm.Show vbModal
  Call ClearCursor
  
PreviewReport_end:
  ReportControl.InReport = False
  If initok Then
    ReportControl.rTarget = rTarget
  End If
  Exit Function
  
PreviewReport_err:
  Call ErrorMessage(ERR_ERROR, Err, "PreviewReport", "ERR_PREVIEW", "Error previewing report")
  Resume PreviewReport_end
End Function

Public Function PreviewPrintReport()
  Dim initok As Boolean, rTarget As PRIVATE_REPORT_TARGET
  
  On Error GoTo PreviewReport_err
  If ReportControl.InReport Then Exit Function
  If ReportControl.rTarget <> RPT_PREPARE Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "PreviewReport", "Cannot Preview a non Preview Report")
  If Not ReportControl.PreviewOK Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "PreviewReport", "Cannot Preview as Report not prepared for Preview")
  If ReportControl.Pages_N <= 0 Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "PreviewReport", "Cannot Preview as no Pages prepared for Preview")
  rTarget = ReportControl.rTarget
  initok = True
  ReportControl.InReport = True
  Call PreviewPrintPageEx(1, ReportControl.Pages_N)
  
PreviewReport_end:
  ReportControl.InReport = False
  If initok Then
    ReportControl.rTarget = rTarget
  End If
  Exit Function
  
PreviewReport_err:
  Call ErrorMessage(ERR_ERROR, Err, "PreviewReport", "ERR_PREVIEW", "Error previewing report")
  Resume PreviewReport_end
End Function

Public Property Get ReportActive() As Boolean
  ReportActive = ReportControl.InReport
End Property

Public Property Get ReportTarget() As REPORT_TARGET
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "ReportTarget", "Cannot get report target as there is no active report")
  If (ReportControl.rTarget <> RPT_PREPARE) And (ReportControl.rTarget <> RPT_PRINTER) Then Call Err.Raise(ERR_INVREPORTTARGET, "ReportTarget", "Cannot get report target as the current target is neither Print or Preview.")
  ReportTarget = ReportControl.rTarget
End Property

Public Property Let ReportGlobal(ByVal RptGlobal As REPORT_CONSTANTS, ByVal NewValue As String)
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "ReportGlobal", "Cannot set report globals as there is no active report")
  If (RptGlobal < 0) Or (RptGlobal > (REPORT_CONSTANTS_N - 1)) Then Call Err.Raise(ERR_SETREPORTCONSTANT, "ReportGlobal", "Cannot set report global " & CStr(RptGlobal) & " as the globals must be in the range 0 to " & CStr(REPORT_CONSTANTS_N - 1))
  Pages(ReportControl.CurPage).statics(RptGlobal) = NewValue
End Property

Public Property Let ReportPageGlobal(ByVal RptGlobal As REPORT_CONSTANTS, ByVal NewValue As Boolean)
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "ReportPageGlobal", "Cannot set report page globals flag as there is no active report")
  If (RptGlobal < 0) Or (RptGlobal > (REPORT_CONSTANTS_N - 1)) Then Call Err.Raise(ERR_SETREPORTCONSTANT, "ReportPageGlobal", "Cannot set report page global flag " & CStr(RptGlobal) & " as the globals must be in the range 0 to " & CStr(REPORT_CONSTANTS_N - 1))
  PageStatics(RptGlobal) = NewValue
End Property

Public Property Let ReportPageGlobalDefault(ByVal RptGlobal As REPORT_CONSTANTS, ByVal NewValue As String)
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "ReportPageGlobalDefault", "Cannot set report page globals default as there is no active report")
  If (RptGlobal < 0) Or (RptGlobal > (REPORT_CONSTANTS_N - 1)) Then Call Err.Raise(ERR_SETREPORTCONSTANT, "ReportPageGlobalDefault", "Cannot set report page global default " & CStr(RptGlobal) & " as the globals must be in the range 0 to " & CStr(REPORT_CONSTANTS_N - 1))
  PageStaticsDefault(RptGlobal) = NewValue
End Property

Public Function ReportTextWidth(String1 As String) As Single
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "ReportTextWidth", "Cannot get report text width as there is no active report")
  ReportTextWidth = GetTextWidth(String1)
End Function

Public Function ReportPageWidth() As Single
  If Not ReportControl.InReport Then Call Err.Raise(ERR_NOACTIVEREPORT, "ReportTextWidth", "Cannot get report text width as there is no active report")
  ReportPageWidth = ReportControl.PageWidth
End Function

Public Function ExportAvailable(ByVal ExportType As REPORT_EXPORTS) As Boolean
  Call xSet("ExportAvailable")
  ExportAvailable = ExportAvailableEx(ExportType)
  Call xReturn("ExportAvailable")
End Function

Public Function PrinterValid() As Boolean
  PrinterValid = IsPrinterAvail(True)  ' force printer detection
End Function

Public Function WrapTextToWidth(WrapString As String, ByVal WrapX As Single, Optional ByVal BreakChars As String = " ,;:-=") As Long
  Call xSet("WrapTextToWidth")
  If Not ReportControl.InReport Then Call IsPrinterAvail(True)  ' force printer detection
  WrapTextToWidth = WrapTextToWidthEx(WrapString, WrapX, BreakChars)
  Call xReturn("WrapTextToWidth")
End Function

Public Function ExportReport(ByVal FileName As String, ByVal ExportType As REPORT_EXPORTS, Optional ByVal Overwrite As Boolean = False) As Boolean
  On Error GoTo ExportReport_err
  Call xSet("ExportReport")
  
  If ReportControl.rTarget <> RPT_PREPARE Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export a Report that has not been prepared")
  If Not ReportControl.PreviewOK Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export a Report that has not been prepared for preview")
  If ReportControl.Pages_N <= 0 Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export as no pages prepared")
  If ReportControl.InReport Then Exit Function
  ExportReport = ExportReportEx(FileName, ExportType, Overwrite, False)
  
ExportReport_end:
  Call xReturn("ExportReport")
  Exit Function
  
ExportReport_err:
  ExportReport = False
  Call ErrorMessage(ERR_ERROR, Err, "ExportReport", "ERR_EXPORT", "Error exporting report to " & FileName)
  Resume ExportReport_end
End Function

Public Function ExportReportStr(ByVal ExportType As REPORT_EXPORTS) As String
  Dim sExport As String
  On Error GoTo ExportReportStr_err
  
  Call xSet("ExportReportStr")
  If ReportControl.rTarget <> RPT_PREPARE Then Err.Raise ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export a Report that has not been prepared"
  If Not ReportControl.PreviewOK Then Err.Raise ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export a Report that has not been prepared for preview"
  If ReportControl.Pages_N <= 0 Then Err.Raise ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export as no pages prepared"
  If ReportControl.InReport Then Exit Function
  If (ExportType = EXPORT_EXCEL) Or (ExportType = EXPORT_WORD) Then Err.Raise ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export as the target (" & ExportType & ") is invalid"
  If ExportReportEx(sExport, ExportType, False, True) Then ExportReportStr = sExport
    
ExportReportStr_end:
  Call xReturn("ExportReportStr")
  Exit Function
  
ExportReportStr_err:
  Call ErrorMessage(ERR_ERROR, Err, "ExportReport", "ERR_EXPORT", "Error exporting report to a string")
  Resume ExportReportStr_end
End Function

Private Sub Class_Initialize()
  Dim i As Long
  
  If RepInitCount = 0 Then
    Set ReportParser = New Parser
    Set ReportParser.ParseSettings = New RepParseSettings
    Call ReportParser.AddParseItem(New ParseArial)
    Call ReportParser.AddParseItem(New ParseBeginSection)
    Call ReportParser.AddParseItem(New ParseEndSection)
    Call ReportParser.AddParseItem(New ParseBoldOn)
    Call ReportParser.AddParseItem(New ParseBoldOff)
    Call ReportParser.AddParseItem(New ParseBox)
    Call ReportParser.AddParseItem(New ParseBW)
    Call ReportParser.AddParseItem(New ParseBWTEXTBOX)
    Call ReportParser.AddParseItem(New ParseBWTEXTBOXL)
    Call ReportParser.AddParseItem(New ParseBWTEXTBOXR)
    Call ReportParser.AddParseItem(New ParseCenter)
    Call ReportParser.AddParseItem(New ParseCenterPage)
    Call ReportParser.AddParseItem(New ParseComma)
    Call ReportParser.AddParseItem(New ParseDate)
    Call ReportParser.AddParseItem(New ParseDummy)
    Call ReportParser.AddParseItem(New ParseEndSkipExport)
    'Call ReportParser.AddParseItem(New ParseEndPFooter)
    'Call ReportParser.AddParseItem(New ParseEndPHeader)
    'Call ReportParser.AddParseItem(New ParseEndRFooter)
    'Call ReportParser.AddParseItem(New ParseEndRHeader)
    Call ReportParser.AddParseItem(New ParseEOLN)
    Call ReportParser.AddParseItem(New ParseEscape)
    Call ReportParser.AddParseItem(New ParseFillRGB)
    Call ReportParser.AddParseItem(New ParseFixedL)
    Call ReportParser.AddParseItem(New ParseFixedR)
    Call ReportParser.AddParseItem(New ParseFont)
    Call ReportParser.AddParseItem(New ParseForeRGB)
    Call ReportParser.AddParseItem(New ParseKEEPTOGETHER)
    Call ReportParser.AddParseItem(New ParseLeft)
    Call ReportParser.AddParseItem(New ParseLeftMarginABS)
    Call ReportParser.AddParseItem(New ParseLeftMargin)
    Call ReportParser.AddParseItem(New ParseLine)
    Call ReportParser.AddParseItem(New ParseLineAbs)
    Call ReportParser.AddParseItem(New ParseLineWidth)
    Call ReportParser.AddParseItem(New ParseNewPage)
    Call ReportParser.AddParseItem(New ParsePage)
    Call ReportParser.AddParseItem(New ParsePop)
    Call ReportParser.AddParseItem(New ParsePush)
    Call ReportParser.AddParseItem(New ParsePushX)
    Call ReportParser.AddParseItem(New ParsePushY)
    Call ReportParser.AddParseItem(New ParseReset)
    Call ReportParser.AddParseItem(New ParseResetColors)
    Call ReportParser.AddParseItem(New ParseRight)
    Call ReportParser.AddParseItem(New ParseRightMargin)
    Call ReportParser.AddParseItem(New ParseStartSkipExport)
    'Call ReportParser.AddParseItem(New ParseStartPFooter)
    'Call ReportParser.AddParseItem(New ParseStartPHeader)
    'Call ReportParser.AddParseItem(New ParseStartRFooter)
    'Call ReportParser.AddParseItem(New ParseStartRHeader)
    Call ReportParser.AddParseItem(New ParseStaticX)
    Call ReportParser.AddParseItem(New ParseTime)
    Call ReportParser.AddParseItem(New ParseTimes)
    Call ReportParser.AddParseItem(New ParseTrim)
    Call ReportParser.AddParseItem(New ParseVAR)
    Call ReportParser.AddParseItem(New ParseVLine)
    Call ReportParser.AddParseItem(New ParseWB)
    Call ReportParser.AddParseItem(New ParseWBTEXTBOXL)
    Call ReportParser.AddParseItem(New ParseWBTEXTBOXR)
    Call ReportParser.AddParseItem(New ParseWingDings)
    Call ReportParser.AddParseItem(New ParseX)
    Call ReportParser.AddParseItem(New ParseXAbs)
    Call ReportParser.AddParseItem(New ParseXRel)
    Call ReportParser.AddParseItem(New ParseY)
    Call ReportParser.AddParseItem(New ParseYAbs)
    Call ReportParser.AddParseItem(New ParseYRel)
    
    For i = 1 To MAX_POINT_SIZE
      FontHeights(i) = i * 20!
    Next i
    Set ReportControl.PreviewForm = New frmPreview
    COLOR_WHITE = rgb(255, 255, 255)
    DefaultFontStyle.FontType = VALID_FONT_TYPE
    DefaultFontStyle.Name = "Arial"
    DefaultFontStyle.Size = 10&
    DefaultFontStyle.bold = False
    DefaultFontStyle.UnderLine = False
    DefaultFontStyle.StrikeThrough = False
    DefaultFontStyle.Italic = False
    DefaultFontStyle.FontHeight = FontHeights(10)
    DefaultFontStyle.Align = ALIGN_LEFT
    Set PrinterFonts = New Collection
    Call RegisterLibrary(Me)
    Call IsPrinterAvail(True)
  End If
  RepInitCount = RepInitCount + 1
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  RepInitCount = RepInitCount - 1
  If RepInitCount = 0 Then
    Call EndReportEx(True, True)
    Set ReportParser = Nothing
    Set ReportControl.PreviewForm = Nothing
  End If
End Sub

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "Report library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get ExportTypeStr(ByVal ExportType As REPORT_EXPORTS) As String
  ExportTypeStr = ExportTypeStrEx(ExportType)
End Property

Public Property Get ExportTypeExt(ByVal ExportType As REPORT_EXPORTS) As String
  ExportTypeExt = ExportTypeExtEx(ExportType)
End Property
