Attribute VB_Name = "Const"
Option Explicit

Public Enum PRINTDLG_OPTIONS
  PAGES_ALL = 0
  PAGES_RANGE
  PAGES_CURRENT
End Enum

Public Enum PRIVATE_REPORT_TARGET
  RPT_PREPARE = PREPARE_REPORT
  RPT_PRINTER = PRINT_REPORT
  RPT_PREVIEW_PRINT = 1024
  RPT_PREVIEW_DISPLAYPAGE
  RPT_CONFIG
  
  RPT_EXPORT = 4096
  RPT_EXPORT_CSV = EXPORT_CSV
  RPT_EXPORT_EXCEL = EXPORT_EXCEL
  RPT_EXPORT_WORD = EXPORT_WORD
  RPT_EXPORT_FIXEDWIDTH = EXPORT_FIXEDWIDTH
  RPT_EXPORT_PDF = EXPORT_PDF
End Enum


Public Enum RepLibraryErrors
  ERR_REPORT_STACK = TCSREPORTER_ERROR
  ERR_NOPRINTER
  ERR_NO_INIT_REPORT
  ERR_NOACTIVEREPORT
  ERR_INVALID_ZOOM_VALUE
  ERR_PARSEFONT
  ERR_NOTPREVIEWREPORT
  ERR_POPFONTSTACK
  ERR_PERCENT
  ERR_PRINTRANGE
  ERR_POPCOLORSTACK
  ERR_PARSELINE
  ERR_POPCOORDSTACK
  ERR_PARSEBOX
  ERR_NOPRINT
  ERR_PAGERECURSIVE
  ERR_SETREPORTCONSTANT
  ERR_REPORTCONSTANT
  ERR_EXPORT
  ERR_INVALID_TARGET
  ERR_SETHEADERFOOTER
  ERR_PARSELINEWIDTH
  ERR_INVREPORTTARGET
  ERR_MARGIN
  ERR_NEWPAGE
  ERR_BEGINSECTION
  ERR_ENDSECTION
  ERR_KEEPTOGETHER
  ERR_GETCHARWIDTHS
  ERR_INVALIDFONT
  ERR_A4
  'ERR_INVALID_PREVIEW_OBJECT = TCSREPORTER_ERROR + 1
  'ERR_NO_INIT_REPORT = TCSREPORTER_ERROR + 2
  'ERR_REPORT_CANCEL = TCSREPORTER_ERROR + 3
End Enum


Public Type Coordinate
  x As Single
  y As Single
  FirstX As Single
End Type

Public Enum FONT_TYPE
  INVALID_FONT_TYPE = 0
  VALID_FONT_TYPE
  PREVIEW_SCREEN_FONT
  SCREEN_FONT
  PRINTER_FONT
End Enum

Public Type FontStyle
  Name As String
  Size As Single
  bold As Boolean
  Italic As Boolean
  UnderLine As Boolean
  StrikeThrough As Boolean
  FontHeight As Single
  Align As ALIGNMENT_TYPE
  FontType As FONT_TYPE
End Type

Public Type ColorType
  ColorValid  As Boolean
  ColorSet As Boolean
  ColorFill As Boolean
  FillColor As Long
  LineColor As Long
  FillStyle As Long
  ForeColor As Long
End Type

' Reporter type declarations
Public Type REPORT_LEVEL
  ReportName As String
  RHeader As String
  RFooter As String
  RFooterH As Single
  
  PHeader As String
  PFooter As String
  PFooterH As Single
  
  OverrideHeader As Boolean
  OverrideFooter As Boolean
End Type

Public Type HTMLOUTPUT
  HTMLString As QString
  OpenDiv As Boolean
  CloseDiv As Boolean
  HTMLFontSet As Boolean
  CurrentX As Single
  CurrentY As Single
  XSetHTML As Boolean
  YSetHTML As Boolean
  Position As Boolean
  FillColor As String
  OrientationString As String
  TopString As String
  BottomString As String
  LeftString As String
  RightString As String
  ContactString As String
  ReportPages As Long
End Type

Public Type REPORT_STATICS
  rTarget As PRIVATE_REPORT_TARGET
  PreviewForm As frmPreview
  Preview As PictureBox
  PreviewTest As PictureBox
    
  Name As String              ' Top Level report name
  ReportDateTime As Date      ' Report start date/time
  CurReport As Long           ' Current Report Level 0=no report active
  CurPage As Long             ' Current report page
  CurPageValid As Boolean     ' Current report page valid (i.e. has text on it)
  Pages_N As Long             ' Total Number of pages
  
  ThrowPageOnSubReport As Boolean
  SuppressNewPageCalc As Boolean
  ' prevent the reporter triggering a New Page at the end of
  ' the current page (used when printing Footers etc)
  DelimitOut As Boolean       ' Delimit the output in PREPARE
  NoRecord As Boolean         ' Do not record this Out Statement if in Preview
  
  ExportHeader As String
  IgnoreExportCR As Boolean
  
  PageHeight As Single
  PageWidth As Single
  
  LeftMargin As Single   ' left margin on the page, default = 1
  RightMargin As Single  ' right margin on the page, default = 0
  
  FirstX As Single       ' X coord after newline, used for STATICX, default = 0
  TrimX As Single        ' trim text
  CenterX As Single
  Zoom As Long
  ZoomLimit As Long
  
  ' support for BEGIN/END SECTION
  BeginSectionX As Single
  BeginSectionY As Single
  OutputSection As Boolean
  Section As String
  
  ' support for fixed width exports
  FixedWidth As Long     ' width of column output
  FW_PadLeft As Boolean ' left/right alignment
  
  Rem track current style for report
  InReport As Boolean     ' am I currently in a report (nested or otherwise)
  PreviewOK As Boolean   ' Finished preparing report
  Orientation As REPORT_ORIENTATION
  
  PrinterNewPage As Boolean  ' when printing flags whether a new page is neccessary
  OnNewPage As Boolean
  
  Rem turn off printing completely
  NoOutput As Boolean
  AbortReport As Boolean
  
  'Report timing and feedback
  StartTime As Long
  EndTime As Long
  
  Rem Track output of Headers/Footers if no body then allow possibility of
  Rem not printing anything
  PageHeaderPrinted As Boolean
  PageTextPrinted As Boolean
  PageFooterPrinted As Boolean
  ReportHeaderPrinted As Boolean
  ReportFooterPrinted As Boolean
  
  'Current Style
  fStyle As FontStyle
  FColor As ColorType
  
  ' PrintDlg settings
  PrintDlgOpt As PRINTDLG_OPTIONS
  PageTo As Long
  PageFrom As Long
  PageCopies As Long
  HTML As HTMLOUTPUT

End Type

Public Const MAXRPTACTIONSTR As Long = 20
Public Const MAX_REPORT_LEVELS As Long = 10
Public Const MAX_POINT_SIZE As Long = 36
Public Const FontMapError As String = "Font Mapping required"
Public Const MAX_PAGE_COPIES As Long = 99

Public PrinterFonts As Collection
Public FontHeights(1 To MAX_POINT_SIZE) As Single
Public DefaultFontStyle As FontStyle
Public RepInitCount As Long
Public ReportParser As Parser

Public rpt(1 To MAX_REPORT_LEVELS) As REPORT_LEVEL
Public ReportControl As REPORT_STATICS

Public NAV_HEADER As Single


Public Const L_LAST_EXPORT As Long = EXPORT_PDF
Public m_LastReportControl As REPORT_STATICS
Private m_Compare As Boolean
Public Const PDF_DRIVER_ABACUS As String = "Abacus PDF Converter"
Public Const PDF_DRIVER_ONE_SOURCE As String = "ONESOURCE PDF Converter"
Public Const PDF_DRIVER_SAGE As String = "Sage PDF Converter"

Public PDF_DRIVER_NAMES() As String
Public g_A4Force As Boolean

Public Sub ReportControlCompare()

If (Not m_Compare) Then
  m_Compare = True
  Exit Sub
End If

Dim r1 As REPORT_STATICS
Dim r2 As REPORT_STATICS


r1 = m_LastReportControl
r2 = ReportControl
  
If r1.rTarget <> r2.rTarget Then Call report_rtarget_diff("rTarget")
If r1.Preview <> r2.Preview Then Call report_rtarget_diff("Preview")
If r1.PreviewTest <> r2.PreviewTest Then Call report_rtarget_diff("PreviewTest")
If r1.Name <> r2.Name Then Call report_rtarget_diff("Name")
If r1.ReportDateTime <> r2.ReportDateTime Then Call report_rtarget_diff("ReportDateTime")
If r1.CurReport <> r2.CurReport Then Call report_rtarget_diff("CurReport")
If r1.CurPage <> r2.CurPage Then Call report_rtarget_diff("CurPage")
If r1.CurPageValid <> r2.CurPageValid Then Call report_rtarget_diff("CurPageValid")
If r1.Pages_N <> r2.Pages_N Then Call report_rtarget_diff("Pages_N")
If r1.ThrowPageOnSubReport <> r2.ThrowPageOnSubReport Then Call report_rtarget_diff("ThrowPageOnSubReport")
If r1.SuppressNewPageCalc <> r2.SuppressNewPageCalc Then Call report_rtarget_diff("SuppressNewPageCalc")
If r1.DelimitOut <> r2.DelimitOut Then Call report_rtarget_diff("DelimitOut")
If r1.NoRecord <> r2.NoRecord Then Call report_rtarget_diff("NoRecord")
If r1.ExportHeader <> r2.ExportHeader Then Call report_rtarget_diff("ExportHeader")
If r1.IgnoreExportCR <> r2.IgnoreExportCR Then Call report_rtarget_diff("IgnoreExportCR")
If r1.PageHeight <> r2.PageHeight Then Call report_rtarget_diff("PageHeight")
If r1.PageWidth <> r2.PageWidth Then Call report_rtarget_diff("PageWidth")
If r1.LeftMargin <> r2.LeftMargin Then Call report_rtarget_diff("LeftMargin")
If r1.RightMargin <> r2.RightMargin Then Call report_rtarget_diff("RightMargin")
If r1.FirstX <> r2.FirstX Then Call report_rtarget_diff("FirstX")
If r1.TrimX <> r2.TrimX Then Call report_rtarget_diff("TrimX")
If r1.CenterX <> r2.CenterX Then Call report_rtarget_diff("CenterX")
If r1.Zoom <> r2.Zoom Then Call report_rtarget_diff("Zoom")
If r1.ZoomLimit <> r2.ZoomLimit Then Call report_rtarget_diff("ZoomLimit")
If r1.BeginSectionX <> r2.BeginSectionX Then Call report_rtarget_diff("BeginSectionX")
If r1.BeginSectionY <> r2.BeginSectionY Then Call report_rtarget_diff("BeginSectionY")
If r1.OutputSection <> r2.OutputSection Then Call report_rtarget_diff("OutputSection")
If r1.Section <> r2.Section Then Call report_rtarget_diff("Section")
If r1.FixedWidth <> r2.FixedWidth Then Call report_rtarget_diff("FixedWidth")
If r1.FW_PadLeft <> r2.FW_PadLeft Then Call report_rtarget_diff("FW_PadLeft")
If r1.InReport <> r2.InReport Then Call report_rtarget_diff("InReport")
If r1.PreviewOK <> r2.PreviewOK Then Call report_rtarget_diff("PreviewOK")
If r1.Orientation <> r2.Orientation Then Call report_rtarget_diff("Orientation")
If r1.PrinterNewPage <> r2.PrinterNewPage Then Call report_rtarget_diff("PrinterNewPage")
If r1.OnNewPage <> r2.OnNewPage Then Call report_rtarget_diff("OnNewPage")
If r1.NoOutput <> r2.NoOutput Then Call report_rtarget_diff("NoOutput")
If r1.AbortReport <> r2.AbortReport Then Call report_rtarget_diff("AbortReport")
If r1.StartTime <> r2.StartTime Then Call report_rtarget_diff("StartTime")
If r1.EndTime <> r2.EndTime Then Call report_rtarget_diff("EndTime")
If r1.PageHeaderPrinted <> r2.PageHeaderPrinted Then Call report_rtarget_diff("PageHeaderPrinted")
If r1.PageTextPrinted <> r2.PageTextPrinted Then Call report_rtarget_diff("PageTextPrinted")
If r1.PageFooterPrinted <> r2.PageFooterPrinted Then Call report_rtarget_diff("PageFooterPrinted")
If r1.ReportHeaderPrinted <> r2.ReportHeaderPrinted Then Call report_rtarget_diff("ReportHeaderPrinted")
If r1.ReportFooterPrinted <> r2.ReportFooterPrinted Then Call report_rtarget_diff("ReportFooterPrinted")
'If r1.fStyle <> r2.fStyle Then Call report_rtarget_diff("fStyle")
'If r1.FColor <> r2.FColor Then Call report_rtarget_diff("FColor")
If r1.PrintDlgOpt <> r2.PrintDlgOpt Then Call report_rtarget_diff("PrintDlgOpt")
If r1.PageTo <> r2.PageTo Then Call report_rtarget_diff("PageTo")
If r1.PageFrom <> r2.PageFrom Then Call report_rtarget_diff("PageFrom")
If r1.PageCopies <> r2.PageCopies Then Call report_rtarget_diff("PageCopies")
'If r1.HTML <> r2.HTML Then Call report_rtarget_diff("HTML")

End Sub
Public Sub report_rtarget_diff(targetproperty As String)
  Debug.Print targetproperty
End Sub
