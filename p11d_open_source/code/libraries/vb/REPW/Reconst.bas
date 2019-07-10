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


Public Const L_LAST_EXPORT As Long = EXPORT_HTML_INTEXP5
