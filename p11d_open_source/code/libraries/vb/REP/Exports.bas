Attribute VB_Name = "Exports"
Option Explicit

Public Enum EXPORT_OUT
  OUT_TEXT
  OUT_CR
  OUT_COMMA
End Enum

Private Enum PDF_PRINTER_AVAILABLE
  NOT_SEARCHED
  NOT_AVAILABLE
  AVAILABLE
End Enum


Private Const MAX_EXCEL_NAME As Long = 18
Private m_Filename As String
Private m_File As Integer
Private m_NewLn As Boolean

' support for fixed width exports
Private m_CurColumn As String
Private m_CurColumnWidth As Long


Private m_Excel As Object 'Excel.Application
Private m_ExcelWB As Object 'Excel.Workbook
Private r1 As Long, c1 As Long
Private m_CellValues() As Variant
Private m_CellValuesNumberFormats() As String
Private m_ColCount As Long


Private m_Word As Word.Application
Private m_WordDoc As Word.Document
Private m_Rng As Word.Range
Public g_cdi As CDIntfEx.CDIntfEx
Public g_cdi_vertical_margin As Long
Public g_cdi_horizontal_margin As Long
Public g_cdi_paper_size As Long
Private m_PDFPrinterAvailable As PDF_PRINTER_AVAILABLE
Private m_PDFPrinterName As String
Private m_LastPrinterName As String
Private m_PDFPrinterLicensee As String
Private m_PDFPrinterLicCode As String

Private Function ClearExports()
  On Error Resume Next
  If m_File > 0 Then
    Close m_File
    m_File = -1
  End If
  If Not m_Excel Is Nothing Then
    If Not m_ExcelWB Is Nothing Then
      If Len(m_ExcelWB.Path) > 0 Then
        Call m_ExcelWB.Close(True)
      Else
        Call m_ExcelWB.Close(True, m_Filename)
      End If
    End If
    Set m_ExcelWB = Nothing
    m_Excel.Quit
    Set m_Excel = Nothing
  End If
  If Not m_Word Is Nothing Then
    If Not m_WordDoc Is Nothing Then
      ' sort out column widths
      Set m_Rng = Nothing
      If m_WordDoc.Tables.Count > 0 Then
        If Val(m_Word.Version) >= 9 Then
          Dim tblObject As Object
          Set tblObject = m_WordDoc.Tables(1)
          Call tblObject.AutoFitBehavior(1) ' wdAutoFitContent = 1
        End If
      End If
      If Len(m_WordDoc.Path) = 0 Then
        Call m_WordDoc.SaveAs(m_Filename)
      End If
      Call m_WordDoc.Close(wdSaveChanges)
    End If
    Set m_WordDoc = Nothing
    m_Word.Quit
    Set m_Word = Nothing
  End If
  m_Filename = ""
  m_ColCount = 1
  ReDim m_CellValues(1 To 1) As Variant
  r1 = 1
  c1 = 1
  m_NewLn = True
  Set ReportControl.HTML.HTMLString = New QString
  ReportControl.HTML.HTMLString = ""
End Function

Private Sub SetMinColumnWidth(ByVal c As Word.Column)
  c.Width = 50
End Sub
      
Private Function OpenExports(fname As String, ByVal ExportType As REPORT_EXPORTS, ByVal Overwrite As Boolean, ByVal ReportAsString As Boolean) As Boolean
  Dim fExists As Boolean
  
  On Error GoTo OpenExports_err
  

  OpenExports = False
  If ReportAsString Then
    If Not IsExportHTML Then Err.Raise ERR_EXPORT, "OpenExports", "Unable to export the export type " & ExportTypeStrEx(ExportType) & " to a string"
  Else
    m_Filename = fname
    If Overwrite And FileExists(m_Filename) Then Call xKill(m_Filename)
    fExists = FileExists(m_Filename)
  End If
  Select Case ExportType
    Case EXPORT_CSV, EXPORT_FIXEDWIDTH
      m_File = FreeFile
      Open fname For Append As m_File
      OpenExports = True
    Case EXPORT_EXCEL
      Set m_Excel = CreateObject("Excel.Application")
      If fExists Then
        Set m_ExcelWB = m_Excel.Workbooks.Open(m_Filename)
      Else
        Set m_ExcelWB = m_Excel.Workbooks.Add
      End If
      Call AddExcelSheet
      OpenExports = True
    Case EXPORT_WORD
      Set m_Word = New Word.Application
      If fExists Then
        Set m_WordDoc = m_Word.Documents.Open(m_Filename)
      Else
        Set m_WordDoc = m_Word.Documents.Add
      End If
      Set m_Rng = m_Word.ActiveDocument.Range(0, 0)
      m_Word.ActiveDocument.PageSetup.Orientation = wdOrientLandscape
      If Val(m_Word.Version) >= 9 Then
        Dim tblsObject As Object
        Set tblsObject = m_Word.ActiveDocument.Tables
        Call tblsObject.Add(m_Rng, 1, 1, , 0)  ' wdAutoFitFixed = 0
      Else
        Call m_Word.ActiveDocument.Tables.Add(m_Rng, 1, 1)
      End If
      Call SetMinColumnWidth(m_Word.ActiveDocument.Tables(1).Columns.Item(1))
      OpenExports = True
    Case EXPORT_HTML_IE, EXPORT_HTML_NETSCAPE, EXPORT_HTML_INTEXP5  'km
      If Not ReportAsString Then
        m_File = FreeFile
        Open fname For Append As m_File
      End If
      Call ConstructHTMLHeader(ReportControl.Name)
      OpenExports = True
    Case EXPORT_PDF
      OpenExports = True
    Case Else
      Call ECASE("OpenExports - ExportType not supported")
  End Select
  
OpenExports_end:
  Exit Function
  
OpenExports_err:
  OpenExports = False
  Resume OpenExports_end
  Resume
End Function

Private Sub AddExcelSheet()
  Dim ws As Object 'Excel.Worksheet
  Dim i As Long
  Dim sname As String
   
  On Error GoTo AddExcelSheet_Err
  sname = ReportControl.Name
  If Len(sname) > MAX_EXCEL_NAME Then sname = Left$(sname, MAX_EXCEL_NAME) & "..."
  'apf add after
  Call m_ExcelWB.Worksheets.Add
  
  On Error GoTo AddExcelSheet_NErr
  Do
    Set ws = m_ExcelWB.Worksheets(sname)
    i = i + 1
    Set ws = m_ExcelWB.Worksheets(sname & "_" & CStr(i))
  Loop Until i >= 10
  If i >= 10 Then Call Err.Raise(ERR_EXPORT, "OpenExports", "Unable to export as the cannot create the worksheet 2 " & sname)
  
AddExcelSheet_end:
  Exit Sub
  
AddExcelSheet_Err:
  Call Err.Raise(ERR_EXPORT, "OpenExports", "Unable to export as the cannot create the worksheet 1 " & sname)
  Resume AddExcelSheet_end
  
AddExcelSheet_NErr:
  If i > 0 Then sname = sname & "_" & CStr(i)
  m_ExcelWB.ActiveSheet.Name = sname
  Resume AddExcelSheet_end
End Sub

Public Function ExportTypeStrEx(ByVal ExportType As REPORT_EXPORTS) As String
  Select Case ExportType
    Case EXPORT_CSV
      ExportTypeStrEx = "CSV Export"
    Case EXPORT_FIXEDWIDTH
      ExportTypeStrEx = "Fixed Width Export"
    Case EXPORT_EXCEL
      ExportTypeStrEx = "Excel Export"
    Case EXPORT_WORD
      ExportTypeStrEx = "Word Export"
    Case EXPORT_HTML_IE
      ExportTypeStrEx = "HTML Export (IE4+)"
    Case EXPORT_HTML_NETSCAPE
      ExportTypeStrEx = "HTML Export (Netscape4+)"
    Case EXPORT_HTML_INTEXP5
      ExportTypeStrEx = "HTML Export (IE5)"
    Case EXPORT_PDF
      ExportTypeStrEx = "PDF Export"
    Case Else
      ExportTypeStrEx = "Unknown Export"
      
  End Select
End Function

Public Function ExportTypeExtEx(ByVal ExportType As REPORT_EXPORTS) As String
  Select Case ExportType
    Case EXPORT_CSV
      ExportTypeExtEx = ".txt"
    Case EXPORT_FIXEDWIDTH
      ExportTypeExtEx = ".fwt"
    Case EXPORT_EXCEL
      ExportTypeExtEx = ".xls"
    Case EXPORT_WORD
      ExportTypeExtEx = ".doc"
    Case EXPORT_HTML_IE
      ExportTypeExtEx = ".htm"
    Case EXPORT_HTML_NETSCAPE
      ExportTypeExtEx = ".htm"
    Case EXPORT_HTML_INTEXP5  'km
      ExportTypeExtEx = ".htm"
    Case EXPORT_PDF
      ExportTypeExtEx = ".pdf"
    Case Else
      ExportTypeExtEx = "Unknown Export"
  End Select
End Function

Public Sub ExportOut(ByVal OutCtrl As EXPORT_OUT, ByVal String1 As String)
  Dim TextWidth As Single
  Dim i As Long
        
        
  On Error GoTo exportout_err
  If OutCtrl = OUT_TEXT Then
    String1 = RemoveChar(String1, vbLf)
    String1 = RemoveChar(String1, vbCr)
    If Len(String1) = 0 Then Exit Sub
  End If
  If ReportControl.NoOutput Then Exit Sub
  Select Case ReportControl.rTarget
    Case EXPORT_CSV
      If OutCtrl = OUT_TEXT Then
        If m_NewLn Then
          Print #m_File, """";
          m_NewLn = False
        End If
        Print #m_File, String1;
      ElseIf OutCtrl = OUT_COMMA Then
        If m_NewLn Then
          Print #m_File, """";
          m_NewLn = False
        End If
        Print #m_File, """,""";
      ElseIf OutCtrl = OUT_CR Then
        If m_NewLn Then
          Print #m_File,
        Else
          Print #m_File, """"
        End If
        m_NewLn = True
      End If
    Case EXPORT_FIXEDWIDTH
      If OutCtrl = OUT_TEXT Then
        m_CurColumn = m_CurColumn & String1
        m_CurColumnWidth = m_CurColumnWidth + Len(String1)
      ElseIf (OutCtrl = OUT_CR) Or (OutCtrl = OUT_COMMA) Then
        If ReportControl.FW_PadLeft Then
          If m_CurColumnWidth < ReportControl.FixedWidth Then
            String1 = String$(ReportControl.FixedWidth - m_CurColumnWidth, " ") & m_CurColumn
          Else
            String1 = Right$(m_CurColumn, ReportControl.FixedWidth)
          End If
        Else
          If m_CurColumnWidth < ReportControl.FixedWidth Then
            String1 = m_CurColumn & String$(ReportControl.FixedWidth - m_CurColumnWidth, " ")
          Else
            String1 = Left$(m_CurColumn, ReportControl.FixedWidth)
          End If
        End If
        Print #m_File, String1;
        m_CurColumn = "": m_CurColumnWidth = 0
        If OutCtrl = OUT_CR Then Print #m_File,
      End If
    Case EXPORT_EXCEL
      If OutCtrl = OUT_TEXT Then
        m_CellValues(c1) = m_CellValues(c1) & String1
      ElseIf OutCtrl = OUT_COMMA Then
        c1 = c1 + 1
        If c1 > m_ColCount Then
          m_ColCount = m_ColCount + 1
          ReDim Preserve m_CellValues(1 To m_ColCount) As Variant
        End If
        m_CellValues(c1) = ""
      ElseIf OutCtrl = OUT_CR Then
        For i = LBound(m_CellValues) To UBound(m_CellValues)
          If IsNumeric(m_CellValues(i)) Then
            m_CellValues(i) = CDbl(m_CellValues(i))
          ElseIf IsDate(m_CellValues(i)) Then
            m_CellValues(i) = CDate(m_CellValues(i))
          End If
        Next i
        m_ExcelWB.ActiveSheet.Range(m_ExcelWB.ActiveSheet.Cells(r1, 1), m_ExcelWB.ActiveSheet.Cells(r1, c1)).value = m_CellValues
        c1 = 1: r1 = r1 + 1
        m_CellValues(c1) = ""
      End If
    Case EXPORT_WORD
      'Call ECASE("ExportAvailableEx not yet implemented")
      If OutCtrl = OUT_TEXT Then
        m_CellValues(c1) = m_CellValues(c1) & String1
        m_Word.ActiveDocument.Tables(1).Cell(r1, c1).Range.Text = m_CellValues(c1)
      ElseIf OutCtrl = OUT_COMMA Then
        If IsNumeric(m_CellValues(c1)) Then
          m_CellValues(c1) = CDbl(m_CellValues(c1))
        End If
        c1 = c1 + 1
        If c1 > m_ColCount Then
          m_ColCount = m_ColCount + 1
          ReDim Preserve m_CellValues(1 To m_ColCount) As Variant
          Call SetMinColumnWidth(m_Word.ActiveDocument.Tables(1).Columns.Add)
          'Call m_Word.ActiveDocument.Tables(1).Columns.DistributeWidth
        End If
        m_CellValues(c1) = ""
      ElseIf OutCtrl = OUT_CR Then
        Call m_Word.ActiveDocument.Tables(1).Rows.Add
        c1 = 1
        r1 = r1 + 1
        m_CellValues(c1) = ""
      End If
    Case EXPORT_HTML_IE, EXPORT_HTML_NETSCAPE, EXPORT_HTML_INTEXP5  'km
      If OutCtrl = OUT_TEXT Then
        Call CloseDiv
        Call SetOpenDiv
        Call SetFontHTML
        If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
          Call ReportControl.HTML.HTMLString.Append("""")
        End If
        Call SetXHTML
        Call SetYHTML
        Call CloseOpenDiv
        Call ReportControl.HTML.HTMLString.Append("<SPAN STYLE=""")
        ReportControl.HTML.HTMLFontSet = False
        ReportControl.HTML.OpenDiv = True
        Call SetFontHTML
        ReportControl.HTML.OpenDiv = False
        Call ReportControl.HTML.HTMLString.Append(""">")
        Call SetHTMLSpaces(String1)
        Call ReportControl.HTML.HTMLString.Append(String1)
        Call ReportControl.HTML.HTMLString.Append("</SPAN>")
        ReportControl.HTML.CurrentX = ReportControl.HTML.CurrentX + GetTextWidth(String1)
        
      ElseIf OutCtrl = OUT_COMMA Then
        'Need to tabulate?
        
      ElseIf OutCtrl = OUT_CR Then
        Call CloseDiv
        ReportControl.HTML.CurrentX = 0
        ReportControl.HTML.CurrentY = ReportControl.HTML.CurrentY + ReportControl.fStyle.FontHeight
        If (ReportControl.rTarget = EXPORT_HTML_IE) Then
          If ReportControl.HTML.CurrentY > ReportControl.PageHeight Then
            Call SetNewHTMLPage
          End If
        End If
      End If
      
    Case Else
      Call ECASE("ExportAvailableEx - unknown export type")
  End Select
  Exit Sub
exportout_err:
  Err.Raise Err.Number, ErrorSource(Err, "ExportOut"), Err.Description
  Resume
End Sub
      

Public Function ExportAvailableEx(ByVal ExportType As REPORT_EXPORTS) As Boolean
  ExportAvailableEx = False
  
  'ExportType = EXPORT_PDF
  Select Case ExportType
    Case EXPORT_CSV, EXPORT_FIXEDWIDTH
           ExportAvailableEx = True
    Case EXPORT_EXCEL
           ExportAvailableEx = isCOMPresent("Excel.Application", WIN32_SERVERPROC)
    Case EXPORT_WORD
           'ExportAvailableEx = False
           ExportAvailableEx = isCOMPresent("Word.Application", WIN32_SERVERPROC)
    Case EXPORT_PDF
      ExportAvailableEx = IsPDFAvailable
    Case EXPORT_HTML_IE, EXPORT_HTML_NETSCAPE, EXPORT_HTML_INTEXP5  'km
           ExportAvailableEx = True
    Case Else
           Call ECASE("ExportAvailableEx - unknown export type")
  End Select
End Function
Private Function IsPDFAvailable() As Boolean
  Dim p As Printer
  Dim i As Long
    
  If (m_PDFPrinterAvailable = NOT_SEARCHED) Then
    m_PDFPrinterAvailable = NOT_AVAILABLE
    For Each p In Printers
      For i = 1 To UBound(PDF_DRIVER_NAMES)
        If IsPDFPrinterDriverEx(p.DeviceName, PDF_DRIVER_NAMES(i)) Then
          m_PDFPrinterAvailable = AVAILABLE
          m_PDFPrinterName = p.DeviceName 'PDF_DRIVER_NAMES(i)
        End If
      Next
    Next
  End If
  IsPDFAvailable = m_PDFPrinterAvailable = AVAILABLE
End Function
Public Function IsCurrentPrinterPDF() As Boolean
  Dim i As Long
  Dim s As String
  
  s = Printer.DeviceName
  For i = 1 To UBound(PDF_DRIVER_NAMES)
    If IsPDFPrinterDriverEx(s, PDF_DRIVER_NAMES(i)) Then
      IsCurrentPrinterPDF = True
      Exit Function
    End If
  Next

End Function
Public Function PDFAmPrinting() As Boolean

  PDFAmPrinting = Not g_cdi Is Nothing
End Function
Public Function A4Force() As Boolean
  A4Force = g_A4Force
End Function
Public Function A4NonPrintableMicroMeters() As Single
  A4NonPrintableMicroMeters = 60#
End Function

Public Function A4NonPrintableMArginTwips() As Single
  A4NonPrintableMArginTwips = (1440 * (A4NonPrintableMicroMeters() / 254)) '340.157480315
End Function
Public Function IsPDFPrinterDriverEx(ByVal printerName As String, ByVal ABACUS_PRINTER_DRIVER As String) As Boolean
  IsPDFPrinterDriverEx = StrComp(printerName, ABACUS_PRINTER_DRIVER) = 0
End Function

Public Function PDFDriverInstall() As Boolean
  Dim serialNumber As String
  
  On Error GoTo err_Err
  
  If (Not IsCurrentPrinterPDF()) Then GoTo err_End
  If IsPDFAvailable = False Then
    Call Err.Raise(ERR_ERROR, "Reporter", "Pdf driver is not available")
  End If
    
  If (g_cdi Is Nothing) Then
    'Printer.Orientation = 1 'EWPDF
    Set g_cdi = New CDIntfEx.CDIntfEx
  Else
    Call Err.Raise(ERR_ERROR, "Reporter", "Can not install pdf driver as has already been installed")
  End If
                  
  serialNumber = "07EFCDAB01000100E0370FBEA11533CE8484F6E47E12547615123438D593268E13F82AE12E446D05067F24A3DB404331D45D96E74588"
  If IsPDFPrinterDriverEx(m_PDFPrinterName, PDF_DRIVER_ABACUS) Or IsPDFPrinterDriverEx(m_PDFPrinterName, PDF_DRIVER_ONE_SOURCE) Then
    g_cdi.DriverInit m_PDFPrinterName 'PDF_DRIVER_ABACUS
    g_cdi.EnablePrinter "Thomson Reuters (Professional)", serialNumber '"07EFCDAB01000100BC59AEFEBF38B9649CF0EE64C144C982C50A46D02B2B3976AC9E4DA9C88520E529023E1508C89DECE79FA977A0FA6D486C3C40BC75D47E"
  ElseIf IsPDFPrinterDriverEx(m_PDFPrinterName, PDF_DRIVER_SAGE) Then
    g_cdi.DriverInit m_PDFPrinterName
    g_cdi.EnablePrinter "Sage (UK) Limited", serialNumber '"07EFCDAB0100010062AE5DBE1E7D10232F4872A235B5E34B5EFC7E6D2EE27CCF7ACBCBBB214A21D8F384C17CF890221DF0893AFD22EE4DDE0621AD1D1921F4"
  Else
      Call Err.Raise(ERR_NOPRINTER, "PDF Driver Install", "Invalid PDF Printer Name:" & m_PDFPrinterName)
  End If
  
  
  g_cdi_horizontal_margin = g_cdi.HorizontalMargin
  g_cdi_vertical_margin = g_cdi.VerticalMargin
  g_cdi_paper_size = g_cdi.PaperSize
  
  g_cdi.HorizontalMargin = A4NonPrintableMicroMeters
  g_cdi.VerticalMargin = A4NonPrintableMicroMeters
  g_cdi.PaperSize = 9 'A4
  PDFDriverInstall = True
  
err_End:
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "PDFDriverInstall"), Err.Description)
End Function
Public Function SetActivePrinter(ByVal Name As String) As String
  Dim prt As Printer
  Dim s As String, lastPrinter As String
  
On Error GoTo err_Err

  lastPrinter = Printer.DeviceName
  Name = LCase$(Name)
  For Each prt In Printers
    s = prt.DeviceName
    s = LCase$(s)
    If (s = Name) Then
      Set Printer = prt
      GoTo err_End
    End If
  Next
  
  Call Err.Raise(380, "SetActivePrinter", "PDF driver is not installed")
err_End:
  SetActivePrinter = lastPrinter
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, Err.Source, "Failed to set the active printer to " & Name & "." & Err.Description)
End Function
Public Sub PDFDriverUninstall()
  If (g_cdi Is Nothing) Then Call Err.Raise(ERR_ERROR, "Reporter", "Can not uninstall pdf driver as has not already been installed")
  
  g_cdi.HorizontalMargin = g_cdi_horizontal_margin
  g_cdi.VerticalMargin = g_cdi_vertical_margin
  g_cdi.PaperSize = g_cdi_paper_size
  'Call g_cdi.RestoreDefaultPrinter
  Call g_cdi.DriverEnd
  Set g_cdi = Nothing
End Sub
Public Function ExportReportEx(exportdest As String, ByVal ExportType As REPORT_EXPORTS, ByVal Overwrite As Boolean, ByVal ReportAsString As Boolean) As Boolean
  
  Static InExport As Boolean
  Dim rTarget As REPORT_TARGET, CurPage As Long, Opened As Boolean
  Dim pdfExport As Boolean
  Dim bSetActivePrinter As Boolean
  Dim cdi As CDIntfEx.CDIntfEx
  Dim lastPrinter As String
  
  On Error GoTo ExportReportEx_err
  
  If ReportControl.rTarget <> RPT_PREPARE And ReportControl.rTarget <> RPT_PREVIEW_DISPLAYPAGE Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export a Report that has not been prepared")
  If Not ReportControl.PreviewOK Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export a Report that has not been prepared for preview")
  If ReportControl.Pages_N <= 0 Then Call Err.Raise(ERR_NOTPREVIEWREPORT, "ExportReport", "Cannot Export as no pages prepared")

  Opened = False
  Call SetCursor
  If Not InExport Then
    InExport = True
    Call ClearExports
    If ReportAsString Then exportdest = "a String"
    If Not ExportAvailableEx(ExportType) Then Call Err.Raise(ERR_EXPORT, "ExportReportEx", "Unable to export as the export type " & ExportTypeStrEx(ExportType))
    rTarget = ReportControl.rTarget
    
    
    Opened = OpenExports(exportdest, ExportType, Overwrite, ReportAsString)
    If Not Opened Then Call Err.Raise(ERR_EXPORT, "ExportReportEx", "Unable to export as the file " & exportdest & " could not be opened")
    
    ReportControl.InReport = True
    ReportControl.CurReport = 1
    CurPage = ReportControl.CurPage
              
    If (ExportType = EXPORT_PDF) Then
       If (Not IsPDFAvailable) Then
          Call Err.Raise(380, "ExportReportEx", "PDF Printer driver is not available")
       End If
       lastPrinter = SetActivePrinter(m_PDFPrinterName)
       
       ReportControl.rTarget = RPT_PRINTER
      
       Call PreviewPrintPageEx(1, ReportControl.Pages_N, exportdest)
       InExport = False
       ExportReportEx = True
       GoTo ExportReportEx_end
    Else
      ReportControl.rTarget = ExportType
    End If
              
    If Not IsExportHTML And (ExportType <> EXPORT_PDF) Then
      Call SetHeaderFooter("", REPORT_HEADER)
      Call SetHeaderFooter("", REPORT_FOOTER)
      Call SetHeaderFooter("", PAGE_HEADER)
      Call SetHeaderFooter("", PAGE_FOOTER)
      Call bOut(ReportControl.ExportHeader & "{EOLN}")
    End If
    
    For ReportControl.CurPage = 1 To ReportControl.Pages_N
      Call bOut(Pages(ReportControl.CurPage).data)
      Call bOut(Pages(ReportControl.CurPage).ExportOnlyFooter)
      If IsExportHTML Then
        Call SetNewHTMLPage
        If (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then
          ReportControl.HTML.CurrentY = ReportControl.HTML.CurrentY + 200
        End If
      End If
    Next ReportControl.CurPage
    
    If IsExportHTML Then
      Call ConstructHTMLFooter
      If Not ReportAsString Then
        Print #m_File, ReportControl.HTML.HTMLString;
      Else
        exportdest = ReportControl.HTML.HTMLString
      End If
    End If
    InExport = False
    ExportReportEx = True
  End If
  
ExportReportEx_end:
  If (Len(lastPrinter) > 0) Then
    
    Call SetActivePrinter(lastPrinter)
  End If
  If Opened Then
    ReportControl.InReport = False
    ReportControl.CurPage = CurPage
    ReportControl.rTarget = rTarget
    Call ClearExports
  End If
  Call ClearCursor
  ReportControl.CurReport = 0
  Exit Function
  
ExportReportEx_err:
  Call ErrorMessage(ERR_ERROR, Err, "ExportReportEx", "Reporter Export", "Error exporting report to " & exportdest)
  InExport = False
  ExportReportEx = False
  Resume ExportReportEx_end
  Resume
End Function




