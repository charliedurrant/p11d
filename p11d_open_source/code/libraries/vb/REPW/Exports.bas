Attribute VB_Name = "Exports"
Option Explicit

Public Enum EXPORT_OUT
  OUT_TEXT
  OUT_CR
  OUT_COMMA
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
Private m_ColCount As Long


Private m_Word As Word.Application
Private m_WordDoc As Word.Document
Private m_Rng As Word.Range

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
    Case Else
      ExportTypeExtEx = "Unknown Export"
  End Select
End Function

Public Sub ExportOut(ByVal OutCtrl As EXPORT_OUT, ByVal String1 As String)
  Dim TextWidth As Single
  
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
        If IsNumeric(m_CellValues(c1)) Then
          m_CellValues(c1) = CDbl(m_CellValues(c1))
        End If
        c1 = c1 + 1
        If c1 > m_ColCount Then
          m_ColCount = m_ColCount + 1
          ReDim Preserve m_CellValues(1 To m_ColCount) As Variant
        End If
        m_CellValues(c1) = ""
      ElseIf OutCtrl = OUT_CR Then
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
  Select Case ExportType
    Case EXPORT_CSV, EXPORT_FIXEDWIDTH
           ExportAvailableEx = True
    Case EXPORT_EXCEL
           ExportAvailableEx = isCOMPresent("Excel.Application", WIN32_SERVERPROC)
    Case EXPORT_WORD
           'ExportAvailableEx = False
           ExportAvailableEx = isCOMPresent("Word.Application", WIN32_SERVERPROC)
    Case EXPORT_HTML_IE, EXPORT_HTML_NETSCAPE, EXPORT_HTML_INTEXP5  'km
           ExportAvailableEx = True
    Case Else
           Call ECASE("ExportAvailableEx - unknown export type")
  End Select
End Function

Public Function ExportReportEx(exportdest As String, ByVal ExportType As REPORT_EXPORTS, ByVal Overwrite As Boolean, ByVal ReportAsString As Boolean) As Boolean
  Static InExport As Boolean
  Dim rTarget As REPORT_TARGET, CurPage As Long, Opened As Boolean
  
  On Error GoTo ExportReportEx_err
  Opened = False
  Call SetCursor
  If Not InExport Then
    InExport = True
    Call ClearExports
    If ReportAsString Then exportdest = "a String"
    If Not ExportAvailableEx(ExportType) Then Call Err.Raise(ERR_EXPORT, "ExportReportEx", "Unable to export as the export type " & ExportTypeStrEx(ExportType))
    rTarget = ReportControl.rTarget
    ReportControl.rTarget = ExportType
    Opened = OpenExports(exportdest, ExportType, Overwrite, ReportAsString)
    If Not Opened Then Call Err.Raise(ERR_EXPORT, "ExportReportEx", "Unable to export as the file " & exportdest & " could not be opened")
    
    ReportControl.InReport = True
    ReportControl.CurReport = 1
    CurPage = ReportControl.CurPage
              
    If Not IsExportHTML Then
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



