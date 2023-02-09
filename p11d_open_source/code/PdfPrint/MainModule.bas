Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private m_consoleOutput As Boolean

Public Sub Main()
  Dim args As String
  Dim logCsv As String
  Dim fr As TCSFileread
  
'EVERY TIME THIS IS BUILD DRAG THE exe over the LinkConsole.vbs

On Error GoTo err_err
  
  
  m_consoleOutput = Not IsRunningInIDE
  
   ' Required in all MConsole.bas supported apps!
  
  If (m_consoleOutput) Then Con.Initialize

  args = Command$
  If Len(args) = 0 Then
    Call ConsoleWriteLine("Exiting app as no command line file set with the reporter output")
    Call MyExitProcess(1)
    GoTo err_end
  End If
  
  ConsoleWriteLine ("Reporter file: " & args)
    
  ConsoleWriteLine ("Opening file: " & args)
  
  logCsv = TextFileLoad(args)
  
  ConsoleWriteLine ("Opened file: " & args)
  
  Call PrintPdfFromReportLog(logCsv)
  
  ConsoleWriteLine ("Finished")
  
err_end:
  Call MyExitProcess(0)
  Exit Sub
err_err:
  ConsoleWriteLine ("Error: " & Err.Description)
  Call MyExitProcess(1)
  Exit Sub
  Resume
End Sub
Private Sub ConsoleWriteLine(message As String)
  If Not m_consoleOutput Then Exit Sub
  ConsoleWriteLine (message)
End Sub
Public Function TextFileLoad(ByVal sPathAndFile As String, Optional ByVal charset As String = "utf-8") As String
  Dim objStream As Stream
  Dim strData As String
  
On Error GoTo err_err

  If (Not FileExists(sPathAndFile)) Then
     Call Err.Raise(1, "TextFileLoad", "Failed load the text file " & sPathAndFile & " as it does not exist")
  End If
  
  Set objStream = New Stream
  objStream.charset = charset
  Call objStream.Open
  Call objStream.LoadFromFile(sPathAndFile)
  strData = objStream.ReadText()
  Call objStream.Close
  TextFileLoad = strData
  
err_end:
  Set objStream = Nothing
  Exit Function
err_err:
  Set objStream = Nothing
  Call Err.Raise(1, ErrorSource(Err, "TextFileLoad"), Err.Description)
  Resume
End Function

Private Sub MyExitProcess(ExitCode As Long)
  If IsRunningInIDE Then Exit Sub
  Call ExitProcess(ExitCode)
End Sub
Private Function PrepareCsv(ByRef logCsv As String) As String
  Dim ret As String
  Dim p0 As Long
  Dim pStart As Long
  Dim continue As Boolean
  Dim lenCrLfLess1 As Long
  Dim foundCount As Long
  
  continue = True
  pStart = Len(logCsv)
  lenCrLfLess1 = Len(vbCrLf) - 1
  foundCount = 0
  
  While continue
    p0 = InStrRev(logCsv, vbCrLf, pStart)
    If (p0 > 0 And p0 = pStart - lenCrLfLess1) Then
      foundCount = foundCount + 1
      pStart = p0 - 1
    Else
      continue = False
    End If
  Wend
  
  If (foundCount > 1) Then
    ret = Left$(logCsv, Len(logCsv) - ((foundCount - 1) * Len(vbCrLf)))
  ElseIf foundCount = 1 Then
    ret = logCsv
  Else
    ret = logCsv & vbCrLf
  End If
  
  PrepareCsv = ret
End Function
Private Sub PrintPdfFromReportLog(logCsv As String)
  Dim csvParser As csvParser
  Dim rowIndex As Long
  Dim typeString As String, value As String
  Dim notificationType As REPORTER_NOTIFICATON_TYPE
  Dim rep As Reporter
  Dim logCsvPrepared As String
  Dim notificationTypeString As String
  
On Error GoTo err_err
  
  ConsoleWriteLine ("Parsing file as a Csv")
  
  Set csvParser = New csvParser
  ConsoleWriteLine ("Prepare file for Csv parsing")
  ConsoleWriteLine ("File len: " & Len(logCsv))
  logCsvPrepared = PrepareCsv(logCsv)
  Call csvParser.Init(logCsvPrepared, False)
  ConsoleWriteLine ("Initialised Csv parser")
  Set rep = New Reporter
  ConsoleWriteLine ("Created reporter object")
  
  For rowIndex = 0 To csvParser.RowCount - 1
    ConsoleWriteLine ("Parse line: " & CStr((rowIndex + 1)))
    ConsoleWriteLine ("1")
    notificationTypeString = csvParser.ValueByIndex(rowIndex, 0)
    ConsoleWriteLine ("2:" & notificationTypeString)
    notificationType = CLng(notificationTypeString)
    ConsoleWriteLine ("3")
    value = csvParser.ValueByIndex(rowIndex, 1)
    ConsoleWriteLine ("Blah")
    Call ProcessReporterLine(notificationTypeString, value)
    ConsoleWriteLine ("Blah2")
    Select Case notificationType
      Case REPORTER_NOTIFICATON_TYPE.A4_FORCE
        rep.A4Force = CBoolean(value)
      Case REPORTER_NOTIFICATON_TYPE.END_REPORT
        Call rep.EndReport
      Case REPORTER_NOTIFICATON_TYPE.EXPORT_HEADER_SET
        rep.ExportHeader = value
      Case REPORTER_NOTIFICATON_TYPE.EXPORT_REPORT
        Call rep.ExportReport(value, EXPORT_PDF, True)
      Case REPORTER_NOTIFICATON_TYPE.FOOTER_SET
        rep.ReportFooter = value
      Case REPORTER_NOTIFICATON_TYPE.HEADER_SET
        rep.ReportHeader = value
      Case REPORTER_NOTIFICATON_TYPE.INIT_REPORT
        Call rep.InitReport("Report", PREPARE_REPORT)
      Case REPORTER_NOTIFICATON_TYPE.OUT_CALLED
        Call rep.Out(value)
      Case REPORTER_NOTIFICATON_TYPE.PAGE_FOOTER_SET
        rep.PageFooter = value
      Case REPORTER_NOTIFICATON_TYPE.PAGE_HEADER_SET
        rep.PageHeader = value
    End Select
  Next
  
  Exit Sub
err_err:
  Call Err.Raise(1, Err.Source + ":PrintPdfFromReportLog", Err.Description)
End Sub
Private Sub ProcessReporterLine(commandType As String, value As String)
  ConsoleWriteLine ("Command=" & commandType & ", value=" & value)
End Sub

