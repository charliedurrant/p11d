Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Sub Main()
  Dim args As String
  Dim logCsv As String
  Dim fr As TCSFileread
  
On Error GoTo err_err
  
  args = Command$
  If Len(args) = 0 Then
    Call MyExitProcess(1)
    GoTo err_end
  End If
  
  Set fr = New TCSFileread
  Call fr.OpenFile(args)
  Call fr.GetFile(logCsv)
  Set fr = Nothing
  
  Call PrintPdfFromReportLog(logCsv)
  
err_end:
  Call MyExitProcess(0)
  Exit Sub
err_err:
  Call MyExitProcess(1)
  Exit Sub
  Resume
End Sub
Private Sub MyExitProcess(exitCode As Long)
  If IsRunningInIDE Then Exit Sub
  Call ExitProcess(exitCode)
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
  
  Set csvParser = New csvParser
  logCsvPrepared = PrepareCsv(logCsv)
  Call csvParser.Init(logCsvPrepared, False)
  Set rep = New Reporter
  
  For rowIndex = 0 To csvParser.RowCount - 1
    notificationType = CLng(csvParser.ValueByIndex(rowIndex, 0))
    value = csvParser.ValueByIndex(rowIndex, 1)
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
End Sub
