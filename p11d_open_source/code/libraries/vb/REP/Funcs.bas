Attribute VB_Name = "Functions"
Option Explicit
Public Enum HEADERFOOTER
  RECALC_ONLY = 0
  REPORT_HEADER
  REPORT_FOOTER
  PAGE_HEADER
  PAGE_FOOTER
  EXPORT_HEADER
End Enum



Public Sub SetOrient(ByVal pmode As REPORT_ORIENTATION, ByVal rTarget As PRIVATE_REPORT_TARGET)
  On Error GoTo SetOrient_err:
  Call xSet("SetOrient")
  'apf2008 altered
  If rTarget <> RPT_PREPARE Then
      If Printer.Orientation <> pmode Then
        Printer.Orientation = pmode
      End If
  End If
  
  
  ReportControl.Orientation = pmode
  
SetOrient_end:
  Call xReturn("SetOrient")
  Exit Sub
  
SetOrient_err:
  ReportControl.Orientation = pmode
  If (rTarget = RPT_PRINTER) Or (rTarget = RPT_PREVIEW_PRINT) Then
    Call ErrorMessage(ERR_INFO, Err, "SetOrient", "Unable to Set Orientation", "Unable to retrieve/set the orientation on the Printer")
  End If
  Resume SetOrient_end
End Sub

Public Function SetHeaderFooter(NewValue As String, hf As HEADERFOOTER, Optional ByVal GetValue As Boolean = False, Optional ByVal Force As Boolean = False) As String
  On Error GoTo SetHeaderFooter_err
  Call xSet("SetHeaderFooter")
  If (ReportControl.CurReport > 0) And (ReportControl.InReport Or Force) Then
    Select Case hf
      Case REPORT_HEADER
            If GetValue Then
              SetHeaderFooter = rpt(ReportControl.CurReport).RHeader
            Else
              rpt(ReportControl.CurReport).RHeader = NewValue
            End If
      Case REPORT_FOOTER
            If GetValue Then
              SetHeaderFooter = rpt(ReportControl.CurReport).RFooter
            Else
              rpt(ReportControl.CurReport).RFooter = NewValue
              If Len(NewValue) > 0 Then
                Call CalcFooterHeights
              Else
                rpt(ReportControl.CurReport).RFooterH = 0
              End If
            End If
      Case PAGE_HEADER
            If GetValue Then
              SetHeaderFooter = rpt(ReportControl.CurReport).PHeader
            Else
              If Force Or Not rpt(ReportControl.CurReport).OverrideHeader Then
                rpt(ReportControl.CurReport).PHeader = NewValue
              End If
            End If
      Case PAGE_FOOTER
            If GetValue Then
              SetHeaderFooter = rpt(ReportControl.CurReport).PFooter
            Else
              If Force Or Not rpt(ReportControl.CurReport).OverrideFooter Then
                rpt(ReportControl.CurReport).PFooter = NewValue
                If Len(NewValue) > 0 Then
                  Call CalcFooterHeights
                Else
                  rpt(ReportControl.CurReport).PFooterH = 0
                End If
              End If
            End If
      Case EXPORT_HEADER
            If GetValue Then
              SetHeaderFooter = ReportControl.ExportHeader
            Else
              ReportControl.ExportHeader = NewValue
            End If
      Case RECALC_ONLY
            Call CalcFooterHeights
      Case Else
            Err.Raise ERR_SETHEADERFOOTER, "SetHeaderFooter", "Unknown HEADERFOOTER type"
    End Select
  End If
SetHeaderFooter_end:
  Call xReturn("SetHeaderFooter")
  Exit Function
  
SetHeaderFooter_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetHeaderFooter", "Setting Header/Footer on Report", "Unable to set the Page or Report header or footer")
  Resume SetHeaderFooter_end:
End Function

Public Function bOut(ByVal s As String) As Boolean
  Dim tmp As String, bNewPage As Boolean, bInPrintSection As Boolean
    
output_section:
  If Len(s) = 0 Then Exit Function
  Call ReportParser.ParseLine(s)
  tmp = ReportParser.PostParseLine
  bOut = DisplayText(tmp)
  If ReportControl.rTarget = RPT_PREPARE Then
    If Not ReportControl.NoRecord Then
      If ReportControl.BeginSectionY <> -1 Then
        ReportControl.Section = ReportControl.Section & s
      Else
        ' if the out statement is more than one page then the complete out statement gets recorded
        ' on the next page - and so on
        ' fix apf 29/4 (not p11d)
        Pages(ReportControl.CurPage).data.Append s
        If ReportControl.DelimitOut Then Pages(ReportControl.CurPage).data.Append "{Z}"
      End If
    End If
    ReportControl.DelimitOut = False
  ElseIf ReportControl.rTarget = RPT_PRINTER Then
    If Not ReportControl.NoRecord Then
      If ReportControl.BeginSectionY <> -1 Then
        ReportControl.Section = ReportControl.Section & s
      End If
    End If
  End If
  ReportControl.NoRecord = False
  bOut = ReportControl.OnNewPage Or bNewPage
  If ReportControl.OutputSection Or ((ReportControl.BeginSectionY <> -1) And ReportControl.OnNewPage) Then
    ReportControl.BeginSectionX = -1!
    ReportControl.BeginSectionY = -1!
    ReportControl.OutputSection = False
    ReportControl.OnNewPage = False
    bInPrintSection = True
    bNewPage = bOut
    s = ReportControl.Section
    ReportControl.Section = ""
    GoTo output_section
  End If
  ReportControl.OnNewPage = False
End Function

Public Function IsPrinterAvail(ByVal CheckPrinter As Boolean) As Boolean
  Static RetVal As Boolean
  Dim s As String
  
  If Not CheckPrinter Then GoTo IsPrinterAvail_end
  
  On Error GoTo IsPrinterAvail_err
  'apf2008 does not detect printer being offline
  s = Printer.DeviceName
  RetVal = True

IsPrinterAvail_end:
  IsPrinterAvail = RetVal
  Exit Function

IsPrinterAvail_err:
  RetVal = False
  Resume IsPrinterAvail_end
End Function

Private Sub CalcFooterHeights()
  Dim curtarget As REPORT_TARGET
  Dim pic As PictureBox, OldCurPageValid As Boolean
    
  Call xSet("CalcFooterHeights")
  curtarget = ReportControl.rTarget
  ReportControl.rTarget = RPT_CONFIG
  OldCurPageValid = ReportControl.CurPageValid
  Set pic = ReportControl.Preview
  Set ReportControl.Preview = ReportControl.PreviewTest
  ReportControl.Preview.Cls
  
  If Len(rpt(ReportControl.CurReport).PFooter) > 0 Then
    Call bOut("{Arial=10,N}{STARTSKIPEXPORT}" & rpt(ReportControl.CurReport).PFooter & vbCrLf & "{ENDSKIPEXPORT}")
    rpt(ReportControl.CurReport).PFooterH = ReportControl.Preview.CurrentY + 144
    ReportControl.Preview.Cls
  Else
    rpt(ReportControl.CurReport).PFooterH = 0
  End If
    
  If Len(rpt(ReportControl.CurReport).PFooter) > 0 Then
    Call bOut("{Arial=10,N}" & rpt(ReportControl.CurReport).RFooter & vbCrLf)
    rpt(ReportControl.CurReport).RFooterH = ReportControl.Preview.CurrentY + 144
  Else
    rpt(ReportControl.CurReport).RFooterH = 0
  End If
  Set ReportControl.Preview = pic
  ReportControl.CurPageValid = OldCurPageValid
  ReportControl.rTarget = curtarget
  Call xReturn("CalcFooterHeights")
End Sub

Public Sub InitNewPage()
  Dim i As Long, yabs As Single
  
  If (Not ReportControl.InReport) Or ((ReportControl.rTarget > RPT_EXPORT) And (Not IsExportHTML)) Or ReportControl.AbortReport Then Exit Sub
  Call xSet("InitNewPage")
  ReportControl.FirstX = 0
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    If ReportControl.CurPage > 0 Then
      If ReportControl.rTarget = RPT_PRINTER Then
        Printer.CurrentY = ReportControl.PageHeight - rpt(ReportControl.CurReport).PFooterH
        ReportControl.SuppressNewPageCalc = True
        Printer.CurrentX = ReportControl.LeftMargin + ReportControl.FirstX
        If Len(rpt(ReportControl.CurReport).PFooter) > 0 Then Call bOut("{Arial=10,N}{STARTSKIPEXPORT}" & rpt(ReportControl.CurReport).PFooter & "{ENDSKIPEXPORT}")
        ReportControl.SuppressNewPageCalc = False
        ReportControl.PrinterNewPage = True
        ReportControl.Pages_N = ReportControl.Pages_N + 1
      End If
    Else
      'EAW portrait.. goes back to portrait
      Printer.Print "";
      Printer.CurrentX = ReportControl.LeftMargin
      Printer.CurrentY = 1
    End If
  ElseIf ReportControl.rTarget = RPT_PREPARE Then
    If ReportControl.CurPage > 0 Then
      yabs = ReportControl.PageHeight - rpt(ReportControl.CurReport).PFooterH
      ReportControl.SuppressNewPageCalc = True
      If Len(rpt(ReportControl.CurReport).PFooter) > 0 Then Call bOut("{YABS=" & CStr(yabs) & "}" & "{Arial=10,N}" & "{STARTSKIPEXPORT}" & rpt(ReportControl.CurReport).PFooter & "{ENDSKIPEXPORT}")
      ReportControl.SuppressNewPageCalc = False
      Pages(ReportControl.CurPage).complete = True
      ReportControl.Pages_N = ReportControl.Pages_N + 1
    End If
    ReportControl.Preview.CurrentX = ReportControl.LeftMargin
    ReportControl.Preview.CurrentY = 1
  End If
  
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREPARE) Then
    ReportControl.CurPageValid = False
    ReportControl.PageHeaderPrinted = False


    ReportControl.CurPage = ReportControl.CurPage + 1
    Call AddPage
    Pages(ReportControl.CurPage).complete = False
    If ReportControl.CurPage = 1 Then
      Pages(ReportControl.CurPage).PageNumber = 1
      Pages(ReportControl.CurPage).PrePageNumber = ""
      Pages(ReportControl.CurPage).PostPageNumber = ""
    Else
      Pages(ReportControl.CurPage).PageNumber = Pages(ReportControl.CurPage - 1).PageNumber + 1
      Pages(ReportControl.CurPage).PrePageNumber = Pages(ReportControl.CurPage - 1).PrePageNumber
      Pages(ReportControl.CurPage).PostPageNumber = Pages(ReportControl.CurPage - 1).PostPageNumber
    End If
    For i = 0 To (REPORT_CONSTANTS_N - 1)
      If PageStatics(i) Then
        Pages(ReportControl.CurPage).statics(i) = PageStaticsDefault(i)
      Else
        If ReportControl.CurPage > 1 Then
          Pages(ReportControl.CurPage).statics(i) = Pages(ReportControl.CurPage - 1).statics(i)
        Else
          Pages(ReportControl.CurPage).statics(i) = ""
        End If
      End If
    Next i
    If ReportControl.rTarget = RPT_PREPARE Then Call bOut("{LEFTMARGINABS=" & ReportControl.LeftMargin & "}")
  End If
  ReportControl.FColor.ColorSet = False
  ReportControl.fStyle.FontType = VALID_FONT_TYPE
    
InitNewPage_exit:
  Call xReturn("InitNewPage")
End Sub


Public Sub PreviewPageEx(Page As Long, picPaper As PictureBox)
  Static InPreview As Boolean
  
  On Error GoTo PreviewPageEx_err
  
  If Not InPreview Then
    InPreview = True
    If Page < 0 Then Beep: GoTo PreviewPageEx_end
    If Page > ReportControl.Pages_N Then Beep: GoTo PreviewPageEx_end
    Set ReportControl.Preview = picPaper
    ReportControl.fStyle.FontType = VALID_FONT_TYPE
    ReportControl.FColor.ColorSet = False
    ReportControl.CurPage = Page
    ReportControl.FirstX = 0
    ReportControl.Preview.Cls
    ReportControl.Preview.CurrentX = ReportControl.LeftMargin
    ReportControl.Preview.CurrentY = 1
    
    ReportControl.SuppressNewPageCalc = True
    Call bOut(Pages(ReportControl.CurPage).data)
    ReportControl.SuppressNewPageCalc = False
  End If
  
PreviewPageEx_end:
  InPreview = False
  Exit Sub
  
PreviewPageEx_err:
  Resume PreviewPageEx_end
End Sub

Public Sub PreviewPrintPageEx(ByVal StartPage As Long, ByVal EndPage As Long, Optional exportName As String = "")
  Static InPreviewPrint As Boolean
  Dim rTarget As REPORT_TARGET
  Dim CurPage As Long
  Dim pdfInstalled As Boolean
  Dim pdffileNameOptionsEx As Long
  Dim pdfdefaultFileName As String
  
  
  On Error GoTo PreviewPrintPageEx_err
  Call xSet("PreviewPrintPageEx")
  Call SetCursor
  
  If Not InPreviewPrint Then
    
    If PDFDriverInstall Then
      pdffileNameOptionsEx = g_cdi.FileNameOptionsEx
      pdfdefaultFileName = g_cdi.DefaultFileName
      g_cdi.FileNameOptionsEx = 1 + 2 ' NoPrompt + UseFileName
      g_cdi.DefaultFileName = exportName
    Else
      
    End If
    InPreviewPrint = True
    If Not IsPrinterAvail(True) Then Err.Raise ERR_NOPRINTER, "PreviewPrintPageEx", "No printer defined or current printer invalid"
    CurPage = ReportControl.CurPage
    ReportControl.SuppressNewPageCalc = True
    Call SetHeaderFooter("", REPORT_HEADER)
    Call SetHeaderFooter("", REPORT_FOOTER)
    Call SetHeaderFooter("", PAGE_HEADER)
    Call SetHeaderFooter("", PAGE_FOOTER)
    rTarget = ReportControl.rTarget
    ReportControl.rTarget = RPT_PREVIEW_PRINT
    Call SetOrient(ReportControl.Orientation, ReportControl.rTarget)
    
    ReportControl.CurPage = 0
    
    Call InitNewPage
    For ReportControl.CurPage = 1 To ReportControl.Pages_N
      If (ReportControl.CurPage >= StartPage) And (ReportControl.CurPage <= EndPage) Then
        Call bOut(Pages(ReportControl.CurPage).data)
        Call InitNewPage
        If Not (((StartPage = ReportControl.CurPage) And (StartPage = EndPage)) Or _
                (EndPage = ReportControl.CurPage)) Then
          Printer.NewPage
          Printer.CurrentX = ReportControl.LeftMargin
          Printer.CurrentY = 1
        End If
      End If
    Next ReportControl.CurPage
    
    Printer.EndDoc 'EWPDF
    m_LastReportControl = ReportControl
    
    ReportControl.SuppressNewPageCalc = False
    ReportControl.CurPage = CurPage
    ReportControl.rTarget = rTarget
    InPreviewPrint = False
  End If
  
PreviewPrintPageEx_end:
  If (Not g_cdi Is Nothing) Then
    g_cdi.FileNameOptionsEx = pdffileNameOptionsEx
    g_cdi.DefaultFileName = pdfdefaultFileName
    Call PDFDriverUninstall
  End If
  
  
  Call xReturn("PreviewPrintPageEx")
  Call ClearCursor
  Exit Sub
  
PreviewPrintPageEx_err:
  InPreviewPrint = False
  Call ErrorMessage(ERR_ERROR, Err, "PreviewPrintPageEx", "Print Report", "Error Printing Report")
  Resume PreviewPrintPageEx_end
  Resume
End Sub

Public Sub A4CheckPrinterSize()
  On Error GoTo err_Err
  
  If (Not A4Force) Then GoTo err_End
  If Printer.PaperSize <> vbPRPSA4 Then
    Call Err.Raise(ERR_A4, "CheckPrinterSize", "The printer paper seems not to be set to A4. Some reports may not print as expected." & vbCrLf & vbCrLf & "If you do not want to get this error either change the printer's paper to A4 or check the box below to not see this error again.")
  End If
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ALLOWIGNORE Or ERR_ERROR, Err, "CheckPrinterSize", "Paper not set to A4", "The paper not correct")
  Resume err_End
End Sub

