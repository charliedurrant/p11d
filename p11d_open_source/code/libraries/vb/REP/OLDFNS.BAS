Attribute VB_Name = "OldFunctions"
Option Explicit

'obsolete
'Private Function PrintTextPrep(s As String) As Boolean
'  Static inPrintTextPrep As Boolean
'  Dim fh As Single, cy As Single
'  Dim s0 As String, s1 As String, p As Long
'  Rem checks if off the bottom and enters new page
'
'  Call xSet("PrintTextPrep")
'  PrintTextPrep = False
'  If ReportControl.SuppressNewPageCalc Then GoTo PrintTextPrep_end
'  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREPARE) Then
'    If inPrintTextPrep Then Call Err.Raise(ERR_PAGERECURSIVE, "PrintTextPrep ", "Attempted to print another new page while processing a new page, reports cannot be recursive")
'    inPrintTextPrep = True
'
'    fh = GetStringHeight(s)
'    If ReportControl.rTarget = RPT_PRINTER Then
'      cy = Printer.CurrentY
'    Else
'      cy = ReportControl.Preview.CurrentY
'    End If
'    If ReportControl.PageHeight < (cy + fh + rpt(ReportControl.CurReport).PFooterH) Then
'      ' print Text before return char before PageBreak - rest After
'      If ReportControl.WrapX > 0! Then
'        p = 0
'      Else
'        p = InStr(s, vbCrLf)
'      End If
'      If p > 0 Then
'        s0 = Left$(s, p - 1)
'        If Len(s0) > 0 Then Call PrintText(s0)
'      End If
'      Call PushFontStyle
'      Call PushColor
'      Call InitNewPage
'      Call PopColor
'      Call PopFontStyle
'      ReportControl.SuppressNewPageCalc = True
'      Call PreOut
'      ReportControl.SuppressNewPageCalc = False
'      ReportControl.OnNewPage = True
'      If ReportControl.WrapX > 0! Then
'        p = 0
'      Else
'        p = InStr(s, vbCrLf)
'      End If
'      If p > 0 Then
'        s1 = Right$(s, Len(s) - (p + 1))
'      Else
'        s1 = s
'      End If
'      'If NotInStr(s1, vbCrLf & " ") <> 0 Then Call PrintText(s1)
'      If Len(s1) > 0 Then Call PrintText(s1)
'      PrintTextPrep = True
'    End If
'    inPrintTextPrep = False
'  End If
'
'PrintTextPrep_end:
'  Call xReturn("PrintTextPrep")
'End Function
'
''obsolete
'Private Function GetStringHeight(s As String) As Single
'  Dim WrapX As Single, tmp0 As String
'  Dim fh As Single, cfh  As Single, pos As Integer
'
'  Call xSet("GetStringHeight")
'  fh = ReportControl.fStyle.FontHeight
'  cfh = fh: pos = 1: WrapX = ReportControl.WrapX
'  Do While pos <> 0
'    pos = InStr(pos, s, vbCrLf)
'    If WrapX > 0! Then
'      If pos = 0 Then
'        tmp0 = Mid$(s, 1)
'      Else
'        tmp0 = Mid$(s, 1, pos - 1)
'        cfh = cfh + fh
'        pos = pos + 2
'      End If
'      cfh = cfh + (fh * CountWrapToWidth(tmp0, WrapX))
'      WrapX = 0!
'    ElseIf pos <> 0 Then
'      cfh = cfh + fh
'      pos = pos + 2
'    End If
'  Loop
'  GetStringHeight = cfh
'  Call xReturn("GetStringHeight")
'End Function
'
'
