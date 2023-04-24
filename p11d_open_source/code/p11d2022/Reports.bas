Attribute VB_Name = "Reports"
Option Explicit

'Rita patel for form validation
' rita.patel@hmrc.gsi.gov.uk
' tel 0207 438 4264

Private Const S_LIGHTGREY As String = "12632256"
Private Const S_P11DB_SHORT_CAPTION_BOX_A As String = "total benefits liable to Class 1A NICs from forms P11D"
Public Const S_P11DB_ADDITIONS_BOX_B As String = "Add any amounts not included in box A on which Class 1A NICs are due"
Public Const S_P11DB_DEDUCTIONS_BOX_C As String = "Deduct any amounts included in box A on which Class 1A NICs are not due"
Private Const S_P11DB_TOTAL_BOX_D As String = "Total of benefits on which Class 1A NICs are due"
Private Const S_P11DB_MUTIPLY_BY_NIC_RATE_BOX_E As String = "Multiply by Class 1A NICs rate"
Private Const S_P11DB_CLASS_1A_PAYABLE_BOX_F As String = "Class 1A NICs payable"
Private Const S_P11DB_AMOUNTS_TAXED_THROUGH_PAYROLL As String = "Amounts taxed through payroll"
Private Const S_P11DB_AMOUNTS_INFORMALLY_TAXED_THROUGH_PAYROLL As String = "informally taxed through payroll"
Private Const S_P11DB_EMPLOYEES_NOT_SUBJECT_TO_CLASS_1A As String = "Employees not subject to Class 1A"


Private Const L_HMIT_COL_1 As Long = 50
Private Const L_HMIT_COL_2 As Long = 63
Private Const L_HMIT_COL_3 As Long = 78
Private Const S_P11D_B = "P11D(b)"
Private Const L_HMIT_COL_4 As Long = 82
Private Const L_HMIT_COL_5 As Long = 94
Private Const L_HMIT_STANDARDBOX_WIDTH = 12
'KA - Private Const L_HMIT_STANDARDBOX_HEIGHT = 2
Private Const L_HMIT_STANDARDBOX_HEIGHT = 1.4
Private Const L_HMIT_SIGNATORYBOX_HEIGHT = 4 'EK Added 25/02/2003

Public Enum EL_MENU_CAPTIONS
  [_ELMC_FIRST_ITEM]
  ELMC_FORMAT = [_ELMC_FIRST_ITEM]
  ELMC_EMPLOYER
  ELMC_EMPLOYEE
  
  ELMC_DATES
  
  ELMC_SUB_REPORTS
  [_ELMC_LAST_ITEM] = ELMC_SUB_REPORTS
End Enum

Public Enum P46_PAYMENT_FREQUENCY
  P46PF_ANNUALLY
  P46PF_QUARTERLY
  P46PF_MONTHLY
  P46PF_WEEKLY
  P46PF_ACTUAL
End Enum

Private Const L_P46_BACKGROUND_COL_WIDTH As Long = 47
Private Const L_P46_COL_1_X As Long = 2
Private Const L_P46_COL_2_X As Long = 51
  

'******************* WK values **************************************
Public Enum WKOUT_TYPE
  WK_BLANK_LINE = 1
  WK_SECTION_HEADER
  WK_SECTION_HEADER_DETAILS
  WK_SECTION_HEADER_VALUE
  WK_SECTION_HEADER_LESS
  WK_SECTION_HEADER_BENEFIT
  WK_SECTION_BREAK
  WK_ITEM_TEXT
  WK_ITEM_DESCRIPTION
  WK_ITEM_TEXT_BOLD
  WK_ITEM_NOTE
  WK_ITEM_subtotal
  WK_ITEM_Total
  WK_SUBTOTAL_ONLY
  WK_TOTAL_ONLY
End Enum

Private Enum ELC_FONT
  ELC_BODY_NORMAL = 1
  ELC_BODY_BOLD
  ELC_TABLE_NORMAL
  ELC_TABLE_NORMAL_RIGHT
  ELC_TABLE_HEADING
  ELC_TABLE_HEADING_RIGHT
  
End Enum


Private Enum EL_TABLE_COL
  [_ELT_FIRST_ITEM] = 1
  ELT_COL1 = [_ELT_FIRST_ITEM]
  ELT_COL2
  ELT_COL3
  ELT_COL4
  [_ELT_LAST_ITEM] = ELT_COL4
End Enum

Private Enum EL_TABLE_DATA
  ELT_SPACING_REP = 1
  ELT_SPACING_EMAIL_TABLE_LINE
  ELT_SPACING_EMAIL_HEADER_LINE
  ELT_COLUMN_HEADER
End Enum


Public Const S_WK_NORMAL_FONT As String = "{Arial=7,n}"
Public Const S_WK_NORMAL_BOLD_FONT As String = "{Arial=7,bn}"
Public Const S_WK_NORMAL_ITALIC_FONT As String = "{Arial=7,i}"
Public Const S_WK_RIGHT_BOLD_FONT As String = "{Arial=7,rb}"
Public Const S_WK_LINE_BREAK_FONT As String = "{Arial=6,n}"

Public Const S_ELMC_MASTER As String = "Control Codes"

Public Const S_WK_HEADER_FONT As String = "{Arial=7,b}"

Public Const L_WK_OTHER_TABLE_COL1 As Long = 5
Public Const L_WK_OTHER_TABLE_COL3 As Long = 87

Public Const L_WK_OTHER_TABLE_COL4 As Long = 97 'change this, change below
Public Const S_WKCOL_2 As String = "{x=97}"

Public Const L_WK_OTHER_TABLE_COL2 As Long = 77 'change this, change below
Public Const S_WKCOL_3 As String = "{x=77}"



'rdc
Public Const L_WK_OTHER_TABLE_COL_11 As Long = 5
Public Const L_WK_OTHER_TABLE_COL_12 As Long = 50
Public Const L_WK_OTHER_TABLE_COL_13 As Long = 60
Public Const L_WK_OTHER_TABLE_COL_14 As Long = 70
Public Const L_WK_OTHER_TABLE_COL_15 As Long = 80
Public Const L_WK_OTHER_TABLE_COL_16 As Long = 88
Public Const L_WK_OTHER_TABLE_COL_17 As Long = 97

'EK
Public Const L_WK_ACCOM_TABLE_COL1 As Long = 5
Public Const L_WK_ACCOM_TABLE_COL2 As Long = 47
Public Const L_WK_ACCOM_TABLE_COL3 As Long = 57
Public Const L_WK_ACCOM_TABLE_COL4 As Long = 77
Public Const L_WK_ACCOM_TABLE_COL5 As Long = 87
Public Const L_WK_ACCOM_TABLE_COL6 As Long = 97


Public Const S_WKCOL_1 As String = "{x=85}"
Public Const S_WK_LEFT_MARGIN As String = "{x=3}"
Public Const L_WK_HEADING_COL2 As Long = 65
Public Const L_WK_HEADING_COL1 As Long = 0


Public Const L_WK_BOXHEIGHT As Long = 2


'MS review to UDT
Private m_sWKColFormats() As String
Private m_lWKNumberOfColumns As Long
Private m_lWKXOffsets() As Long

'******************* End WK values **************************************

'************************* WK functions ******************************
Public Sub WKBenefitHeader(rep As Reporter, ben As IBenefitClass, bPrintDescription As Boolean)

  On Error GoTo WKBenefitHeader_ERR

  Call xSet("WKBenefitHeader")
  Call HMITSectionHeader(rep, p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION), p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_FORM_CAPTION) & ": " & IIf(bPrintDescription, ben.value(ITEM_DESC), ""))
  Call rep.Out(vbCrLf & vbCrLf)
  
WKBenefitHeader_END:
  Call xReturn("WKBenefitHeader")
  Exit Sub
WKBenefitHeader_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKBenefitHeader", "WK Header", "Error printing the WK benefit header.")
  Resume WKBenefitHeader_END
  Resume
End Sub

Public Sub ReportBanner(rep As Reporter, BannerText As String)
  If rep Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ReportBanner", "Reporter is nothing.")
  Call HMITBanner(rep, BannerText & " " & p11d32.Rates.value(TaxFormYear), True)
End Sub
Public Function WKMainHeader(rep As Reporter, ben As IBenefitClass, ee As Employee) As Boolean
   
  On Error GoTo WKMainHeader_ERR

  Call xSet("WKMainHeader")
  
  If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WkMainHeader", "The employee is nothing.")
  'Set ee = GetParentFromBenefit(ben, GPBF_EMPLOYEE)
  
  rep.Out ("{BEGINSECTION}")
  'JN to see wrap uu into another sub which just allows p11d expenses text to be changed
  Call ReportBanner(rep, "P11D EXPENSES AND BENEFITS")
  Call rep.Out(vbCrLf & vbCrLf)
  Call HMITEmpDetails(rep, ee)
  Call WKOut(rep, WK_SECTION_BREAK)
  rep.Out ("{ENDSECTION}")

  
  WKMainHeader = True
  
WKMainHeader_END:
  Call xReturn("WKMainHeader")
  Exit Function
WKMainHeader_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKMainHeader", "WK Header", "Error printing the WK header.")
  Resume WKMainHeader_END
End Function

Private Function WKOutputRef(ByVal OutPutRef As Variant) As String
  If Len(OutPutRef) > 0 Then
    WKOutputRef = S_WK_NORMAL_BOLD_FONT & S_WKCOL_1 & OutPutRef
  End If
End Function
Private Function WKTableItem(ByVal iCol As Long, ByVal sValue As String)
  
  WKTableItem = WKTableItem & "{Arial=7"
  If Len(m_sWKColFormats(iCol)) > 0 Then WKTableItem = WKTableItem & "," & m_sWKColFormats(iCol)
  WKTableItem = WKTableItem & "}{x=" & m_lWKXOffsets(iCol) & "}" & sValue
  
End Function
Public Function WKTableTotals(rep As Reporter, ParamArray TotalValues()) As Boolean
  Dim i As Long, s As String, j As Long

  On Error GoTo WKTableTotals_ERR

  Call xSet("WKTableTotals")

  If ((UBound(TotalValues) - LBound(TotalValues)) + 1) <> m_lWKNumberOfColumns Then Call Err.Raise(ERR_TOTALS_NOT_EQUAL_PARAMARRAY, "WKTableTotals", "Total values ubound of param array not equal to no of columns.")
    
  For i = 1 To 3
    For j = 1 To m_lWKNumberOfColumns
      Select Case i
        Case 1
          If Len(TotalValues(j - 1)) > 0 Then s = s & WKTableItem(j, "{line=-7}")
        Case 2
          s = s & WKTableItem(j, TotalValues(j - 1))
        Case 3
          If Len(TotalValues(j - 1)) > 0 Then s = s & WKTableItem(j, "{line=-7,d}")
      End Select
      
    Next
    s = s & vbCrLf
  Next
  
  Call rep.Out(s)
  
  WKTableTotals = True
  
WKTableTotals_END:
  Call xReturn("WKTableTotals")
  Exit Function
WKTableTotals_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKTableTotals", "WK Table Totals", "Error printing the WK table totals.")
  Resume WKTableTotals_END
  Resume
End Function
Public Sub WKTblOtherTypeTable(rep As Reporter, ByVal sCol1Caption As String, ByVal sCol2Caption As String, ByVal sCol3Caption As String, ByVal sCol4Caption As String)
  Call WKTblColXOffsets(L_WK_OTHER_TABLE_COL1, L_WK_OTHER_TABLE_COL2, L_WK_OTHER_TABLE_COL3, L_WK_OTHER_TABLE_COL4)
  Call WKTblColFormats("n", "rn", "rn", "rn")
  Call WKTableHeadings(rep, sCol1Caption, sCol2Caption, sCol3Caption, sCol4Caption)
End Sub

'CAD review 20/02 good
' rdc Differs to the above WKTblOtherTypeTable procedure to handle up to 7 columns/headings within the working-paper benefits report
Public Sub WKTblOtherTypeTableWithDates(rep As Reporter, ByVal sCol1Caption As String, ByVal sCol2Caption As String, ByVal sCol3Caption As String, ByVal sCol4Caption As String, ByVal sCol5Caption As String, ByVal sCol6Caption As String, ByVal sCol7Caption As String)
  Call WKTblColXOffsets(L_WK_OTHER_TABLE_COL_11, L_WK_OTHER_TABLE_COL_12, L_WK_OTHER_TABLE_COL_13, L_WK_OTHER_TABLE_COL_14, L_WK_OTHER_TABLE_COL_15, L_WK_OTHER_TABLE_COL_16, L_WK_OTHER_TABLE_COL_17)
  Call WKTblColFormats("n", "rn", "rn", "rn", "rn", "rn", "rn")
  Call WKTableHeadings(rep, sCol1Caption, sCol2Caption, sCol3Caption, sCol4Caption, sCol5Caption, sCol6Caption, sCol7Caption)
End Sub


Public Function WKOut(ByVal rep As Reporter, ByVal OutputType As WKOUT_TYPE, Optional ByVal OutputText As Variant, Optional ByVal OutputValue As Variant, Optional ByVal OutPutRef As Variant, Optional ByVal bCurrency As Boolean = False, Optional ByVal bNegative As Boolean = False) As Boolean
  On Error GoTo WKOut_ERR
  
  Call xSet("WKOut")
  
  If IsMissing(OutputText) Then OutputText = "" Else OutputText = CStr(OutputText)
  If IsMissing(OutputValue) Then OutputValue = "" Else OutputValue = FormatWN(OutputValue, IIf(bCurrency, "£", ""), bNegative)
  If IsMissing(OutPutRef) Then OutPutRef = "" Else OutPutRef = CStr(OutPutRef)
  
  Select Case OutputType
    Case WK_BLANK_LINE
      Call rep.Out(vbCrLf)
    Case WK_SECTION_HEADER, WK_SECTION_HEADER_DETAILS, WK_SECTION_HEADER_LESS, WK_SECTION_HEADER_BENEFIT, WK_SECTION_HEADER_VALUE
      Select Case OutputType
        Case WK_SECTION_HEADER_DETAILS
          OutputText = "Details:"
        Case WK_SECTION_HEADER_VALUE
          OutputText = "Value:"
        Case WK_SECTION_HEADER_LESS
          OutputText = "Less:"
        Case WK_SECTION_HEADER_BENEFIT
         OutputText = "Benefit:"
      End Select
      Call rep.Out(S_WK_NORMAL_BOLD_FONT & "{x=0}" & vbCrLf & "{x=3}" & OutputText & "{Arial=6}" & vbCrLf & vbCrLf)
    Case WK_SECTION_BREAK
      Call rep.Out(vbCrLf & "{line}" & vbCrLf)
    Case WK_ITEM_TEXT
      Call rep.Out(S_WK_NORMAL_FONT & S_WK_LEFT_MARGIN & OutputText & _
                   WKOutputRef(OutPutRef) & _
                   S_WK_RIGHT_BOLD_FONT & S_WKCOL_2 & OutputValue & vbCrLf)
    
    Case WK_ITEM_TEXT_BOLD
      Call rep.Out(S_WK_NORMAL_BOLD_FONT & S_WK_LEFT_MARGIN & OutputText & _
                   WKOutputRef(OutPutRef) & _
                   S_WK_RIGHT_BOLD_FONT & S_WKCOL_2 & OutputValue & vbCrLf)
    Case WK_ITEM_DESCRIPTION
      Call rep.Out(S_WK_NORMAL_FONT & S_WK_LEFT_MARGIN & OutputText & S_WK_NORMAL_BOLD_FONT & OutputValue & vbCrLf)
    Case WK_ITEM_NOTE
      Call rep.Out(S_WK_NORMAL_ITALIC_FONT & S_WK_LEFT_MARGIN & OutputText & vbCrLf)
    Case WK_ITEM_subtotal
      Call WKOut(rep, WK_SUBTOTAL_ONLY)
      Call rep.Out(S_WK_NORMAL_FONT & S_WK_LEFT_MARGIN & OutputText & _
                   WKOutputRef(OutPutRef) & _
                   S_WK_RIGHT_BOLD_FONT & S_WKCOL_2 & OutputValue & vbCrLf)
      
    Case WK_ITEM_Total
      Call rep.Out(vbCrLf)
      Call WKOut(rep, WK_SUBTOTAL_ONLY)
      Call rep.Out(S_WK_NORMAL_FONT & S_WK_LEFT_MARGIN & OutputText & _
                   WKOutputRef(OutPutRef) & _
                   S_WK_RIGHT_BOLD_FONT & S_WKCOL_2 & OutputValue & vbCrLf)
      Call WKOut(rep, WK_TOTAL_ONLY)
    Case WK_SUBTOTAL_ONLY
      Call rep.Out(S_WKCOL_2 & "{line=-7}" & vbCrLf)
    Case WK_TOTAL_ONLY
      Call rep.Out(S_WKCOL_2 & "{line=-7,d}" & vbCrLf)
    Case Else
      Call ECASE("Unknown output type in WKOut.")
  End Select

  WKOut = True

WKOut_END:
  Call xReturn("WKOut")
  Exit Function
WKOut_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKOut", "WK Out", "Error priting a worksheet, output type = " & OutputType & ".")
  Resume WKOut_END
End Function
Public Function WKTableTotalBen(rep As Reporter, ByVal ben As IBenefitClass)
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim value As Variant, Benefit As Variant, MadeGood As Variant
    
  BenArr(1) = ben.BenefitClass
  Call rep.Out(vbCrLf & vbCrLf)
  Call WKTableTotals(rep, "", FormatWN(ben.value(ITEM_VALUE)), FormatWN(ben.value(ITEM_MADEGOOD_NET), , True), FormatWN(ben.value(ITEM_BENEFIT)))

End Function
Public Function WKTableRow(rep As Reporter, ParamArray RowData()) As Boolean
  Dim lOffset As Long, lNumCols As Long, i As Long, s As String
   
  On Error GoTo WKTableRow_ERR
  
  Call xSet("WKTableRow")
  
  lOffset = LBound(RowData) - 1
  lNumCols = UBound(RowData) - lOffset
  
  If lNumCols <> m_lWKNumberOfColumns Then
    Call ECASE("WKTableRow" & vbCrLf & vbCrLf & "Parameters:" & lNumCols & vbCrLf & "Offsets:" & m_lWKNumberOfColumns)
    GoTo WKTableRow_END
  End If

  For i = 1 To m_lWKNumberOfColumns
    If Len(m_sWKColFormats(i)) > 0 Then s = s & "{Arial=7," & m_sWKColFormats(i)
    s = s & "}{x=" & m_lWKXOffsets(i) & "}" & RowData(i + lOffset)
  Next i
  
  If Len(s) > 0 Then
    rep.Out s & vbCrLf
  End If
  
WKTableRow_END:
  Call xReturn("WKTableRow")
  Exit Function
WKTableRow_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKTableRow", "WK Table Row", "Error creating a WK table row.")
  Resume WKTableRow_END
End Function
Public Function WKTableBlankRow(rep As Reporter) As Boolean
   
  On Error GoTo err_err
  
  Call xSet("WKTableBlankRow")
  
  Call rep.Out(vbCrLf)
  
err_end:
  Call xReturn("WKTableBlankRow")
  Exit Function
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "WKTableBlankRow", "WK Table Blank Row", "Error creating a WK table blank row.")
  Resume err_end
End Function

Public Function WKTableHeadings(rep As Reporter, ParamArray Headings()) As Boolean
  Dim lOffset As Long, lNumCols As Long, lNumRows As Long
  Dim i As Long, j As Long, k As Long, s As String, sHeadingLines() As String
  
  On Error GoTo WKTableHeadings_ERR
  
  Call xSet("WKTableHeadings")
  
  lOffset = LBound(Headings) - 1
  lNumCols = UBound(Headings) - lOffset
  
  If lNumCols > m_lWKNumberOfColumns Then
    Call ECASE("WKTableHeadings" & vbCrLf & vbCrLf & "Parameters:" & lNumCols & vbCrLf & "Offsets:" & m_lWKNumberOfColumns)
    Exit Function
  End If
    
  For i = 1 To lNumCols
   k = 0
    j = InStr(1, Headings(i + lOffset), "~")
    Do While j
      k = k + 1
      j = InStr(j, Headings(i + lOffset), "~")
      If j = 0 Then Exit Do
      j = j + 1
    Loop
    lNumRows = Max(lNumRows, k)
  Next i
  
  If lNumRows = 0 Then lNumRows = 1
  
  For j = 1 To lNumRows
    s = ""
    For i = 1 To lNumCols
      Call GetDelimitedValues(sHeadingLines, CStr(Headings(i + lOffset)), , , "~")
      If j <= UBound(sHeadingLines) Then
        If Len(m_sWKColFormats(i)) > 0 Then
          s = s & "{Arial=7," & m_sWKColFormats(i) & "}"
        Else
          s = s & S_WK_NORMAL_FONT
        End If
        If Len(sHeadingLines(j)) > 0 Then
          s = s & "{x=" & m_lWKXOffsets(i) & "}" & sHeadingLines(j)
        End If
      End If
    Next i
    s = s & vbCrLf
    Call rep.Out(s)
  Next j
  
  Call rep.Out(vbCrLf)
  
  WKTableHeadings = True
  
WKTableHeadings_END:
  Call xReturn("WKTableHeadings")
  Exit Function
WKTableHeadings_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKTableHeadings", "WK Table Headings", "Error creating the WK Table Headings.")
  Resume WKTableHeadings_END
  Resume
End Function

Public Function WKTableColTotals(ParamArray TotalCols()) As Boolean
  Dim lOffset As Long, lNumCols As Long, i As Long
  
  On Error GoTo WKTableColTotals_ERR
  
  Call xSet("WKTableColTotals")
  
  lOffset = LBound(TotalCols) - 1
  lNumCols = UBound(TotalCols) - lOffset
  
  If lNumCols <> m_lWKNumberOfColumns Then
    Call ECASE("WKTableColTotals" & vbCrLf & vbCrLf & "Parameters:" & lNumCols & vbCrLf & "Offsets:" & m_lWKNumberOfColumns)
    Exit Function
  End If
  
  ReDim m_bWKTotalCols(1 To m_lWKNumberOfColumns)
  
  For i = 1 To m_lWKNumberOfColumns
    m_bWKTotalCols(i) = TotalCols(i + lOffset)
  Next i
  
WKTableColTotals_END:
  Call xReturn("WKTableColTotals")
  Exit Function
WKTableColTotals_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKTableColTotals", "WK Table Col Totals", "Error setting the WK table col totals.")
  Resume WKTableColTotals_END
End Function


Public Function WKTblColXOffsets(ParamArray XOffsets()) As Boolean
  Dim lOffset As Long, i As Long
  
  On Error GoTo WKTblColXOffsets_ERR
  
  Call xSet("WKTblColXOffsets")
  
  lOffset = LBound(XOffsets) - 1
  m_lWKNumberOfColumns = UBound(XOffsets) - lOffset
  ReDim m_lWKXOffsets(1 To m_lWKNumberOfColumns)
  
  For i = 1 To m_lWKNumberOfColumns
    m_lWKXOffsets(i) = CLng(XOffsets(i + lOffset))
  Next i
  
  WKTblColXOffsets = True
  
WKTblColXOffsets_END:
  Call xReturn("WKTblColXOffsets")
  Exit Function
WKTblColXOffsets_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKTblColXOffsets", "WK Tbl Col X Offsets", "Error setting the worksheet table X offsets.")
  Resume WKTblColXOffsets_END
  Resume
End Function
Public Function WKTblColFormats(ParamArray ColFormats()) As Boolean
  Dim lOffset As Long, lNumCols As Long, i As Long
  
  On Error GoTo WKTblColFormats_ERR
  
  Call xSet("WKTblColFormats")
  
  lOffset = LBound(ColFormats) - 1
  
  lNumCols = UBound(ColFormats) - lOffset
  
  If lNumCols <> m_lWKNumberOfColumns Then
    Call ECASE("WKTblColFormats" & vbCrLf & vbCrLf & "Parameters:" & lNumCols & vbCrLf & "lOffsets:" & m_lWKNumberOfColumns)
    Exit Function
  End If
  
  ReDim m_sWKColFormats(1 To m_lWKNumberOfColumns)
  
  For i = 1 To m_lWKNumberOfColumns
    m_sWKColFormats(i) = ColFormats(i + lOffset)
  Next i
  
  WKTblColFormats = True
  
WKTblColFormats_END:
  Call xReturn("WKTblColFormats")
  Exit Function
WKTblColFormats_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKTblColFormats", "WK Tbl Col Formats", "Error specifying the column formats for a work sheet.")
  Resume WKTblColFormats_END
End Function
'************************* END WK functions ******************************
Public Function StartAutoSTD(ByVal sSQL As String, db As Database, sReportHeader As String, ByVal Dest As REPORT_TARGET, Optional sTitle As String = "Errors") As Boolean
  Dim rs As Recordset
  
  On Error GoTo StartAutoSTD_ERR
  
  Call xSet("StartAutoSTD")
  
  If db Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "StartAutoSTD", "The db is nothing")
  Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
  If rs Is Nothing Then Call Err.Raise(ERR_RS_IS_NOTHING, "StartAutoSTD", "The recordset created from sql = " & sSQL & " on database " & db.Name & " is nothing.")
 
  StartAutoSTD = ReportErrors(rs, sReportHeader, Dest, sTitle)
  
StartAutoSTD_END:
  Call xReturn("StartAutoSTD")
  Exit Function
StartAutoSTD_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "StartAutoSTD", "Start Auto STD", "Error starting a standard auto report.")
  Resume StartAutoSTD_END
  Resume
End Function

Public Sub SetActivePrinter(ByVal Name As String)
  Dim rep As Reporter
  
  Set rep = ReporterNew(Nothing)
  Call rep.SetActivePrinter(Name)
  
End Sub

Public Function ReportErrors(rs As Recordset, sReportHeader As String, ByVal Dest As REPORT_TARGET, Optional sTitle As String = "Errors") As Boolean
  Dim ac As AutoClass
  Dim rep As Reporter
  
  On Error GoTo ReportErrors_Err

  Call xSet("ReportErrors")
  Set ac = New AutoClass
  If Not ac.InitAutoData("ReportErrors", rs) Then GoTo ReportErrors_End
  Set rep = ReporterNew(Nothing)
  If Not rep.InitReport(sTitle & vbCrLf & vbCrLf, Dest, LANDSCAPE, True) Then GoTo ReportErrors_End
  
  ac.dateFormat = "DD/MM/YYYY hh:mm:ss"
  ac.ReportHeader = sReportHeader
  
  If sReportHeader = "Magnetic Media Warnings and Errors" Then
    Call ac.AddFieldFormat("Description", "{WRAP}")
    Call ac.AddFieldFormat("Benefit", "{WRAP}")
  End If
  
  If Left$(sReportHeader, 18) = "Import errors for:" Then
    Call ac.AddFieldFormat("ErrorDescription", "{WRAP}")
  End If
  
  'ac.ReportFormat = "{Font=Arial,20}"
  Call ac.ShowReport(rep)
  rep.EndReport
  If Dest = PREPARE_REPORT Then rep.PreviewReport

ReportErrors_End:
  Set rep = Nothing
  If Not ac Is Nothing Then
    Call ac.Kill
    Set ac = Nothing
  End If
  Call xReturn("ReportErrors")
  Exit Function

ReportErrors_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReportErrors", "Report Errors", "Error displaying the import errors")
  Resume ReportErrors_End
  Resume
End Function

Private Function TickOutPlusX(lXStart As Long, v As Variant) As String
  TickOutPlusX = "{x=" & lXStart & "}" & TickOut(v)
End Function
Private Function BoxText(vText As Variant) As String
  BoxText = " " & AddEscapeChars(vText) & " "
End Function
Private Function HMITColText(sText As String, lColumn As Long)
  HMITColText = "{Arial=6,n}{x=" & lColumn + 2 & "}" & sText
End Function
Private Function OutLineBoxR(lXStart, lWidth As Long, sngHeight As Single, vText As Variant) As String
  OutLineBoxR = "{Times=10,bi}{x=" & lXStart & "}{WBTEXTBOXR=" & lWidth & "," & sngHeight & "," & BoxText(vText) & "}"
End Function
Private Function OutLineBoxL(lXStart, lWidth As Long, sngHeight As Single, vText As Variant) As String
  OutLineBoxL = "{Times=10,bi}{x=" & lXStart & "}{WBTEXTBOXL=" & lWidth & "," & sngHeight & "," & BoxText(vText) & "}"
End Function
Private Function FillBox(lXStart, lWidth As Long, sngHeight As Single, sText As String) As String
  FillBox = "{Arial=10,nb}{x=" & lXStart & "}{BWTEXTBOX=" & lWidth & "," & sngHeight & "," & sText & "}"
End Function
Private Function FillBoxHeader(lXStart, lWidth As Long, sngHeight As Single, sText As String) As String
  FillBoxHeader = "{Arial=10,nb}{x=" & lXStart & "}{BWTEXTBOXL=" & lWidth & "," & sngHeight & "," & sText & "}"
End Function
Private Function HMITColTextLower(sText As String, lColumn As Long)
  HMITColTextLower = "{Arial=6,n}{x=" & lColumn & "}" & sText
End Function
Private Function FillBoxNIC(lXStart, lWidth As Long, sngHeight As Single, sText As String) As String
  FillBoxNIC = "{Arial=9,nb}{x=" & lXStart & "}{BWTEXTBOX=" & lWidth & "," & sngHeight & "," & sText & "}"
End Function
Private Function FillBoxNICHeader(lXStart, lWidth As Long, sngHeight As Single, sText As String) As String
  FillBoxNICHeader = "{Arial=7,nb}{x=" & lXStart & "}{BWTEXTBOX=" & lWidth & "," & sngHeight & "," & sText & "}"
End Function

Private Function HMITText(sText, Optional lXStart As Long = 0) As String
  '{Arial=12,B}
  HMITText = "{x=" & lXStart + 4 & "}" & "{XREL=100}{YREL=50}{Arial=8,nb}" & sText
End Function
Private Function HMITEquals() As String
  HMITEquals = "{Arial=11,B}{x=" & L_HMIT_COL_3 - 2 & "}="
End Function
Private Function HMITMinus(lXStart As Long) As String
  HMITMinus = "{Arial=11,B}{x=" & L_HMIT_COL_2 - 0.81 & "}-"
End Function
Private Function HMITStandardCol(sCaption As String, vBoxValue As Variant, bBoxNumber, sBoxNumber As String, lColX As Long)
  HMITStandardCol = OutLineBoxR(lColX, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, vBoxValue) & _
                    IIf(bBoxNumber, FillBox(lColX, 4, 2, sBoxNumber), "") & _
                    "{Arial=7}{x=5}" & sCaption & vbCrLf & "{Arial=6}" & vbCrLf & vbCrLf
End Function
Private Function HMITMielageCaptionsOut(lXStart) As String

  HMITMielageCaptionsOut = "{Arial=6}{x=" & lXStart - 2 & "}2,499 or less" & _
                           "{x=" & lXStart + 11 - 2 & "}2,500 to 17,999" & _
                           "{x=" & lXStart + 22 - 2 & "}18,000 or more"
End Function
Private Function HMITMileageTicks(CompanyCar As IBenefitClass, lXStart As Long) As String
  HMITMileageTicks = TickOutPlusX(lXStart, GetBenItem(CompanyCar, car_P46lowMiles)) & TickOutPlusX(lXStart + 11, GetBenItem(CompanyCar, car_P46MediumMiles)) & TickOutPlusX(lXStart + 22, GetBenItem(CompanyCar, car_P46HighMiles))
End Function
Private Function HMITFieldTrim(ByVal s As String, ByVal MaxLength As Long, Optional Append As String = "...") As String
  If Not p11d32.ReportPrint.HMITFieldTrim Then
    HMITFieldTrim = s
    Exit Function
  End If
  
  If Len(s) <= MaxLength Then
    HMITFieldTrim = s
    Exit Function
  End If
  
  If Len(Append) > MaxLength Then
    Call Err.Raise(ERR_STRING_TOO_LONG, ErrorSource(Err, "HMITFieldTrim"), "Append string is longer than MaxLength")
  End If
  
  s = Left$(s, MaxLength - Len(Append))
  HMITFieldTrim = s & Append
End Function
Private Function HMITCarMakeAndModel(ByVal benCar As IBenefitClass) As String
  HMITCarMakeAndModel = HMITFieldTrim(GetBenItem(benCar, car_Make_db) & " " & GetBenItem(benCar, car_Model_db), 30)
End Function
Private Function HMITCar(rep As Reporter, ee As Employee, CompanyCar1 As IBenefitClass, lCar1Number As Long, CompanyCar2 As IBenefitClass, lCar2Number As Long) As Long
  Dim vTotalCarBenefit As Variant, vTotalFuelBenefit As Variant
  Dim FuelTypeString_Car1 As String
  Dim FuelTypeString_Car2 As String
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim totalAccessoriesCar1 As String
  Dim totalAccessoriesCar2 As String
  Dim Withdrawndate As String

  Const HMIT_CAR_COL1 As Long = 33
  Const HMIT_CAR_COL2 As Long = 64
  
  On Error GoTo HMITCar_ERR
  Call xSet("HMITCar")
    
  'Make & model
  Call HMITSectionHeader(rep, HMIT_F, "Cars and car fuel " & "{Arial=6,ni}If more than two cars were made available, either at the same time or in succession, please give details on a separate sheet")   '& "{Arial=6}", "{Arial=6,i}" & vbCrLf & "{Arial=6}"))
  
  If lCar1Number = 1 Then
    BenArr(1) = BC_COMPANY_CARS_F
    Call SumBenefitFWNRPT(ee, 0, 0, 0, vTotalCarBenefit, BenArr)
    BenArr(1) = BC_FUEL_F
    Call SumBenefitFWNRPT(ee, 0, 0, 0, vTotalFuelBenefit, BenArr)
  End If
        
  'Make And model
  
  Call rep.Out(vbCrLf & HMITColText("Car " & CStr(lCar1Number), HMIT_CAR_COL1) & _
               HMITColText("Car " & CStr(lCar2Number), HMIT_CAR_COL2) & vbCrLf)
  
  'FIX XX by CAD 24/04/2002 to fix make model bug preventing printing added HMITFieldTrim
  Call rep.Out(OutLineBoxL(HMIT_CAR_COL1, 30, L_HMIT_STANDARDBOX_HEIGHT, HMITCarMakeAndModel(CompanyCar1)) & _
                 OutLineBoxL(HMIT_CAR_COL2, 30, L_HMIT_STANDARDBOX_HEIGHT, HMITCarMakeAndModel(CompanyCar2)) & _
                 LineText("{x=6}Make and Model"))
  
  
  
'Date first registered
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(CompanyCar1, car_Registrationdate_db)) & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(CompanyCar2, car_Registrationdate_db)) & _
               LineText("{x=6}Date first registered"))
  
' KA: New section for 2002 - 2003. Boxes for reporting CO2 emissions and engine size. Replaces business mileage.
  
' KA: Approved CO2 emissions
  
  'no C02 box pos
  
  Dim noCO2BoxXPos1 As Long, noCO2BoxXPos2 As Long
  noCO2BoxXPos1 = L_HMIT_COL_1 - 3
  noCO2BoxXPos2 = L_HMIT_COL_4 - 3
    
  Call rep.Out("{x=6}{Arial=7,n}Approved CO2 emissions figure for cars " & _
  OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetCO2DisplayFigure(CompanyCar1) & "g/km") & _
  TickOutPlusX(noCO2BoxXPos1, GetBenItem(CompanyCar1, car_p46NoApprovedCO2Figure_db)) & _
  "{x=51}{Arial=6,ni}See P11D Guide for" & _
  OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetCO2DisplayFigure(CompanyCar2) & "g/km") & _
  TickOutPlusX(noCO2BoxXPos2, IIf(CompanyCar2 Is Nothing, False, GetBenItem(CompanyCar2, car_p46NoApprovedCO2Figure_db))) & _
  "{x=83}{Arial=6,ni}See P11D Guide for" & vbCrLf & _
  "{x=6}{Arial=7,n}registered on or after 1 January 1998 " & _
  "{Arial=6,ni}Tick" & _
  "{Arial=7,n}" & _
  "{x=51}{Arial=6,ni}details of cars that have" & _
  "{x=83}{Arial=6,ni}details of cars that have" & vbCrLf & _
  "{x=6}{Arial=6,ni}box if the car does not have an approved CO2 figure" & _
  "{x=51}{Arial=6,ni}no approved CO2 figure " & _
  "{x=83}{Arial=6,ni}no approved CO2 figure " & S_WK_LINE_BREAK_FONT & vbCrLf & vbCrLf)
  
'Approved zero emissions milage
  Call rep.Out("{x=6}{Arial=7,n}Approved zero emissions mileage if hybrid's" & _
  OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetApprovedZeroEmissionsMileage(CompanyCar1) & " miles") & _
  OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetApprovedZeroEmissionsMileage(CompanyCar2) & " miles") & _
  S_WK_LINE_BREAK_FONT & vbCrLf & "{x=6}{Arial=7,n}CO2 emissions is between 1-50 (inclusive)" & S_WK_LINE_BREAK_FONT & vbCrLf & vbCrLf)
  
  
'KA: Engine size
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(CompanyCar1, car_enginesize_db) & "cc") & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(CompanyCar2, car_enginesize_db) & "cc") & _
               LineText("{x=6}Engine size"))
  
'KA: Type of fuel or power
  If Not CompanyCar1 Is Nothing Then
    FuelTypeString_Car1 = IIf(GetBenItem(CompanyCar1, car_P46WithdrawnWithoutReplacement), "", GetBenItem(CompanyCar1, car_p46FuelTypeString))
  End If
  If Not CompanyCar2 Is Nothing Then
    FuelTypeString_Car2 = IIf(GetBenItem(CompanyCar2, car_P46WithdrawnWithoutReplacement), "", GetBenItem(CompanyCar2, car_p46FuelTypeString))
  End If
  
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH / 2, L_HMIT_STANDARDBOX_HEIGHT, FuelTypeString_Car1) & _
                 OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH / 2, L_HMIT_STANDARDBOX_HEIGHT, FuelTypeString_Car2) & _
                 "{x=6}{Arial=7,n}Type of fuel or power used " & "{Arial=6,ni}Please use the key" & "{Arial=7,n}" & vbCrLf & _
                 "{x=6}{Arial=6,i}letter shown in the P11D Guide" & vbCrLf & vbCrLf)
  
  
  
  'LK - code added to only show dates when they differ from the start or end of the year
    Call rep.Out(OutLineBoxR(HMIT_CAR_COL1 + 3, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, IIf(GetBenItem(CompanyCar1, Car_AvailableFrom_db) <> p11d32.Rates.value(TaxYearStart), GetBenItem(CompanyCar1, Car_AvailableFrom_db), "")) & _
                 OutLineBoxR(HMIT_CAR_COL1 + L_HMIT_STANDARDBOX_WIDTH + 6, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, IIf(GetBenItem(CompanyCar1, Car_AvailableTo_db) <> p11d32.Rates.value(TaxYearEnd), GetBenItem(CompanyCar1, Car_AvailableTo_db), "")) & _
                 OutLineBoxR(HMIT_CAR_COL2 + 3, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, IIf(GetBenItem(CompanyCar2, Car_AvailableFrom_db) <> p11d32.Rates.value(TaxYearStart), GetBenItem(CompanyCar2, Car_AvailableFrom_db), "")) & _
                 OutLineBoxR(HMIT_CAR_COL2 + L_HMIT_STANDARDBOX_WIDTH + 6, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, IIf(GetBenItem(CompanyCar2, Car_AvailableTo_db) <> p11d32.Rates.value(TaxYearEnd), GetBenItem(CompanyCar2, Car_AvailableTo_db), "")) & _
                 HMITColTextLower("From", HMIT_CAR_COL1) & _
                 HMITColTextLower("To", HMIT_CAR_COL1 + L_HMIT_STANDARDBOX_WIDTH + 4) & _
                 HMITColTextLower("From", HMIT_CAR_COL2) & _
                 HMITColTextLower("To", L_HMIT_COL_3 + 2) & _
                 ("{x=6}{Arial=7,n}Dates car was available{Arial=6,ni} Do not complete the" & vbCrLf & _
                 "{x=6}{Arial=6,ni}'From' box if the car was available on " & DateValReadToScreen(p11d32.Rates.value(LastTaxYearEnd)) & vbCrLf & _
                 "{x=6}{Arial=6,ni}or the 'To' box if it continued to be" & vbCrLf) & _
                 "{x=6}available on " & DateValReadToScreen(p11d32.Rates.value(NextTaxYearStart)) & vbCrLf & vbCrLf)

  
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar1, car_ListPrice_db)) & _
             OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar2, car_ListPrice_db)) & _
             "{x=6}{Arial=7,n}List price of car " & "{Arial=6,ni}Including car and standard" & vbCrLf & _
             "{x=6}{Arial=6,ni}accessories only: if there is no list price, or if it is a" & "{Arial=7,n}" & vbCrLf & _
             "{x=6}{Arial=6,ni}classic car, employers see booklet 480" & vbCrLf & _
             "{x=6}{Arial=6,ni}" & vbCrLf & vbCrLf)

  If Not CompanyCar1 Is Nothing Then
    totalAccessoriesCar1 = FormatWNRPT(CompanyCar1.value(car_AccessoriesNew_db) + CompanyCar1.value(car_AccessoriesOriginal_db) - CompanyCar1.value(car_CheapAccessories_db), "£")
  End If
  If Not CompanyCar2 Is Nothing Then
    totalAccessoriesCar2 = FormatWNRPT(CompanyCar2.value(car_AccessoriesNew_db) + CompanyCar2.value(car_AccessoriesOriginal_db) - CompanyCar2.value(car_CheapAccessories_db), "£")
  End If
    
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, totalAccessoriesCar1) & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, totalAccessoriesCar2) & _
               "{x=6}{Arial=7,n}Accessories {Arial=6,ni}All non-standard accessories," & vbCrLf & _
               "{x=6}{Arial=6,ni}see P11D Guide" & vbCrLf & vbCrLf)
               
  'capital contrib
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar1, car_CapitalContributionRestricted)) & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar2, car_CapitalContributionRestricted)) & _
               "{x=6}{Arial=7,n}Capital contributions (maximum £5,000) the" & vbCrLf & _
               "{x=6}employee made towards the cost of car or" & vbCrLf & _
               "{x=6}accessories" & vbCrLf & vbCrLf)
               
  'made good for private use
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar1, car_ActualAmountMadeGood)) & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar2, car_ActualAmountMadeGood)) & _
               "{x=6}{Arial=7,n}Amount paid by employee for private use of" & vbCrLf & _
               "{x=6}the car" & vbCrLf & vbCrLf)
               
  'free fuel withdrawn
  
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(CompanyCar1, Car_FuelAvailableTo_calc)))
  Call rep.Out(TickOutPlusX(L_HMIT_COL_1 - 3, GetBenItem(CompanyCar1, car_fuelreinstated_calc)))
  
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(CompanyCar2, Car_FuelAvailableTo_calc)))
  Call rep.Out(TickOutPlusX(L_HMIT_COL_4 - 3, GetBenItem(CompanyCar2, car_fuelreinstated_calc)))
  
  Call rep.Out("{x=6}{Arial=7,n}Date free fuel was withdrawn" & vbCrLf)
  Call rep.Out("{x=6}{Arial=6,ni}Tick if reinstated in year (see P11D Guide)" & vbCrLf & vbCrLf)
             
  
  'cash equivalent
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar1, car_carBenefit)) & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar2, car_carBenefit)) & _
               "{x=6}{Arial=7,n}Cash equivalent or relevant amount" & vbCrLf & "{x=6}{Arial=7,n}for each car" & vbCrLf & vbCrLf)
               
  'total cash equivalent
  If lCar1Number = 1 Then
    Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, vTotalCarBenefit) & _
                 FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BC_COMPANY_CARS_F, BCT_HMIT_BOX_NUMBER)) & _
                 FillBoxNIC(L_HMIT_COL_5, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
                 "{x=6}{Arial=8,nb}Total cash equivalent or relevant amount of all cars made available in " & DateValReadToScreen(p11d32.Rates.value(TaxFormYear)) & vbCrLf & vbCrLf)
  End If
  
  'cash equivalent on fuel
  Call rep.Out(OutLineBoxR(HMIT_CAR_COL1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar1, car_FuelBenefit)) & _
               OutLineBoxR(HMIT_CAR_COL2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(CompanyCar2, car_FuelBenefit)) & _
               "{x=6}{Arial=7,n}Cash equivalent or amount foregone in" & vbCrLf & "{x=6}{Arial=7,n}respect of fuel for each car" & vbCrLf & vbCrLf)
                 

  'ONLY PRINT WITHDRAWN IF VALID
  
                  
  If lCar1Number = 1 Then
  'total cash equiv on fuel for all cars
    Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, vTotalFuelBenefit) & _
                 FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BC_FUEL_F, BCT_HMIT_BOX_NUMBER)) & _
                 FillBoxNIC(L_HMIT_COL_5, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
                 "{x=6}{Arial=8,nb}Total cash equivalent or amount foregone in respect of fuel for all cars made available in " & vbCrLf & "{x=6}{Arial=8,nb}" & DateValReadToScreen(p11d32.Rates.value(TaxFormYear)) & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf)
  End If
HMITCar_END:
  Call xReturn("HMITCar")
  Exit Function
HMITCar_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITCar", "HMIT Car", "Error printing a company car for the HMIT return")
  Resume HMITCar_END
  Resume
End Function
Public Sub HMITBanner(rep As Reporter, sBannerText As String, Optional bIRFlag As Boolean = True)
  Call rep.Out(HMITBannerString(sBannerText, bIRFlag))
End Sub


Public Function HMITBannerString(sBannerText As String, Optional bIRFlag As Boolean = True) As String
  Dim sDraft As String
  If p11d32.ReportPrint.DraftReports Then
    sDraft = "DRAFT"
  Else
    sDraft = ""
  End If
  'HMITBannerString = "{Arial=10,bi}{X=" & IIf(bIRFlag, 15, 0) & "}{BWTEXTBOXR=" & IIf(bIRFlag, 82, 97) & ",3," & AddEscapeChars(sBannerText) & " }" & _
  '      "{Arial=10,b}" & vbCrLf & "{x=3}" & IIf(bIRFlag, "HM Revenue", "") & vbCrLf & _
  '      "{Arial=13,b}{x=3}& Customs"
        
        
  HMITBannerString = "{Arial=12,bi}{X=3}" & sDraft & "{LEFT}{X=52}" & AddEscapeChars(sBannerText) & _
        "{Arial=10,b}" & vbCrLf & "{x=3}" & IIf(bIRFlag, "HM Revenue", "") & vbCrLf & _
        IIf(bIRFlag, "{Arial=13,b}{x=3}& Customs", "")
  
End Function

'Public Function HMITBannerStringSecondLine(sBannerText As String, Optional bIRFlag As Boolean = True) As String
'  HMITBannerStringSecondLine = "{Arial=6,n}" & vbCrLf & "{Arial=12,bi}{X=" & IIf(bIRFlag, 15, 0) & "}{BWTEXTBOXR=" & IIf(bIRFlag, 82, 97) & ",3," & AddEscapeChars(sBannerText) & " }"
'End Function
Private Function HMITBullet(Optional lXStart As Long = 0) As String
  HMITBullet = "{Arial=10,nb}{x=" & lXStart & "}{WB}•"
End Function
Private Sub HMITPageHeader(rep As Reporter, ee As Employee)
  'Page 1 Title
  Call HMITBanner(rep, "P11D EXPENSES AND BENEFITS " & IIf(p11d32.AppYear > 2001, " ", "") & p11d32.Rates.value(TaxFormYear)) 'km - changed 2000 to 2001
  'notes to employer
  Call rep.Out(vbCrLf & "{PUSHY}")
  

                  
    Call rep.Out("{yrel=0}{x=3}{Arial=8,nb}{FORERGB=8224125}Make sure your entries are clear on both sides of the form.")
  'notes to emplyer
    Call rep.Out("{yrel=0}{y=2}{FILLRGB=" & 14803425 & "}" & _
               "{x=48}{BOX=50,5.5,F}{RESETCOLORS}" & _
               "{X=49}{YREL=5}{Arial=9,nb}Note to employer" & vbCrLf & "{YREL=4}" & _
               "{Arial=8,n}{x=49}Fill in this return for a director or an employee for the year to 5 April " & Format$(p11d32.Rates.value(TaxYearEnd), "YYYY") & "." & vbCrLf & _
               "{x=49}Send all your P11Ds and one " & S_P11D_B & " to your HMRC office by 6 July " & Format$(p11d32.Rates.value(TaxYearEnd), "YYYY") & "." & vbCrLf & _
               "{x=49}Don't include payrolled benefits on the P11D, go to www.gov.uk/guidance" & vbCrLf & _
               "{x=49}/paying-your-employees-expenses-and-benefits-through-your-payroll")

                
                  
  'notes to employee
  Call rep.Out("{POP}")
    
    Call rep.Out("{yrel=1000}{y=8}{FILLRGB=" & S_LIGHTGREY & "}" & _
               "{x=48}{BOX=50,4.5,F}{RESETCOLORS}" & _
               "{X=49}{YREL=5}{Arial=9,b}Note to employee" & vbCrLf & "{YREL=4}" & _
               "{Arial=8,n}{x=49}Keep this form in a safe place, You'll need it to complete your " & Format$(p11d32.Rates.value(TaxYearStart), "YYYY") & " to " & Format$(p11d32.Rates.value(TaxYearEnd), "YYYY") & vbCrLf & _
               "{x=49}tax return if you get one. The box numberings on this form are the same as" & vbCrLf & _
               "{x=49}on the 'Employment' page of the tax return")
  
  '********
  
  
  
  'section line
  'Call rep.Out("{x=3}{LINE=94}")
  
End Sub
Private Sub HMITEmpDetails(rep As Reporter, ee As Employee)
  Dim ben As IBenefitClass
  Dim benEmployer As IBenefitClass
  
  Set benEmployer = p11d32.CurrentEmployer
  Set ben = ee
  'Employers/employees details
  'was 5.5
  Call rep.Out("{x=3}{y=5}{Arial=3,n}" & vbCrLf & _
               "{x=3}{Arial=7,n}Employer name" & vbCrLf & _
               OutLineBoxL(3, 38, L_HMIT_STANDARDBOX_HEIGHT, benEmployer.Name) & vbCrLf & _
               "{x=3}{Arial=3,n}" & vbCrLf & _
               "{x=3}{Arial=7,n}Employer PAYE reference" & vbCrLf & _
               OutLineBoxL(3, 25, L_HMIT_STANDARDBOX_HEIGHT, benEmployer.value(employer_Payeref_db)) & vbCrLf & _
               "{Arial=3,n}" & vbCrLf & "{x=3}{Arial=7,n}Employee name" & vbCrLf & _
               OutLineBoxL(3, 38, L_HMIT_STANDARDBOX_HEIGHT, "Surname: " & ben.value(ee_Surname_db)) & vbCrLf & _
               "{Arial=3,n}" & vbCrLf & "{Arial=7,n}" & OutLineBoxL(3, 38, L_HMIT_STANDARDBOX_HEIGHT, "First name(s): " & ee.ForeNames) & "{x=44}{Arial=7,n}If a director tick here " & TickOut(ben.value(ee_Director_db)) & "{Arial=7,n}{x=60}Date of birth {Arial=7,i}in figures (if known)" & "{Arial=7,n}" & OutLineBoxR(78.5, 18, L_HMIT_STANDARDBOX_HEIGHT, DateValReadToScreenOnlyValidDates(ben.value(ee_DOB_db))) & vbCrLf & _
               "{x=3}{Arial=3,n}" & vbCrLf & _
               "{x=3}{Arial=7,n}Works number / department" & "{x=40}National Insurance number" & vbCrLf & _
               OutLineBoxL(3, 23, L_HMIT_STANDARDBOX_HEIGHT, ee.PersonnelNumber) & OutLineBoxR(40, 21, L_HMIT_STANDARDBOX_HEIGHT, ben.value(ee_NINumber_db)) & _
               "{x=78}{Arial=7,n}Gender M - Male F - Female" & _
               OutLineBoxL(94, 2, L_HMIT_STANDARDBOX_HEIGHT, IIf(ben.value(ee_Gender_db) <> S_GENDER_NA, ben.value(ee_Gender_db), "")))

'"{x=3}{Arial=7,n}Employee name{x=78}Date of birth {Arial=7,i}in figures (if known)"
   Call rep.Out(vbCrLf & vbCrLf & vbCrLf)
 '              OutLineBoxL(50, 32, L_HMIT_STANDARDBOX_HEIGHT, ee.FullName) & _
 '              "{X=93}" & TickOut(ben.value(ee_Director_db)) & _
 '              "{Arial=6,n}{x=85}If a director" & vbCrLf & _
 '              "{Arial=6,n}{x=85}tick here" & vbCrLf & vbCrLf & "{Arial=3,n}" & vbCrLf & _

 '              "{x=50}Works number / department" & _
 '              "{Arial=7,nr}{x=90} National Insurance number" & vbCrLf & "{Arial=3,n}" & vbCrLf & _

 '              OutLineBoxL(50, 23, L_HMIT_STANDARDBOX_HEIGHT, ee.PersonnelNumber) & _
 '              OutLineBoxR(75, 21, L_HMIT_STANDARDBOX_HEIGHT, ben.value(ee_NINumber_db)) & vbCrLf & vbCrLf)
  
  
'
  
  Call rep.Out("{x=3}{LINE=94}")
'  If p11d32.AppYear = 2000 Then
'    Call rep.Out("{x=3}{Arial=7,nb}From 6 April 2000 employers pay Class 1A National Insurance contributions on more benefits." & vbCrLf & _
'                  "{x=3}{Arial=7,nb}These are shown in boxes which have a " & FillBoxNICHeader(28, 2, 1, "1A") & "{Arial=7,nb}{x=31}indicator" & vbCrLf & vbCrLf)
'  Else
    'Call rep.Out("{x=3}{Arial=7,nb}Employers pay Class 1A National Insurance contributions on most benefits. These are shown in boxes which have a " & FillBoxNICHeader(74, 2, 1, "1A") & "{Arial=7,nb}{x=77}indicator" & vbCrLf & vbCrLf)
    Call rep.Out("{x=3}{Arial=7,nb}Employers pay Class 1A National Insurance contributions on most benefits. These are shown in boxes which have a [1A] indicator" & vbCrLf & vbCrLf)
'  End If
End Sub
Private Sub HMITWhichSection(rep As Reporter, ee As Employee, lSections As Long)
  Dim i As Long
  Dim lCarsFound As Long, lBenefitStartIndexCars As Long
  Dim lLoansFound As Long, lBenefitStartIndexLoans As Long
  
  
  lBenefitStartIndexCars = 1
  lBenefitStartIndexLoans = 1
  'print all the sections
   For i = 0 To HMIT_SECTIONS.[_HMIT_COUNT] - 1
    If lSections And 2 ^ i Then
      Select Case i
        Case HMIT_SECTIONS.HMIT_F
          Call HMITSection(rep, ee, i, lBenefitStartIndexCars, lCarsFound)
        Case HMIT_SECTIONS.HMIT_H
          Call HMITSection(rep, ee, i, lBenefitStartIndexLoans, lLoansFound)
      Case Else
        Call HMITSection(rep, ee, i, 0, 0)
      End Select
    End If
  Next
  'do remaining loans/cars
  If lSections And 2 ^ HMIT_SECTIONS.HMIT_F Then
    Do While HMITSection(rep, ee, HMIT_SECTIONS.HMIT_F, lBenefitStartIndexCars, lCarsFound)
    Loop
  End If

  If lSections And 2 ^ HMIT_SECTIONS.HMIT_H Then
    Do While HMITSection(rep, ee, HMIT_SECTIONS.HMIT_H, lBenefitStartIndexLoans, lLoansFound)
    Loop
  End If
      
End Sub
Private Function SumBenefitFWNRPT(ee As Employee, Description As Variant, value As Variant, MadeGood As Variant, Benefit As Variant, BenArr() As BEN_CLASS, Optional VALUE_ENUM As Long = ITEM_VALUE, Optional MADEGOOD_ENUM As Long = ITEM_MADEGOOD_NET, Optional BENEFIT_ENUM As Long = ITEM_BENEFIT, Optional sIRDesc As String) As Long
  SumBenefitFWNRPT = ee.SumBenefit(Description, value, MadeGood, Benefit, BenArr(), VALUE_ENUM, MADEGOOD_ENUM, BENEFIT_ENUM, sIRDesc)
  If SumBenefitFWNRPT Then
    value = FormatWNRPT(value)
    MadeGood = FormatWNRPT(MadeGood)
    Benefit = FormatWNRPT(Benefit)
  End If
End Function
Private Function HMITSection(rep As Reporter, ee As Employee, HMITS As HMIT_SECTIONS, lBenefitStartIndex As Long, lBenefitsFound As Long) As Boolean
  Dim Description As String, value As Variant, MadeGood As Variant, Benefit As Variant
  Dim loans As loans, l As Long
  Dim benEmployer As IBenefitClass
  Dim benEmployee As IBenefitClass 'ek for section E change
  
  Dim BenArr() As BEN_CLASS
  
  ' PS 2/04 TTP#194 for having item in Entertainment box
  Dim tempDesc As String
  Dim tempvalue As Variant
  Dim tempmadegood As Variant
  Dim tempbenefit As Variant
  Dim entertainmenttick As Boolean
  Dim sIRDesc As String
  Dim tempDesc2 As String
  
  On Error GoTo HMITSection_ERR
  
  Call xSet("HMITSection")
  
  Set benEmployer = p11d32.CurrentEmployer
  Set benEmployee = p11d32.CurrentEmployer.CurrentEmployee 'ek for section E change
  Call rep.Out("{BEGINSECTION}")
  
  Call rep.Out("{x=3}{LINE=94}")
  
  Select Case HMITS
    Case HMIT_SECTIONS.HMIT_A
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_ASSETSTRANSFERRED_A
      Call HMITSectionHeader(rep, HMITS, "Assets transferred (cars, property, goods or other assets)")
      Call HMITColHeaders(rep, "Cost/market value", "or amount foregone", "Amount made good", "or from which tax deducted", "Cash equivalent", "or relevant amount")
      Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Description of asset", True)
    Case HMIT_SECTIONS.HMIT_B
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_PAYMENTS_ON_BEFALF_B
      Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr, , , , sIRDesc)
      If Len(sIRDesc) > 0 And Not LCase(sIRDesc) = "other" Then
        tempDesc2 = sIRDesc
      Else
        tempDesc2 = Description
      End If
      Call rep.Out(FillBox(3, 3, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.HMITSectionToHMITCode(HMITS)) & HMITText("{x=8}Payments made on behalf of employee") & vbCrLf & vbCrLf & _
               OutLineBoxL(23, (L_HMIT_COL_2 + L_HMIT_STANDARDBOX_WIDTH) - 23, 1.6, tempDesc2) & _
               FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BenArr(1), BCT_HMIT_BOX_NUMBER)) & _
               OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, Benefit) & _
               LineText("Description of payment") & _
               FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BenArr(1), BCT_HMIT_BOX_NUMBER)))
      BenArr(1) = BC_TAX_NOTIONAL_PAYMENTS_B
      Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr)
      Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, Benefit) & _
                   LineText("Tax on notional payments made during the year not borne by employee within 90 days of " & p11d32.Rates.value(TaxYearEnd)))
    Case HMIT_SECTIONS.HMIT_C
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_VOUCHERS_AND_CREDITCARDS_C
      Call HMITSectionHeader(rep, HMITS, p11d32.Rates.BenClassTo(BC_VOUCHERS_AND_CREDITCARDS_C, BCT_FORM_CAPTION))
      Call HMITColHeaders(rep, "Gross amount", "or amount foregone", "Amount made good", "or from which tax deducted", "Cash equivalent", "or relevant amount")
      Call HMITAssetsTransferredType(rep, ee, BenArr, "{x=8}Value of vouchers and payments made using credit cards or tokens" & vbCrLf & "{x=8}{Arial=6,ni} (for qualifying childcare vouchers read the P11D Guide)", False)
    Case HMIT_SECTIONS.HMIT_D
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_LIVING_ACCOMMODATION_D
      Call HMITSectionHeader(rep, HMITS, "Living accommodation")
      Call HMITColHeaders(rep, "", "", "", "", "Cash equivalent", "or relevant amount")
      Call HMITVanTypeNIC(rep, ee, BenArr, "{x=8}Cash equivalent or relevant amount of accommodation provided for the employee, or his/her family or household." & vbCrLf & "{x=8}Exceptions do not apply if using optional renumeration arrangements" & "{Arial=6,ni} (read P11D Guide)")
    Case HMIT_SECTIONS.HMIT_E
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_EMPLOYEE_CAR_E
      Call HMITSectionHeader(rep, HMITS, "Mileage allowance payments not taxed at source")
      Call HMITColHeaders(rep, "", "", "", "", "", "Taxable amount")
      Call HMITVanType(rep, ee, BenArr, "{x=8}Enter the mileage allowances in excess of the exempt amounts only where you have been unable to tax" & _
      vbCrLf & "{x=8}this under PAYE. The exceptions do not apply if using optional renumeration arrangements" & _
      "{Arial=6,ni}(read P11D Guide" & _
      ")")
    Case HMIT_SECTIONS.HMIT_F
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_COMPANY_CARS_F
      HMITSection = HMITCollections(rep, ee, lBenefitStartIndex, lBenefitsFound, ee.benefits, BenArr)
'      Call rep.Out("{x=3}{LINE=94}")
    Case HMIT_SECTIONS.HMIT_G
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_NONSHAREDVANS_G
      Call HMITSectionHeader(rep, HMITS, "Vans and van fuel" & vbCrLf)
      Call HMITVanTypeNIC(rep, ee, BenArr, "{x=8}Total cash equivalent or amount foregone in respect of all vans made available in " & p11d32.Rates.value(TaxFormYear), , , nsvans_benefit_van_only)
      Call HMITVanTypeNIC(rep, ee, BenArr, "{x=8}Total cash equivalent or amount foregone of fuel for all vans made available in " & p11d32.Rates.value(TaxFormYear), , , nsvans_fuel_benefit, p11d32.Rates.BenClassTo(BC_NONSHAREDVANS_FUEL_G, BCT_HMIT_BOX_NUMBER))
    Case HMIT_SECTIONS.HMIT_H
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_LOAN_OTHER_H
      'ReDim BenArr(1 To 2)'RK redundant 18/03/03
      'BenArr(1) = BC_LOAN_HOME_H
      'BenArr(2) = BC_LOAN_OTHER_H
      l = ee.GetLoansBenefitIndex
      If l > 0 Then
        Set loans = ee.benefits(l)
        If lBenefitStartIndex = 1 Then Call loans.SortLoans
        HMITSection = HMITCollections(rep, ee, lBenefitStartIndex, lBenefitsFound, loans.loans, BenArr)
      Else
        HMITSection = HMITCollections(rep, ee, lBenefitStartIndex, lBenefitsFound, Nothing, BenArr)
      End If
    Case HMIT_SECTIONS.HMIT_I
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_PRIVATE_MEDICAL_I
      Call HMITSectionHeader(rep, HMITS, "Private medical treatment or insurance")
      Call HMITColHeaders(rep, "Cost to you", "or amount foregone", "Amount made good", "or from which tax deducted", "Cash equivalent", "or relevant amount")
      Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Private medical treatment or insurance", False)
    Case HMIT_SECTIONS.HMIT_J
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_QUALIFYING_RELOCATION_J
      Call HMITSectionHeader(rep, HMITS, "Qualifying relocation expenses payments and benefits" & vbCrLf)
      Call rep.Out("{x=8}{Arial=8,ni}Non-qualifying benefits and expenses go in sections M and N below" & vbCrLf)
      Call HMITVanTypeNIC(rep, ee, BenArr, "Excess over " & FormatWN(L_RELOCEXEMPT) & " of all qualifying relocation expenses payments and benefits for each move", , , , p11d32.Rates.BenClassTo(BC_QUALIFYING_RELOCATION_J, BCT_HMIT_BOX_NUMBER))
    Case HMIT_SECTIONS.HMIT_K
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_SERVICES_PROVIDED_K
      Call HMITSectionHeader(rep, HMITS, "Services supplied")
      Call HMITColHeaders(rep, "Cost to you", "or amount foregone", "Amount made good", "or from which tax deducted", "Cash equivalent", "or relevant amount")
      Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Services supplied to the employee", False)
    Case HMIT_SECTIONS.HMIT_L
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_ASSETSATDISPOSAL_L
      Call HMITSectionL(rep, ee, BenArr())
      
'    Case HMIT_SECTIONS.HMIT_M
'      ReDim BenArr(1 To 1)
'      BenArr(1) = BC_SHARES_M
'      l = SumBenefitFWNRPT(ee, Description, value, MadeGood, benefit, BenArr)
'      Call rep.Out(FillBox(3, 3, L_HMIT_STANDARDBOX_HEIGHT, "M") & HMITText("{x=8}Shares") & vbCrLf & vbCrLf & _
'                   TickOutPlusX(L_HMIT_COL_1, l) & _
'                   "{Arial=7,n}{x=8}Tick the box if during the year there have been share-related benefits" & vbCrLf & _
'                   "{X=8}for the employee" & vbCrLf & vbCrLf)
    Case HMIT_SECTIONS.HMIT_M
      Call HMITSectionHeader(rep, HMIT_M, "Other items (including subscriptions and professional fees)")
      Call HMITColHeaders(rep, "Cost to you", "or amount foregone", "Amount made good", "or from which tax deducted", "Cash equivalent", "or relevant amount")
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_CLASS_1A_M
      Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr)
      Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Description of other items", True)
      'Call rep.Out(vbCrLf)
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_NON_CLASS_1A_M
      Call HMITAssetsTransferredType(rep, ee, BenArr, "{x=8}Description of other items", True)
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_INCOME_TAX_PAID_NOT_DEDUCTED_M
      Call HMITColHeaders(rep, "", "", "", "", "", "Tax Paid")
      Call HMITVanType(rep, ee, BenArr, "{x=8}Income Tax paid but not deducted from director's remuneration")
    Case HMIT_SECTIONS.HMIT_N
      Call HMITSectionHeader(rep, HMIT_N, "{x=8}Expenses payments made on behalf of the employee")
      Call HMITColHeaders(rep, "Cost to you", "or amount foregone", "Amount made good", "or from which tax deducted", "Taxable payment", "or relevant amount")
      ReDim BenArr(1 To 1)
      BenArr(1) = BC_TRAVEL_AND_SUBSISTENCE_N
      Call HMITAssetsTransferredType(rep, ee, BenArr, "{x=8}Travelling and subsistence payments - Cost to you or amount" & _
                  vbCrLf & "{x=8}foregone" & "{Arial=6,ni} (Except mileage allowance payments for employee's own car - see section E)", False)
      BenArr(1) = BC_ENTERTAINMENT_N
      'PS TTP#194
      entertainmenttick = ee.SumBenefit(tempDesc, tempvalue, tempmadegood, tempbenefit, BenArr())
      Call HMITAssetsTransferredType(rep, ee, BenArr, "{PUSH}" & TickOutPlusX(L_HMIT_COL_1 - 5, IIf(entertainmenttick, IIf(benEmployer.value(employer_CT_db), True, -2), 0)) & "{POP}{Arial=7,n}" & "Entertainment {Arial=7,i}(trading organisations read P11D Guide and~{Arial=7,n}{Arial=7,i}then enter a tick or a cross as appropriate here){Arial=7,n}" & "{Arial=7}", False)

      BenArr(1) = BC_GENERAL_EXPENSES_BUSINESS_N
      BenArr(1) = BC_PHONE_HOME_N
      Call HMITAssetsTransferredType(rep, ee, BenArr, "{x=8}Payments for use of home telephone", False)
      BenArr(1) = BC_NON_QUALIFYING_RELOCATION_N
        Call HMITAssetsTransferredType(rep, ee, BenArr, "{x=8}Non-qualifying relocation expenses " & "{Arial=6,ni}" & " (those not shown in sections J or M){Arial=7,b}", False)
      ReDim BenArr(1 To 2)
      BenArr(1) = BC_OOTHER_N
      BenArr(2) = BC_CHAUFFEUR_OTHERO_N
      Call HMITAssetsTransferredType(rep, ee, BenArr, "{x=8}Description of other expenses", True)
  End Select
  
  Call rep.Out("{ENDSECTION}")
  
HMITSection_END:
  
  Call xReturn("HMITSection")
  Exit Function
HMITSection_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITSection", "HMIT Section", "Error printing a section on the HMIT return, section index = " & HMITS & ".")
  Resume HMITSection_END
  Resume
End Function
Private Function HMITCollectionsSet(ben1 As IBenefitClass, ben2 As IBenefitClass, ben As IBenefitClass) As Boolean
  On Error GoTo HMITCollectionsSet_ERR
  
  Call xSet("HMITCollectionsSet")
  
  If ben1 Is Nothing Then
    Set ben1 = ben
    ' EK added 1/04 TTP#194
    If p11d32.AppYear = 2003 And ben.BenefitClass = BC_COMPANY_CARS_F Then
      HMITCollectionsSet = True
    End If
  Else
    HMITCollectionsSet = True
    Set ben2 = ben
  End If
        
HMITCollectionsSet_END:
  Call xReturn("HMITCollectionsSet")
  Exit Function
HMITCollectionsSet_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITCollectionsSet", "HMIT Collections Set", "Error setting valid benefits for HMIT collections.")
  Resume HMITCollectionsSet_END
End Function
Private Function HMITCollections(rep As Reporter, ee As Employee, lBenefitStartIndex, lCollectionItemsFound As Long, BenefitCollection As ObjectList, BenArr() As BEN_CLASS) As Boolean
  Dim i As Long, j As Long
  Dim ben As IBenefitClass
  Dim ben1 As IBenefitClass, ben2 As IBenefitClass, bendummy As IBenefitClass
  Dim bFirst As Boolean
  On Error GoTo HMITCollections_ERR
  
  Call xSet("HMITCollections")
    
  bFirst = lBenefitStartIndex
  
  If Not BenefitCollection Is Nothing Then
    For i = lBenefitStartIndex To BenefitCollection.count
      Set ben = BenefitCollection(i)
      If Not ben Is Nothing Then
        If BenClassInArray(BenArr, ben.BenefitClass) Then
          If BenefitIsLoan(ben.BenefitClass) Then
            If ben.value(ITEM_BENEFIT_REPORTABLE) Then
              HMITCollections = HMITCollectionsSet(ben1, ben2, ben)
              lBenefitStartIndex = i
            End If
          Else
            HMITCollections = HMITCollectionsSet(ben1, ben2, ben)
            lBenefitStartIndex = i
          End If
          If HMITCollections Then Exit For
        End If
      End If
    Next
  End If
  
  lBenefitStartIndex = lBenefitStartIndex + 1
  'do dependant things on the benclass
  If Not ben1 Is Nothing Or lCollectionItemsFound = 0 Then
    If BenClassInArray(BenArr, BC_COMPANY_CARS_F) Then
      
      ' EK 1/04 TTP#194
      If p11d32.AppYear = 2003 Then
        Call HMITCar(rep, ee, ben1, lCollectionItemsFound + 1, bendummy, 0)
      Else
        Call HMITCar(rep, ee, ben1, lCollectionItemsFound + 1, ben2, lCollectionItemsFound + 2)
      End If
    
    ElseIf BenClassInArray(BenArr, BC_LOAN_OTHER_H) Then
      Call HMITLoan(rep, ben1, lCollectionItemsFound + 1, ben2, lCollectionItemsFound + 2)
    End If
  End If
      
'  If p11d32.AppYear = 2003 And BenClassInArray(BenArr, BC_COMPANY_CARS_F) Then
 '    lCollectionItemsFound = lCollectionItemsFound + 1
  'Else
     lCollectionItemsFound = lCollectionItemsFound + 2
  'End If
  
  
HMITCollections_END:
  Call xReturn("HMITCollections")
  Exit Function
HMITCollections_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITCollections", "HMIT Collections", "Error in HMIT collections for HMIT print")
  Resume HMITCollections_END
  Resume
End Function
Public Function BenClassInArray(BenArr() As BEN_CLASS, bc As BEN_CLASS) As Boolean
  Dim i As Long
  
  On Error GoTo BenClassInArray_ERR
  
  Call xSet("BenClassInArray")
  
  If IsArray(BenArr) Then
    For i = LBound(BenArr) To UBound(BenArr)
      If bc = BenArr(i) Then
        BenClassInArray = True
        Exit For
      End If
    Next
  Else
    Call Err.Raise(ERR_NOT_ARRAY, "BenClassInArray", "The variable passed is not an array.")
  End If
  
BenClassInArray_END:
  Call xReturn("BenClassInArray")
  Exit Function
BenClassInArray_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BenClassInArray", "Ben Class In Array", "Error determining if benclass is in BenArr.")
  Resume BenClassInArray_END
  Resume
End Function
Public Function HMITDate(v As Variant) As String
  HMITDate = Format$(v, "dd mmmm yyyy")
End Function

Private Sub HMITLoan(rep As Reporter, Loan1 As IBenefitClass, lLoan1Number As Long, Loan2 As IBenefitClass, lLoan2Number As Long)
  On Error GoTo HMITLoan_ERR
  
  Call xSet("HMITLoan")
  
  Call HMITSectionHeader(rep, HMIT_H, "Interest-free and low interest loans" & vbCrLf & vbCrLf)
  rep.Out ("{Arial=7,ni}{x=8}If the total amount outstanding on all loans does not exceed " & FormatWN(L_LOANDEMINIMUS) & " at any time in the year, there is no need to complete this section" & vbCrLf & "{Arial=7,ni}{x=8}unless the load is provided under an optional renumeration arrangement when the threshold does not apply")
  
  'Number of borrowers
  Call rep.Out(vbCrLf & HMITColText("Loan " & CStr(lLoan1Number), L_HMIT_COL_2) & _
               HMITColText("Loan " & CStr(lLoan2Number), L_HMIT_COL_4) & vbCrLf & vbCrLf)
  
  Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH / 2, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_HMIT_NoOfBorrowers, "")) & _
               OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH / 2, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_HMIT_NoOfBorrowers, "")) & _
               "{x=8}{Arial=7,n}Number of joint borrowers {Arial=7,ni} (if applicable){Arial=7,n}" & vbCrLf & vbCrLf & vbCrLf)
               
  'Outstanding Start
'  If p11d32.AppYear = 2000 Then 'km
'    Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_OpenOutstanding)) & _
'                 OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_OpenOutstanding)) & _
'                 "{x=8}{Arial=7,n}Amount oustanding at " & HMITDate(p11d32.Rates.value(LastTaxYearEnd)) & " or at date loan was made if later" & vbCrLf & vbCrLf & vbCrLf)
'  Else
    Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_OpenOutstanding)) & _
                 OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_OpenOutstanding)) & _
                 "{x=8}{Arial=7,n}Amount outstanding at " & HMITDate(p11d32.Rates.value(LastTaxYearEnd)) & " or at date loan was made if later" & vbCrLf & vbCrLf & vbCrLf)
'  End If
  
  'Outstanding end
'  If p11d32.AppYear = 2000 Then 'km
'    Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_CloseOutstanding)) & _
'                 OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_CloseOutstanding)) & _
'                 "{x=8}{Arial=7,n}Amount oustanding at " & HMITDate(p11d32.Rates.value(TaxYearEnd)) & " or at date loan was discharged if earlier" & vbCrLf & vbCrLf & vbCrLf)
'  Else
    Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_CloseOutstanding)) & _
                 OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_CloseOutstanding)) & _
                 "{x=8}{Arial=7,n}Amount outstanding at " & HMITDate(p11d32.Rates.value(TaxYearEnd)) & " or at date loan was discharged if earlier" & vbCrLf & vbCrLf & vbCrLf)
'  End If
  
  'Max outstanding
  Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_MaxOutstandingAtAnyPoint)) & _
               OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_MaxOutstandingAtAnyPoint)) & _
               "{x=8}{Arial=7,n}Maximum amount outstanding at any time in the year" & vbCrLf & vbCrLf & vbCrLf)
              
  'Total interest
  Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_InterestPaid_db)) & _
               OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_InterestPaid_db)) & _
               "{x=8}{Arial=7,n}Total amount of interest paid by the borrower in " & p11d32.Rates.value(TaxFormYear) & "{Arial=6,i}     - enter 'NIL' if none was paid{Arial=7}" & vbCrLf & vbCrLf & vbCrLf)
               
  
  'Date loan was made
   Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(Loan1, ln_MadeDate)) & _
             OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(Loan2, ln_MadeDate)) & _
             "{x=8}{Arial=7,n}Date loan was made in " & p11d32.Rates.value(TaxFormYear) & " (if applicable)" & vbCrLf & vbCrLf & vbCrLf)
  'Date loan was discharged
   Call rep.Out(OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(Loan1, ln_DischargedDate)) & _
             OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItem(Loan2, ln_DischargedDate)) & _
             "{x=8}{Arial=7,n}Date loan was discharged in " & p11d32.Rates.value(TaxFormYear) & " (if applicable)" & vbCrLf & vbCrLf & vbCrLf)
 
  
  'cash equivalent on all loans
  
  Call rep.Out(FillBox(L_HMIT_COL_2 - 4, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BC_LOANS_H, BCT_HMIT_BOX_NUMBER)) & _
               FillBoxNIC(L_HMIT_COL_3 - 3, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
               OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan1, ln_Benefit)) & _
               FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BC_LOANS_H, BCT_HMIT_BOX_NUMBER)) & _
               FillBoxNIC(L_HMIT_COL_5, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
               OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, GetBenItemFWNRPT(Loan2, ln_Benefit)) & _
               "{x=8}{Arial=7,n}Cash equivalent or relevant amount of loans after deducting any interest paid" & vbCrLf & "{x=8}{Arial=7,n}by the borrower" & vbCrLf & vbCrLf)
  
  Call xReturn("HMITLoan")
  
HMITLoan_END:
  Call xReturn("HMITLoan")
  Exit Sub
HMITLoan_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITLoan", "HMIT Loan", "Error printing the loans section of the HMIT return.")
  Resume HMITLoan_END
  Resume
End Sub
Private Function LineText(ByVal sText As String, Optional ByVal bItalic As Boolean = False) As String
  Dim s() As String
  Dim sFont As String
  
  On Error GoTo LineText_ERR
  
  Call xSet("LineText")
  
  If bItalic Then
    sFont = S_WK_NORMAL_ITALIC_FONT
  Else
    sFont = S_WK_NORMAL_FONT
  End If
  
  sFont = sFont & "{x=8}"
  
  If Len(sText) Then
    If GetDelimitedValues(s, sText, , , "~") > 1 Then
      LineText = sFont & s(1) & vbCrLf & "{x=8}" & s(2) & vbCrLf & vbCrLf
    Else
      If bItalic Then
        LineText = sFont & sText & vbCrLf & vbCrLf & vbCrLf
      Else
        LineText = sFont & "{YREL=100}" & sText & "{YREL=-100}" & S_WK_LINE_BREAK_FONT & vbCrLf & vbCrLf & vbCrLf
      End If
    End If
  Else
    LineText = LineText & S_WK_LINE_BREAK_FONT & vbCrLf & vbCrLf & vbCrLf
  End If
  
LineText_END:
  Call xReturn("LineText")
  Exit Function
LineText_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LineText", "Line Text", "Error outputting line text.")
  Resume LineText_END
  Resume
End Function
Public Function HMITFooter(sP11DType As String, Optional ee As Employee, Optional bAppYear As Boolean = True, Optional bManual As Boolean = False) As String
  Dim sEEDetails As String
  Dim ben As IBenefitClass
  Dim s As String
  If Not (ee Is Nothing) Then
    Set ben = ee
    sEEDetails = " - " + ee.FullName & "(" + ben.value(ee_NINumber_db) & ")"
  End If
  HMITFooter = "{x=0}{Arial=8,i}" & sP11DType & IIf(bAppYear, "(" & p11d32.AppYear + 1 & ")", "") & IIf(bManual, "Man", "") & "(Substitute)(" & app.companyName & ")" & TimeStampReport & sEEDetails
End Function
Public Function Report_HMIT(rep As Reporter, ee As Employee) As Boolean
  On Error GoTo Report_HMIT_ERR
  Dim y1 As Long, y2 As Long
  
  Call xSet("Report_HMIT")
  
  'EK change to this as formerly never printing footer, (defaultreportindex = 4 is employee letter - email)
  If Not (p11d32.ReportPrint.Destination = REPD_FILE_HTML And p11d32.ReportPrint.ExportOption = EXPORT_HTML_INTEXP5) Then
    If Not (p11d32.ReportPrint.DefaultReportIndex = RPT_EMPLOYEE_LETTER_EMAIL And p11d32.ReportPrint.ExportOption = EXPORT_HTML_INTEXP5) Then
      rep.PageFooter = HMITFooter("P11D", ee)
    End If
  End If
  
  Call rep.Out("{BEGINSECTION}")
  Call HMITPageHeader(rep, ee)
  Call HMITEmpDetails(rep, ee)
  Call rep.Out("{ENDSECTION}")

  Call HMITWhichSection(rep, ee, p11d32.ReportPrint.HMITSections_PRINT)
  Report_HMIT = True
Report_HMIT_END:
  Call xReturn("Report_HMIT")
  Exit Function
Report_HMIT_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Report_HMIT", "Report HMIT", "Error printing the HMIT P11D return.")
  Resume Report_HMIT_END
  Resume
End Function
Public Function Report_PrintedEmployees(rep As Reporter, PrintedEmployees As ObjectList)
  Dim ben As IBenefitClass
  Dim bc As BEN_CLASS
  Dim j As Long
    
  Const L_COL_SURNAME As Long = 0
  Const L_COL_FIRSTNAME As Long = 33
  Const L_COL_TITLE As Long = 66
  Const L_COL_PNUM As Long = 99
  Const S_TITLEFORMAT As String = "{Arial=8,b}"
  'JN code Ummmm
    
On Error GoTo Report_PrintedEmployees_ERR

  Call xSet("Report_PrintedEmployees")
  
  Call ReportBanner(rep, "Printed Employees")
  
  bc = BC_EMPLOYEE
  
  Call rep.Out(vbCrLf)
  Call WKTblColXOffsets(L_COL_SURNAME, L_COL_FIRSTNAME, L_COL_TITLE, L_COL_PNUM)
  Call WKTblColFormats("n", "n", "n", "rn")
  Call WKTableHeadings(rep, S_TITLEFORMAT & S_SURNAME, S_TITLEFORMAT & S_FIRSTNAME, S_TITLEFORMAT & S_TITLE, S_TITLEFORMAT & S_PNUM)
          
  For j = 1 To PrintedEmployees.count
    Set ben = PrintedEmployees(j)
    If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Report_PrintedEmployees", "The employee is nothing.")
    Call WKTableRow(rep, ben.value(ee_Surname_db), ben.value(ee_Firstname_db), ben.value(ee_Title_db), ben.value(ee_PersonnelNumber_db))
  Next

  Call WKTableTotals(rep, "", "", "", "Total printed " & PrintedEmployees.count)
  
  Report_PrintedEmployees = True
  
Report_PrintedEmployees_END:
  Call xReturn("Report_PrintedEmployees")
  Exit Function

Report_PrintedEmployees_ERR:
  
  Call ErrorMessage(ERR_ERROR, Err, "Report_PrintedEmployees", "Printed Employees Report", "Error in Printed Employees Report")
  Resume Report_PrintedEmployees_END
  Resume
End Function
Public Function P11DbAdditionsDescription(benEmployer As IBenefitClass, Optional ByVal defaultValueIfNotUSerValueEntered As Boolean = True) As String
  P11DbAdditionsDescription = P11DbAdditionsDeductionsDescription(benEmployer, employer_NIC_AdjustmentAddTotal, employer_addClass1ADescription_db, S_P11DB_AMOUNTS_TAXED_THROUGH_PAYROLL)
End Function
Public Function P11DbDeductionsDescription(benEmployer As IBenefitClass) As String
  P11DbDeductionsDescription = P11DbAdditionsDeductionsDescription(benEmployer, employer_NIC_AdjustmentDeductTotal, employer_deductClass1ADescription_db, S_P11DB_EMPLOYEES_NOT_SUBJECT_TO_CLASS_1A)
End Function
Public Function P11DbAdditionsDeductionsDescription(benEmployer As IBenefitClass, ByVal employer_NIC_AdjustmentAdditionDeductTotalEnumValue As Long, ByVal employer_NIC_AdjustmentAdditionDeductTotalDescriptionEnumValue As Long, ByVal defaultDescription As String) As String
  If (benEmployer.value(employer_NIC_AdjustmentAdditionDeductTotalEnumValue) = 0) Then
    P11DbAdditionsDeductionsDescription = ""
  Else
    Dim descriptionEntered As String
    descriptionEntered = Trim$(benEmployer.value(employer_NIC_AdjustmentAdditionDeductTotalDescriptionEnumValue))
    If Len(descriptionEntered) = 0 Then
      P11DbAdditionsDeductionsDescription = defaultDescription
    Else
      P11DbAdditionsDeductionsDescription = descriptionEntered
    End If
  End If
End Function
Private Function Report_P11db_NIC_Rate_Formatted() As String
  Report_P11db_NIC_Rate_Formatted = (p11d32.Rates.value(carNICRate) * 100) & "%"
End Function
Public Function Report_P11db(rep As Reporter, benEmployer As IBenefitClass)
  Dim add1A As Long
  Dim deduct1A As Long
  Dim sNICPErcentage As String, sBenefitsPotentiallyWithClass1A As String
  Dim sNICDue As String
  Dim bPrintAdjustmentValues As Boolean
  Dim l As Long
  Dim ey As Employer
  
  On Error GoTo p11dbform_err
  
  Call xSet("p11dbform")
  
  Call benEmployer.Calculate
    
  sNICPErcentage = Report_P11db_NIC_Rate_Formatted()
  sNICDue = FormatWNRPT(benEmployer.value(ITEM_NIC_CLASS1A_BENEFIT), , , True)
  add1A = benEmployer.value(employer_NIC_AdjustmentAddTotal)
  deduct1A = benEmployer.value(employer_NIC_AdjustmentDeductTotal)
  bPrintAdjustmentValues = (add1A > 0) Or (deduct1A > 0)
  sBenefitsPotentiallyWithClass1A = FormatWNRPT(benEmployer.value(employer_TotalBenefitsPotentiallySubjectToClass1A))
  rep.PageFooter = HMITFooter(S_P11D_B, , , True)
  
  Call rep.Out("{PUSHY}")
  Call rep.Out(vbCrLf & "{Arial=13,b}{x=3}HM Revenue" & vbCrLf & _
               "{Arial=16,b}{x=3}&Customs" & vbCrLf & vbCrLf)
  Call rep.Out("{POP}")
  
  'if statement and "DRAFT" added by IK. 23/05/2003
  If p11d32.ReportPrint.DraftReports Then
    Call rep.Out("{x=20}{Arial=16,b} DRAFT")
  End If
  Call rep.Out("{x=40}{Arial=12,b} Return of Class 1A National Insurance contributions due" & vbCrLf & _
               "{Arial=12,b}{x=40}Return of expenses and benefits - Employer declaration " & vbCrLf & vbCrLf & _
               "{Arial=12,b}{x=70}" & "Year ended 5 April " & Year(p11d32.Rates.value(TaxYearEnd)) & vbCrLf & vbCrLf)
  
  'KA: Employer reference
  
  Call rep.Out(OutLineBoxL(16, 28, L_HMIT_STANDARDBOX_HEIGHT, benEmployer.value(employer_Payeref_db)) & _
                         "{Arial=8,n}{x=4}Employer PAYE" & vbCrLf & _
                         "{Arial=8,n}{x=4}reference" & vbCrLf & vbCrLf)
  
  'KA: Accounts Office reference
  
  Call rep.Out(OutLineBoxL(16, 28, L_HMIT_STANDARDBOX_HEIGHT, benEmployer.value(employer_IRAccountsOfficeReference_db)) & _
                         "{Arial=8,n}{x=4}Accounts office" & "{x=54}" & vbCrLf & _
                         "{Arial=8,n}{x=4}reference" & vbCrLf & vbCrLf)
  
  
  'KA: Addresses
  
  Call rep.Out("{Arial=8,nb}{x=4}Employer name and address" & _
              "{Arial=8,nb}{x=54}Please return this form to the address shown below" & vbCrLf & _
               OutLineBoxL(54, 45, 8, benEmployer.value(employer_IRTaxOffice_db) & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf) & _
               OutLineBoxL(4, 45, 8, benEmployer.Name & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf) & vbCrLf & _
               "{x=54}" & benEmployer.value(employer_IRAddressLine1_db) & "{x=4}" & benEmployer.value(employer_AddressLine1_db) & vbCrLf & _
               "{x=54}" & benEmployer.value(employer_IRAddressLine2_db) & "{x=4}" & benEmployer.value(employer_AddressLine2_db) & vbCrLf & _
               "{x=54}" & benEmployer.value(employer_IRAddressLine3_db) & "{x=4}" & benEmployer.value(employer_AddressLine3_db) & vbCrLf & _
               "{x=54}" & benEmployer.value(employer_IRAddressLine4_db) & "{x=4}" & benEmployer.value(employer_AddressLine4_db) & vbCrLf & _
               "{x=54}" & benEmployer.value(employer_IRPostcode_db) & "{x=4}" & benEmployer.value(employer_AddressPostCode_db) & vbCrLf & vbCrLf)
       
  
    
    Call rep.Out("{Arial=9,nb}{x=4}If this replaces a return that was issued automatically it may not show all of your details. If so, fill in the top" & _
                vbCrLf & "{Arial=9,nb}{x=4}of this return before you send it to your HM Revenue & Customs (HMRC) office." & _
                vbCrLf & vbCrLf & "{Arial=9,nb}{x=4}Please read the notes overleaf before completing this return." & _
                 vbCrLf & "{Arial=9,nb}{x=4}Don't declare any amounts already reported under the Taxed Award Scheme arrangements." & vbCrLf & vbCrLf)
    
    
    '******************* SECTION 1 CAD ammended 2006 and I added these comments to help !!! ***************************
    
    'BOX A all the benefits potentiall subject to c1A
    Call rep.Out("{x=4}{BOX=95,17.8}")
    Call BoxBulletText(rep, "1", "{Arial=10,bn}Class 1A National Insurance contributions (NICs) due" & vbCrLf, 4)
    rep.Out (OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, sBenefitsPotentiallyWithClass1A) & _
             FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "A") & _
             FillBoxNIC(L_HMIT_COL_5 + 1, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
             "{Arial=9,n}{x=7}Enter the " & S_P11DB_SHORT_CAPTION_BOX_A & ", (this is the total of the Class 1A" & vbCrLf & _
             "{Arial=9,n}{x=7}NICs boxes on forms P11D) and/or the total benefits that have been taxed through your payroll." & vbCrLf & _
             "{Arial=9,n}{x=7}There's a quick guide to working out whether Class 1A NICs are due in Part 2 of the CWG5" & vbCrLf & _
             "{x=7}if you're not sure" & vbCrLf & vbCrLf)
             
    Call rep.Out("{Arial=9,nb}{x=18}Please note: if you need to adjust the figures entered in box A, don't complete box C below," & vbCrLf & _
             "{x=9}" & TickOut(add1A > 0 Or deduct1A > 0) & _
             "{Arial=9,nb}{x=18}tick this box and complete Section 4 overleaf." & vbCrLf & vbCrLf)  'km - added Carriage Return
    
    
    'box b + c
    Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, sNICPErcentage) & _
                FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "B") & _
                "{Arial=9,n}{x=7}" & S_P11DB_MUTIPLY_BY_NIC_RATE_BOX_E & vbCrLf & vbCrLf & _
                "{Arial=7,n}{x=79}box A x rate in box B" & vbCrLf)
    Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, IIf(bPrintAdjustmentValues, "", sNICDue)))
    
    
    Call rep.Out(FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "C") & _
                FillBoxNIC(L_HMIT_COL_5 + 1, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
                "{Arial=9,nb}{x=7}" & S_P11DB_CLASS_1A_PAYABLE_BOX_F & "{Arial=9,n} (Don't fill this in if you're making an adjustment in Section 4.)" & vbCrLf)
  
  
    
  ' *********************** SECTION 2 ************************************************************8
  Call rep.Out(vbCrLf & vbCrLf & "{x=4}{BOX=95,36}")
  Call BoxBulletText(rep, "2", "{Arial=10,bn}Employer declaration" & vbCrLf & vbCrLf, 4)
    
    Call rep.Out("{Arial=9,in}{x=5}Tick the relevant box and fill in the appropriate details." & vbCrLf & vbCrLf & _
                 "{x=5}" & TickOut(benEmployer.value(employer_EerDeclSect2Chk1_db)) & "{Arial=9,n}{yrel=-10}{x=9}No expenses payments or benefits that must be returned on forms P11D or taxed through payroll have been or will be" & vbCrLf & _
                 "{x=9}provided for the year ended 5 April " & Year(p11d32.Rates.value(TaxYearEnd)) & vbCrLf & vbCrLf & _
                 "{x=5}" & TickOut(benEmployer.value(employer_EerDeclSect2Chk2_db)) & "{Arial=9,n}{yrel=-10}{x=9}I confirm that all details of expenses payments and benefits that must be returned on forms P11D for the year ending" & vbCrLf & _
                 "{x=9}5 April " & Year(p11d32.Rates.value(TaxYearEnd)) & "{Arial=9,b} are enclosed " & "{Arial=9,n}with this declaration. I confirm that I have declared any Class 1A NICs due for expenses" & vbCrLf & "{x=9}payments and benefits that have been taxed through payroll. I declare that the details provided are fully and truly" & vbCrLf & _
                 "{x=9}stated to the best of my knowledge and belief." & vbCrLf & vbCrLf)
    
    
    Call rep.Out("{x=5}" & TickOut(benEmployer.value(employer_EerDeclSect2Chk3_db)) & OutLineBoxL(35, 23, L_HMIT_STANDARDBOX_HEIGHT, benEmployer.value(employer_IRTaxOffice_db)) & OutLineBoxL(L_HMIT_COL_3 + 1, L_HMIT_STANDARDBOX_WIDTH + 4, L_HMIT_STANDARDBOX_HEIGHT, "") & _
                 "{Arial=9,n}{x=9}{yrel=0}Forms P11D for the year ended" & "{x=59}HMRC office on" & vbCrLf & _
                 "{Arial=9,n}{x=9}{yrel=0}5 April " & Year(p11d32.Rates.value(TaxYearEnd)) & " were sent to{x=59}" & vbCrLf & vbCrLf & vbCrLf)
    Call rep.Out("{x=5}I confirm that details of expenses payments and benefits that must be returned on forms P11D or taxed through payroll" & vbCrLf & _
                 "{x=5}have been sent to HMRC." & vbCrLf & vbCrLf & _
                 "{Arial=9,nb}{x=5}I declare that all the details on this form are fully and truly stated to the best of my knowledge and belief." & vbCrLf & vbCrLf)
  
  Call rep.Out(OutLineBoxL(26, 43, L_HMIT_SIGNATORYBOX_HEIGHT, "") & vbCrLf & OutLineBoxL(L_HMIT_COL_3 + 1, L_HMIT_STANDARDBOX_WIDTH + 4, L_HMIT_STANDARDBOX_HEIGHT, "") & _
               "{x=5}{Arial=9,bn}Signature of employer" & _
               "{x=73}{Arial=9,bn}Date" & vbCrLf & vbCrLf & vbCrLf & _
               "{Arial=9,ni}{x=5}The declaration should be signed by the employer or any person authorised to do so." & vbCrLf & vbCrLf & _
               OutLineBoxL(26, 43, L_HMIT_SIGNATORYBOX_HEIGHT, "") & vbCrLf & _
               "{x=5}{Arial=9,bn}Capacity in which signed" & vbCrLf & vbCrLf)
           
  
    
  'page 2
  Call rep.Out("{NEWPAGE}")
  
  '***************************** SECTION 3 ***********************************************
    



  Call rep.Out(vbCrLf & "{x=4}{BOX=95,37}")
    Call BoxBulletText(rep, "3", "{Arial=10,bn}Notes for employer" & vbCrLf & vbCrLf, 4)
    
    Call rep.Out("{Arial=9,nb}{x=5}Class 1A National Insurance contributions (NICs) due" & vbCrLf)
    Call rep.Out("{Arial=9,n}{x=5}You need to pay Class 1A NICs on taxable expenses and benefits, unless Class 1 or Class 1B NICs are due. The boxes" & vbCrLf)
    Call rep.Out("{x=5}marked '1A' on the P11D indicate that you need to pay Class 1A NICs but you also need to include benefits you've" & vbCrLf)
    Call rep.Out("{x=5}payrolled, and not included on form P11D. You can find more information in booklet CWG5 'Class 1A National Insurance" & vbCrLf)
    Call rep.Out("{x=5}contributions on benfits in kind, A guide for employers'." & vbCrLf & vbCrLf)
    
    
      Call rep.Out("{x=5}You need to pay Class 1A NICs shown on the return to the Accounts Office. For details on how to pay, go to" & vbCrLf)
      Call rep.Out("{x=5}www.gov.uk/pay-class-1a-national-insurance" & vbCrLf)
      
      Call rep.Out("{x=5}Your payment must reach us by:" & vbCrLf & vbCrLf & _
                HMITBullet(7) & "{Arial=9,n}{x=9}19 July{Arial=9,n} if paying is by post" & vbCrLf & _
                HMITBullet(7) & "{Arial=9,n}{x=9}22 July{Arial=9,n} if paying by an approved electronic method" & vbCrLf & vbCrLf)
                
      Call rep.Out("{Arial=9,n}{x=5}Please, note that if 22 July falls on a non-banking day, you'll need to pay early unless you are using Faster Payments." & vbCrLf & _
                "{Arial=9,n}{x=5}There's more information on our webiste. We charge interest on late payments." & vbCrLf & _
                "{x=5}The filing deadline is 6 July. If we've not received your return by 19 July, we'll charge penalties. The amount we charge" & vbCrLf & _
                "{x=5}is £100 for each month or part month the return is outstanding, for each 50 employees or part batch of 50." & vbCrLf & vbCrLf)
                    
      Call rep.Out("{Arial=9,nb}{x=5}P11D Forms" & vbCrLf)
      Call rep.Out("{Arial=9,n}{x=5}You must complete a P11D for each employee or director who receives taxable expenses or benefits from you, or" & vbCrLf)
      Call rep.Out("{x=5}from a thrird party by your arrangement unless:" & vbCrLf)
      
      Call rep.Out(HMITBullet(7) & "{Arial=9,n}{x=9}you registered online before the start of the tax year to payroll all taxable expenses or benefits, and you've" & vbCrLf)
      Call rep.Out("{Arial=9,n}{x=9}taxed them in full" & vbCrLf)
      Call rep.Out(HMITBullet(7) & "{Arial=9,n}{x=9}the expenses were covered by an exemption, or an agreed bespoke benchmark rate" & vbCrLf)
      Call rep.Out(HMITBullet(7) & "{Arial=9,n}{x=9}You've arranged a PAYE Settlement Agreement with us" & vbCrLf & vbCrLf)
      
      
      Call rep.Out("{Arial=9,n}{x=5}Whether you taxed benefits through your payroll or not, you need to give your employees a copy of the" & vbCrLf)
      Call rep.Out("{Arial=9,n}{x=5}information you've reported on a P11D or your Full Payment Submissions on or before 6 July, so that they can" & vbCrLf)
      Call rep.Out("{Arial=9,n}{x=5}complete a tax return if they get one." & vbCrLf & vbCrLf)
      
                        
      

'************************* SECTION 4 ADJUSTMENTS *************************************


Call rep.Out(vbCrLf & "{x=4}{BOX=95,42}")
Call BoxBulletText(rep, "4", "{Arial=10,bn}Adjustments to Class 1A NICs" & vbCrLf & vbCrLf, 4)
  
  Call rep.Out("{Arial=9,nb}{x=5}Complete this section if you need to adjust the total benefits shown as liable to Class 1A NICs." & vbCrLf & _
                "{Arial=9,n}{x=5}Paragraph 18 of CWG5 explains circumstances in which you may need to make adjustments." & vbCrLf & vbCrLf & vbCrLf)
  
  
  Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, IIf(bPrintAdjustmentValues, sBenefitsPotentiallyWithClass1A, "")) & _
           FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "A") & _
           FillBoxNIC(L_HMIT_COL_5 + 1, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
           "{Arial=9,n}{x=7}Enter the total benefits liable to Class 1A NICs from Section 1, box A overleaf." & vbCrLf & vbCrLf & vbCrLf)  'km

'box A
Call rep.Out(HMITBullet(7) & "{Arial=9,n}{x=9}" & S_P11DB_ADDITIONS_BOX_B & vbCrLf & _
            "{Arial=7,n}{x=79}Amount to be added" & vbCrLf & _
            OutLineBoxL(25, 50, L_HMIT_STANDARDBOX_HEIGHT, P11DbAdditionsDescription(benEmployer)) & _
            OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, FormatWNRPT(add1A)) & _
            FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "B") & _
            FillBoxNIC(L_HMIT_COL_5 + 1, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
            "{Arial=9,n}{x=9}Brief description" & vbCrLf & vbCrLf & vbCrLf)
            
Call rep.Out(HMITBullet(7) & "{Arial=9,n}{x=9}" & S_P11DB_DEDUCTIONS_BOX_C & vbCrLf & _
            "{Arial=7,n}{x=79}Amount to be deducted" & vbCrLf & _
            OutLineBoxL(25, 50, L_HMIT_STANDARDBOX_HEIGHT, P11DbDeductionsDescription(benEmployer)) & _
            OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, IIf(bPrintAdjustmentValues, FormatWNRPT(deduct1A), "")) & _
            FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "C") & _
            "{Arial=9,n}{x=9}Brief description" & vbCrLf & vbCrLf & vbCrLf)
            
            'benefit1A + add1A - deduct1A
Call rep.Out("{Arial=7,n}{x=79}box A + box B minus box C" & vbCrLf & _
             OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, IIf(bPrintAdjustmentValues, FormatWNRPT(benEmployer.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A)), "")) & _
             FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "D") & _
             FillBoxNIC(L_HMIT_COL_5 + 1, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
             "{Arial=9,nb}{x=7}" & S_P11DB_TOTAL_BOX_D & vbCrLf & vbCrLf & vbCrLf & vbCrLf)
          

  Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, sNICPErcentage) & _
                FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "E") & _
                "{Arial=9,n}{x=7}" & S_P11DB_MUTIPLY_BY_NIC_RATE_BOX_E & vbCrLf & vbCrLf & vbCrLf)
  Call rep.Out("{Arial=7,n}{x=79}box D x rate in box E" & vbCrLf & _
                OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH + 1, L_HMIT_STANDARDBOX_HEIGHT, IIf(bPrintAdjustmentValues, sNICDue, "")) & _
                FillBox(L_HMIT_COL_3 + 1, 3, L_HMIT_STANDARDBOX_HEIGHT, "F") & _
                FillBoxNIC(L_HMIT_COL_5 + 1, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
                "{Arial=9,nb}{x=7}" & S_P11DB_CLASS_1A_PAYABLE_BOX_F & vbCrLf & vbCrLf & vbCrLf)

Report_P11db = True

p11dbform_end:
  Call xReturn("p11dbform")
  Exit Function
p11dbform_err:
  Call ErrorMessage(ERR_ERROR, Err, "P11DBForm", "ERR_UNDEFINED", "Undefined error.")
  Resume p11dbform_end
  Resume
End Function

Private Sub DoReport_WKSub(rep As Reporter, ben As IBenefitClass, ByVal CurBenClass As BEN_CLASS, bFirst As Boolean, BenOtherPrinted() As Boolean, ee As Employee)

  If ben.BenefitClass = CurBenClass Then
    If Not bFirst Then
      Call WKMainHeader(rep, ben, ee)
      bFirst = True
    End If
    If IsBenOtherClass(ben.BenefitClass) Then
      If Not BenOtherPrinted(ben.BenefitClass) Then
        Call ben.PrintWk(rep)
        BenOtherPrinted(ben.BenefitClass) = True
      End If
    Else
      Call ben.PrintWk(rep)
    End If
  End If
End Sub

Public Function Report_WK(rep As Reporter, ee As Employee) As Boolean
  Dim i As Long, j As Long, k As Long
  Dim RelevantBenClasses() As BEN_CLASS, BenClassCount As Long, CurBenClass As BEN_CLASS
  Dim BenOtherPrinted() As Boolean, LoansCol As loans
  Dim ben As IBenefitClass
  Dim bFirst As Boolean
  Dim bEmployeeCarCall As Boolean
  Dim orderedBenfits As ObjectList
  
  bEmployeeCarCall = True
    
  On Error GoTo Report_WK_ERR
  
  Call xSet("Report_WK")
  ReDim Preserve BenOtherPrinted(BC_FIRST_ITEM To BC_UDM_BENEFITS_LAST_ITEM)

  For i = BC_FIRST_ITEM To BC_UDM_BENEFITS_LAST_ITEM
    BenOtherPrinted(i) = Not IsBenOtherClass(i)
  Next i
    
  BenClassCount = ee.GetRelevantBenClassesSorted(RelevantBenClasses)
  For i = 1 To BenClassCount
    CurBenClass = RelevantBenClasses(i)
    For j = 1 To ee.benefits.count
      
      Set ben = ee.benefits(j)
      If Not ben Is Nothing Then
        If ben.BenefitClass = BC_LOANS_H Then
          Set LoansCol = ben
          For k = 1 To LoansCol.loans.count
            Set ben = LoansCol.loans(k)
            If Not ben Is Nothing Then Call DoReport_WKSub(rep, ben, CurBenClass, bFirst, BenOtherPrinted, ee)
          Next k
        Else
          If bEmployeeCarCall = True Or Not ben.BenefitClass = BC_EMPLOYEE_CAR_E Then
            Call DoReport_WKSub(rep, ben, CurBenClass, bFirst, BenOtherPrinted, ee)
          End If
          If ben.BenefitClass = BC_EMPLOYEE_CAR_E Then
            bEmployeeCarCall = False
          End If
        End If
      End If
    Next j
  bEmployeeCarCall = True
    
  Next i
  
  
Report_WK_END:
  Call xReturn("Report_WK")
  Exit Function
  
Report_WK_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Report_WK", "Report HMIT", "Error printing the working papers.")
  Resume Report_WK_END
  Resume
End Function
Private Function P46PaymentFrequencyTickOut(ByVal ben As IBenefitClass, ByVal ppf As P46_PAYMENT_FREQUENCY) As String
  P46PaymentFrequencyTickOut = TickOut(P46PaymentFrequencyEx(ben, ppf))
End Function
Private Function P46PaymentFrequencyEx(ByVal ben As IBenefitClass, ByVal ppf As P46_PAYMENT_FREQUENCY) As Boolean
  Dim b As Boolean
  
  If ben.value(car_P46WithdrawnWithoutReplacement) Or ben.value(car_MadeGood_db) = 0 Then
    b = False
  Else
    b = ben.value(car_p46PaymentFrequency_db) = ppf
    
  End If
  P46PaymentFrequencyEx = b
End Function
Private Sub RepOutCrLf(rep As Reporter, formattedText As String, Optional crLFXOffset As Long = -1)
  Dim s As String
  Dim sReplace As String
    
  sReplace = vbCrLf
  If (crLFXOffset <> -1) Then
    sReplace = sReplace & "{x=" & crLFXOffset & "}"
  End If
  
  s = Replace$(formattedText, "\n", sReplace)
  Call rep.Out(s)
End Sub
Private Sub P46CaptionCol2(rep As Reporter, caption As String, Optional Font As String = "{Arial=10,n}")
  Call P46Caption(rep, caption, L_P46_COL_2_X + 1, Font)
End Sub
Private Sub P46CaptionCol1(rep As Reporter, caption As String, Optional Font As String = "{Arial=10,n}")
  Call P46Caption(rep, caption, L_P46_COL_1_X + 1, Font)
End Sub
Private Sub P46TickRowCol1(rep As Reporter, caption As String, value As Boolean)
  Call P46TickRow(rep, L_P46_COL_1_X + 1, caption, value)
End Sub

Private Sub P46TickRowCol2(rep As Reporter, caption As String, value As Boolean)
  Call P46TickRow(rep, L_P46_COL_2_X + 1, caption, value)
End Sub
Private Sub P46TickRow(rep As Reporter, xoffset As Long, caption As String, value As Boolean)
  Call rep.Out("{x=" & (xoffset + 41) & "}")
  rep.Out (TickOut(value))
  Call rep.Out("{x=" & (xoffset) & "}")
  Call P46Caption(rep, caption, xoffset)
End Sub

Private Sub P46Caption(rep As Reporter, caption As String, xoffset As Long, Optional Font As String = "{Arial=10,n}")
  If (Len(caption) > 0) Then
    Call RepOutCrLf(rep, Font & "{x=" & xoffset & "}" & caption & vbCrLf & "{Arial=6,n}" & vbCrLf, xoffset)
  End If
End Sub
Private Sub P46InputBoxFullColumnLength(rep As Reporter, caption As String, value As String, xoffset As Long, Optional height As Single = L_HMIT_STANDARDBOX_HEIGHT)
  Call P46Caption(rep, caption, xoffset)
  
  Call rep.Out(OutLineBoxL(xoffset, 45, height, value))
  Call rep.Out("{Arial=10,n}" & vbCrLf & "{Arial=6,n}" & vbCrLf & "{Arial=10,n}")
End Sub
Private Sub P46InputBoxFullColumnLengthCol1(rep As Reporter, caption As String, value As String, Optional height As Single = L_HMIT_STANDARDBOX_HEIGHT)
  Call P46InputBoxFullColumnLength(rep, caption, value, L_P46_COL_1_X + 1, height)
End Sub
Private Sub P46InputBoxFullColumnLengthCol2(rep As Reporter, caption As String, value As String, Optional height As Single = L_HMIT_STANDARDBOX_HEIGHT)
  Call P46InputBoxFullColumnLength(rep, caption, value, L_P46_COL_2_X + 1, height)
End Sub
Private Sub P46BackgroundBox(rep As Reporter, Title As String, xoffset As Long, width As Long, height As Double)
  'Call RepOutCrLf(rep, "{x=" & xoffset & "}{FillRGB=15790320}{BOX=" & width & "," & height & ", F}{FillRGB=" & RGB(255, 255, 255) & "}")
  Call RepOutCrLf(rep, "{Arial=6,n}\n{x=" & xoffset & "}{BOX=" & width & "," & height & "}")
  Call RepOutCrLf(rep, "{Arial=6,n}\n")
  Call RepOutCrLf(rep, "{Arial=11,bn}{x=" & (xoffset + 1) & "}" & Title & "\n{Arial=6,n}\n{Arial=10,n}")
End Sub

Private Sub P46BackgroundBoxCol2(rep As Reporter, Title As String, height As Double)
  Call P46BackgroundBox(rep, Title, L_P46_COL_2_X, L_P46_BACKGROUND_COL_WIDTH, height)
End Sub
Private Sub P46BackgroundBoxCol1(rep As Reporter, Title As String, height As Double)
  Call P46BackgroundBox(rep, Title, L_P46_COL_1_X, L_P46_BACKGROUND_COL_WIDTH, height)
End Sub
Private Sub P464TickOut(rep As Reporter, xoffset As Long, caption1 As String, value1 As Boolean, caption2 As String, value2 As Boolean, caption3 As String, value3 As Boolean, caption4 As String, value4 As Boolean)
  Call P462TickOut(rep, xoffset, caption1, value1, caption2, value2)
  Call P462TickOut(rep, xoffset, caption3, value3, caption4, value4)
End Sub
Private Sub P462TickOut(rep As Reporter, xoffset As Long, caption1 As String, value1 As Boolean, caption2 As String, value2 As Boolean)
  Call rep.Out("{x=" & (xoffset + 16) & "}" & TickOut(value1))
  Call rep.Out("{x=" & (xoffset + 43) & "}" & TickOut(value2))
  Call P46Caption(rep, caption1 & "{x=" & (xoffset + 21) & "}" & caption2, xoffset + 1)
End Sub
Private Sub P46FuelTypeLine(rep As Reporter, caption As String, fuelChar As String)
  Call RepOutCrLf(rep, "{x=" & (L_P46_COL_1_X + 43) & "}" & " {Arial=10,b}" & fuelChar & "{Arial=10,n}" & HMITBullet(L_P46_COL_1_X + 1))
  Call P46CaptionCol1(rep, "  " & caption)
End Sub

Public Function Report_P46CarBeforeApril2018(rep As Reporter, ee As Employee, dDateFrom As Date, dDateTo As Date) As Boolean
  Dim P46Cars As ObjectList
  Dim p46car As IBenefitClass
  Dim i As Long, ben As IBenefitClass
  Dim benEmployer As IBenefitClass
  Dim bCarDetails  As Boolean, bFuelNoMadeGood As Boolean
  On Error GoTo Report_P46Car_err
  Call xSet("Report_P46Car")
  
  Set benEmployer = p11d32.CurrentEmployer
   
  If ee Is Nothing Then Call Err.Raise(ERR_NO_EMPLOYEE, "Report_P46Car", "The employee is nothing can not print the P46 car return.")
  If Not ee.GetP46Cars(P46Cars, dDateFrom, dDateTo) Then GoTo Report_P46Car_end

  'IK 17/06/2003 getting the total number of cars. Search lNumCars to see usage
  Dim CompanyCars As ObjectList
  Set CompanyCars = BenefitsOfType(ee, BC_COMPANY_CARS_F)
  Dim lNumCars As Long
  lNumCars = CompanyCars.count
  Set CompanyCars = Nothing
  
  rep.PageFooter = HMITFooter("P46(Car)(" & p11d32.AppYear & ")", ee, False)

  Set ben = ee
  
  For i = 1 To P46Cars.count
    Set p46car = P46Cars(i)
    With p46car
              
         Call rep.Out("{BEGINSECTION}")
        
         Call rep.Out(vbCrLf & "{x=51}{Arial=14,b}Car provided for the private use" & vbCrLf & _
                    "{Arial=13,b}{x=3}HM Revenue" & "{Arial=14,b}{x=53}of an employee or a director" & vbCrLf & _
                    "{Arial=16,b}{x=3}&Customs" & vbCrLf)
   

        If p11d32.ReportPrint.DraftReportsp46 Then Call rep.Out("{Arial=16,b}{x=30}DRAFT" & vbCrLf)
          
          
        Call RepOutCrLf(rep, "\n{x=2}{Arial=14,b}Use from 6 April 2014 onwards\n")

        Call rep.Out("{PUSHY}")
  
         
          Call RepOutCrLf(rep, "{Arial=10,n}\nYou must complete this form if there is a change that affects\ncar benefits for an employee earning at " & _
                           "the rate of £8,500\na year or more or a director for whom a car is made available\nfor private use." & _
                           "Complete and return this form within 28 days\nof the end of the quarter to " & Format$(p11d32.Rates.value(P46Quarter1End), "d mmmm") & ", " & Format$(p11d32.Rates.value(P46Quarter2End), "d mmmm") & ", " & Format$(p11d32.Rates.value(P46Quarter3End), "d mmmm") & _
                           " or\n" & Format$(p11d32.Rates.value(P46Quarter4End), "d mmmm") & " in which the change takes place.\n\n", 2)
          
          Call P46BackgroundBoxCol1(rep, "Employer's details", 16)
          
          Call P46InputBoxFullColumnLengthCol1(rep, "Name", benEmployer.Name)
          Call P46InputBoxFullColumnLengthCol1(rep, "Phone number", benEmployer.value(employer_contactnumber_db))
          Call P46InputBoxFullColumnLengthCol1(rep, "PAYE reference", benEmployer.value(employer_Payeref_db))
          
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")
          
          Call P46BackgroundBoxCol1(rep, "Employee's or Director's details", 31.57)
          Call P46InputBoxFullColumnLengthCol1(rep, "Name", ee.FullName)
          Call P46InputBoxFullColumnLengthCol1(rep, "National Insurance number", ben.value(ee_NINumber_db))
          Call P46InputBoxFullColumnLengthCol1(rep, "Date of birth (if known) DD MM YYYY", DateValReadToScreenOnlyValidDates(ben.value(ee_DOB_db)))
          Call P46CaptionCol1(rep, "Gender")
          Call RepOutCrLf(rep, "{x=" & (L_P46_COL_1_X + 1) & "}{Arial=10,n}Male ")
          Call rep.Out(TickOut(ben.value(ee_Gender_db) = S_GENDER_MALE))
          Call RepOutCrLf(rep, "{x=" & (L_P46_COL_1_X + 15) & "}{Arial=10,n}Female ")
          Call rep.Out(TickOut(ben.value(ee_Gender_db) = S_GENDER_FEMALE))
          
          Call rep.Out("{POP}")
          
          Call P46BackgroundBoxCol2(rep, "General details", 58)
          Call P46CaptionCol2(rep, "Show here and on Page 2 any changes that have\nbeen made", "{Arial=10,b}\n")
          Call P46TickRowCol2(rep, "We provided the employee or director with a car\nwhich is available for private use.", p46car.value(car_P46FirstProvidedWithCar))
          
          'CAD 2010, will be back in for 2011 PAYEonline!
          If (p11d32.ReportPrint.P46PrintReplacedP46s) And p46car.value(car_P46CarProvidedReplaced) Then
           Call P46TickRowCol2(rep, "We replaced a car provided to the employee or\ndirector with another car which is available for\nprivate use.", p46car.value(car_P46CarProvidedReplaced))
           Call P46CaptionCol2(rep, "If the employee has more than one car available for private\nuse, please give details of the car that you replaced.", "{Arial=10,i}")
           Call P46InputBoxFullColumnLengthCol2(rep, "Make and model", IIf(.value(car_p46ReplacementForMultipleCarsMakeAndModel), .value(car_CarReplacedMake_db), ""))
           Call P46InputBoxFullColumnLengthCol2(rep, "", IIf(.value(car_p46ReplacementForMultipleCarsMakeAndModel), .value(car_CarReplacedModel_db), ""))
           Call P46InputBoxFullColumnLengthCol2(rep, "Engine size", IIf(.value(car_p46ReplacementForMultipleCarsMakeAndModel) And .value(car_CarReplacedEngineSize_db) <> 0, .value(car_CarReplacedEngineSize_db) & " cc", ""))
          End If
          
          
          
          Call P46TickRowCol2(rep, "We provided the employee or director with a second\nor further car, which is available for private use.", p46car.value(car_P46SecondCar))
          Call P46TickRowCol2(rep, "The employee has started to earn at a rate of\n£8,500 a year or more, or has become a director.", False)
          Call P46TickRowCol2(rep, "We have withdrawn a car provided to the employee\nor director and have not replaced it.", p46car.value(car_P46WithdrawnWithoutReplacement))
          
          Call P46CaptionCol2(rep, "If you ticked this box, please complete the boxes below,\nand then go straight to the declaration overleaf. Do not\ncomplete the other sections.", "{Arial=10,i}")
          Call P46InputBoxFullColumnLengthCol2(rep, "Date withdrawn DD MM YYYY", IIf(p46car.value(car_P46WithdrawnWithoutReplacement), .value(Car_AvailableTo_db), ""))
          Call P46CaptionCol2(rep, "Please give details of the car withdrawn.")
          Call P46InputBoxFullColumnLengthCol2(rep, "Make and model", IIf(.value(car_P46WithdrawnWithoutReplacement), .value(car_Make_db), ""))
          Call P46InputBoxFullColumnLengthCol2(rep, "", IIf(p46car.value(car_P46WithdrawnWithoutReplacement), .value(car_Model_db), ""))
          
          Call P46InputBoxFullColumnLengthCol2(rep, "Engine size", IIf(.value(car_P46WithdrawnWithoutReplacement) And .value(car_enginesize_db) <> 0, .value(car_enginesize_db) & " cc", ""))
          
          Call rep.Out("{ENDSECTION}")
          Call rep.Out("{NEWPAGE}")
          
          Call rep.Out("{BEGINSECTION}")
          Call rep.Out("{PUSHY}")
          Call P46BackgroundBoxCol1(rep, "Details of the car provided:", 82)
          Call P46InputBoxFullColumnLengthCol1(rep, "{Arial=10,b}Make and model", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_Make_db)))
          Call P46InputBoxFullColumnLengthCol1(rep, "", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_Model_db)))
          Call P46InputBoxFullColumnLengthCol1(rep, "Engine size", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_enginesize_db)))
          
          Call P46CaptionCol1(rep, "Please tick one of these boxes to show the engine size")
          
          Call P464TickOut(rep, L_P46_COL_1_X, "up to 1400cc", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46LowCC)), "2001cc or more", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46HighCC)), "1401-2000cc", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46MediumCC)), "unknown", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46NoCC)))
          Call P46InputBoxFullColumnLengthCol1(rep, "Date first registered DD MM YYYY", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_Registrationdate_db)))

          Call P46CaptionCol1(rep, "Emissions", "{Arial=10,b}")
          Call P46CaptionCol1(rep, "Give details of the approved CO2 emissions figure at the\ndate of first registration")
          
          Call P46CaptionCol1(rep, "Grams of CO2 per kilometre" & OutLineBoxR(L_P46_COL_1_X + 24, 15, L_HMIT_STANDARDBOX_HEIGHT, IIf(.value(car_P46WithdrawnWithoutReplacement) Or .value(car_p46CarbonDioxide_db) = 0 Or Year(.value(car_Registrationdate_db)) < 1998 Or .value(car_p46NoApprovedCO2Figure_db) Or (.value(car_p46FuelType_db) = CCFT_ELECTRIC), "", .value(car_p46CarbonDioxide_db))))
          Call P46CaptionCol1(rep, "If you have not filled in a figure for the approved CO2\nemissions, please show the reason:")
          Call P46TickRowCol1(rep, "Car was first registered before 1998, or", IIf(.value(car_P46WithdrawnWithoutReplacement), False, IIf(Year(.value(car_Registrationdate_db)) < 1998, True, False)))
          Call P46TickRowCol1(rep, "1998 or later car, for which there is no approved\nCO2 emissions figure (for example some personal\nimports from outside the European Community)", IIf(.value(car_P46WithdrawnWithoutReplacement) Or Year(.value(car_Registrationdate_db)) < 1998, False, .value(car_p46NoApprovedCO2Figure_db)))
          
          Call P46CaptionCol1(rep, "Type of Fuel or power used", "{Arial=10,b}")
          Call P46CaptionCol1(rep, "{Arial=10,u}Key letter{Arial=10,n} - use the list of key letters below to find\nthe appropriate key letter and enter it in the box below:")
          Call P46CaptionCol1(rep, "Type{x=" & (L_P46_COL_1_X + 39) & "}Key letter", "{Arial=10,u}")

          Call P46FuelTypeLine(rep, "Petrol", P46FuelTypeStrings(CCFT_PETROL).Letter)
          Call P46FuelTypeLine(rep, "Diesel", P46FuelTypeStrings(CCFT_DIESEL).Letter)
          Call P46FuelTypeLine(rep, "Euro IV emissions standard diesel", P46FuelTypeStrings(CCFT_EUROIVDIESEL).Letter)
          
          
          
          Call P46CaptionCol1(rep, "Alternative fuel/power types:", "{Arial=10,u}")
          Call P46FuelTypeLine(rep, "Hybrid electric\nA hybrid car combines a petrol engine\nwith an electric motor", P46FuelTypeStrings(CCFT_HYBRID).Letter)
          
          Call P46FuelTypeLine(rep, "Zero emmission car (one which cannot in any\ncircumstances emit CO2 by being driven)\nincluding electric", P46FuelTypeStrings(CCFT_ELECTRIC).Letter)
          Call P46FuelTypeLine(rep, "Bi-fuel\nFor a gas and petrol car that had an approved\nCO2 emissions figure for gas at first registration", P46FuelTypeStrings(CCFT_BIFUEL_WITH_CO2_FOR_GAS).Letter)
          Call P46FuelTypeLine(rep, "E85\nFor a car manufactured to be able to run on E85,\na mixture of petrol and at least 85% bioethanol", P46FuelTypeStrings(CCFT_E85_BIO_ENTHANOL_AND_PETROL).Letter)
          Call P46FuelTypeLine(rep, "Conversion or older bi-fuel\nFor a gas and petrol car that only had an approved\nCO2 emissions figure for petrol at first registration", P46FuelTypeStrings(CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B).Letter)
          Call P46CaptionCol1(rep, "Key Letter{Arial=10,n}{x=" & (L_P46_COL_1_X + 42) & "}{WBTEXTBOXL=3,1.4, " & IIf(.value(car_P46WithdrawnWithoutReplacement), "", IIf(.value(car_p46FuelType_db) = 1, "D", .value(car_p46FuelTypeString))) & "}", "{Arial=10,b}")
          Call P46CaptionCol1(rep, "If you think that the car uses a type of fuel that is not mentioned\nabove please contact your HM Revenue & Customs office.", "{Arial=8,n}")

          Call rep.Out("{POP}")
          Call P46BackgroundBoxCol2(rep, "Details of the car provided:", 37)
          Call P46CaptionCol2(rep, "Price and employee contributions", "{Arial=10,b}")
          Call P46InputBoxFullColumnLengthCol2(rep, "Price of the car (not the price acually paid, but the price for\ntax purposes - normally the list price at the date of\nfirst registration)", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_ListPrice_db)))))
          Call P46InputBoxFullColumnLengthCol2(rep, "Price of accessories not included in the price of the car", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_Accessories)))))
          Call P46InputBoxFullColumnLengthCol2(rep, "Date the car was first made available to the employee\nDD MM YYYY", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(Car_AvailableFrom_db)))
          Call P46InputBoxFullColumnLengthCol2(rep, "Capital contribution (if any) made by the employee\ntowards the cost of the car and for accessories", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_capitalcontribution_db)))))
          Call P46InputBoxFullColumnLengthCol2(rep, "Sum that the employee is required to pay (if any) for\nthe private use of the car", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_MadeGood_db)))))
          Call P46CaptionCol2(rep, "If so, how often?")
          Call P464TickOut(rep, L_P46_COL_2_X, "Weekly", P46PaymentFrequencyEx(p46car, P46PF_WEEKLY), "Quarterly", P46PaymentFrequencyEx(p46car, P46PF_QUARTERLY), "Monthly", P46PaymentFrequencyEx(p46car, P46PF_MONTHLY), "Yearly", P46PaymentFrequencyEx(p46car, P46PF_ANNUALLY) Or P46PaymentFrequencyEx(p46car, P46PF_ACTUAL))
          
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")
          
          Call P46BackgroundBoxCol2(rep, "Fuel for private use", 19)
          Call P46CaptionCol2(rep, "Is fuel provided for private use?")
          Call P46CaptionCol2(rep, "Tick 'Yes' if the employee is provided with any fuel at all for\nprivate use, including any combination of petrol and gas,\nor petrol for a hybrid electric car.")
          Call P46CaptionCol2(rep, "Do not tick 'Yes' if only electricity is provided.")
          Call P462TickOut(rep, L_P46_COL_2_X, "Yes", IIf(.value(car_P46WithdrawnWithoutReplacement) Or .value(car_p46FuelType_db) = CCFT_ELECTRIC, False, .value(car_privatefuel_db)), "No", IIf(.value(car_P46WithdrawnWithoutReplacement), False, (Not .value(car_privatefuel_db) Or (.value(car_privatefuel_db) And .value(car_p46FuelType_db) = CCFT_ELECTRIC))))
          Call P46CaptionCol2(rep, "If 'Yes', must the employee pay for all fuel used for private\nmotoring and do you expect them to continue to do so?")
          Call P462TickOut(rep, L_P46_COL_2_X, "Yes", IIf(.value(car_privatefuel_db), (IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_requiredmakegood_db))), False), "No", IIf(.value(car_privatefuel_db), (IIf(.value(car_P46WithdrawnWithoutReplacement), False, (Not .value(car_requiredmakegood_db)))), False))
          
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")

          Call P46BackgroundBoxCol2(rep, "Declaration", 22)
          Call P46CaptionCol2(rep, "I declare that the information I have given is correct\naccording to the best of my knowledge and belief.")
          Call P46InputBoxFullColumnLengthCol2(rep, "Signature", "", 4)
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")

          Call P46InputBoxFullColumnLengthCol2(rep, "Capacity in which signed", "")
          Call P46InputBoxFullColumnLengthCol2(rep, "Date DD MM YYYY", "")
          
        Call rep.Out("{ENDSECTION}")
        Call rep.Out("{NEWPAGE}")
        
    End With
  Next
  
  Report_P46CarBeforeApril2018 = True
Report_P46Car_end:
  Call xReturn("Report_P46Car")
  Exit Function
  
Report_P46Car_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "Report_P46Car", "P46 Car Report", "Error printing P46 Car...")
  Resume Report_P46Car_end
  Resume
      
End Function
Public Function Report_P46Car(rep As Reporter, ee As Employee, dDateFrom As Date, dDateTo As Date) As Boolean
  Report_P46Car = Report_P46CarAfterApril2018(rep, ee, dDateFrom, dDateTo)
End Function

Public Function Report_P46CarAfterApril2018(rep As Reporter, ee As Employee, dDateFrom As Date, dDateTo As Date) As Boolean
  Dim P46Cars As ObjectList
  Dim p46car As IBenefitClass
  Dim i As Long, ben As IBenefitClass
  Dim benEmployer As IBenefitClass
  Dim bCarDetails  As Boolean, bFuelNoMadeGood As Boolean
  On Error GoTo Report_P46Car_err
  Dim v As Variant
  
  Call xSet("Report_P46Car")
  
  Set benEmployer = p11d32.CurrentEmployer
   
  If ee Is Nothing Then Call Err.Raise(ERR_NO_EMPLOYEE, "Report_P46Car", "The employee is nothing can not print the P46 car return.")
  If Not ee.GetP46Cars(P46Cars, dDateFrom, dDateTo) Then GoTo Report_P46Car_end

  'IK 17/06/2003 getting the total number of cars. Search lNumCars to see usage
  Dim CompanyCars As ObjectList
  Set CompanyCars = BenefitsOfType(ee, BC_COMPANY_CARS_F)
  Dim lNumCars As Long
  lNumCars = CompanyCars.count
  Set CompanyCars = Nothing
  
  rep.PageFooter = HMITFooter("P46(Car)(" & p11d32.AppYear & ")", ee, False)

  Set ben = ee
  
  For i = 1 To P46Cars.count
    Set p46car = P46Cars(i)
    With p46car
              
         Call rep.Out("{BEGINSECTION}")
        
         Call rep.Out(vbCrLf & "{x=51}{Arial=14,b}Car provided for the private use" & vbCrLf & _
                    "{Arial=13,b}{x=3}HM Revenue" & "{Arial=14,b}{x=53}of an employee or a director" & vbCrLf & _
                    "{Arial=16,b}{x=3}&Customs" & vbCrLf)
   

        If p11d32.ReportPrint.DraftReportsp46 Then Call rep.Out("{Arial=16,b}{x=30}DRAFT" & vbCrLf)
          
          
        Call RepOutCrLf(rep, "\n{x=2}{Arial=14,b}Use from 6 April 2018 onwards\n")

        
  
         
          Call RepOutCrLf(rep, "{Arial=10,n}\nYou must fill in this form if there's a change that affects car benefits that hasn't been payrolled for an employee or a director\nfor " & _
                           "whom a car is made available for private use. Complete and return this form within 28 days of the end of the quarter\nto " & _
                           Format$(p11d32.Rates.value(P46Quarter1End), "d mmmm") & ", " & Format$(p11d32.Rates.value(P46Quarter2End), "d mmmm") & ", " & Format$(p11d32.Rates.value(P46Quarter3End), "d mmmm") & _
                           " or " & Format$(p11d32.Rates.value(P46Quarter4End), "d mmmm") & " in which the change takes place.\n\nYou musn't fill in this form if you're taxing the employee's car benefit through the payrolling service. If you are, you should\ninclude the benefit on the employee's Full Payment Submission (FPS).\n\n" & _
                           "For guidance, go to www.gov.uk/guidance/payrolling-tax-employees-benefits-and-expenses-through-your-payroll\n\n", 2)
          
        Call rep.Out("{PUSHY}")
          Call P46BackgroundBoxCol1(rep, "Employer's details", 16)
          
          Call P46InputBoxFullColumnLengthCol1(rep, "Name", benEmployer.Name)
          Call P46InputBoxFullColumnLengthCol1(rep, "Phone number", benEmployer.value(employer_contactnumber_db))
          Call P46InputBoxFullColumnLengthCol1(rep, "PAYE reference", benEmployer.value(employer_Payeref_db))
          
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")
          
          Call P46BackgroundBoxCol1(rep, "Employee's or director's details", 31.57)
          Call P46InputBoxFullColumnLengthCol1(rep, "Name", ee.FullName)
          Call P46InputBoxFullColumnLengthCol1(rep, "National Insurance number", ben.value(ee_NINumber_db))
          Call P46InputBoxFullColumnLengthCol1(rep, "Date of birth (if known) DD MM YYYY", DateValReadToScreenOnlyValidDates(ben.value(ee_DOB_db)))
          Call P46CaptionCol1(rep, "Gender")
          Call RepOutCrLf(rep, "{x=" & (L_P46_COL_1_X + 1) & "}{Arial=10,n}Male ")
          Call rep.Out(TickOut(ben.value(ee_Gender_db) = S_GENDER_MALE))
          Call RepOutCrLf(rep, "{x=" & (L_P46_COL_1_X + 15) & "}{Arial=10,n}Female ")
          Call rep.Out(TickOut(ben.value(ee_Gender_db) = S_GENDER_FEMALE))
          
          Call rep.Out("{POP}")
          
          Call P46BackgroundBoxCol2(rep, "General details", 49.57)
          Call P46CaptionCol2(rep, "Show here and on Page 2 any changes that have\nbeen made", "{Arial=10,b}\n")
          Call P46TickRowCol2(rep, "We provided the employee or director with a car\nwhich is available for private use.", p46car.value(car_P46FirstProvidedWithCar))
          
          'CAD 2010, will be back in for 2011 PAYEonline!
          If (p11d32.ReportPrint.P46PrintReplacedP46s) And p46car.value(car_P46CarProvidedReplaced) Then
           Call P46TickRowCol2(rep, "We replaced a car provided to the employee or\ndirector with another car which is available for\nprivate use.", p46car.value(car_P46CarProvidedReplaced))
           Call P46CaptionCol2(rep, "If the employee has more than one car available for private\nuse, please give details of the car that you replaced.", "{Arial=10,i}")
           Call P46InputBoxFullColumnLengthCol2(rep, "Make and model", IIf(.value(car_p46ReplacementForMultipleCarsMakeAndModel), .value(car_CarReplacedMake_db), ""))
           Call P46InputBoxFullColumnLengthCol2(rep, "", IIf(.value(car_p46ReplacementForMultipleCarsMakeAndModel), .value(car_CarReplacedModel_db), ""))
           Call P46InputBoxFullColumnLengthCol2(rep, "Engine size", IIf(.value(car_p46ReplacementForMultipleCarsMakeAndModel) And .value(car_CarReplacedEngineSize_db) <> 0, .value(car_CarReplacedEngineSize_db) & " cc", ""))
          End If
          
          
          
          Call P46TickRowCol2(rep, "We provided the employee or director with a second\nor further car, which is available for private use.", p46car.value(car_P46SecondCar))
          Call P46TickRowCol2(rep, "We have withdrawn a car provided to the employee\nor director and have not replaced it.", p46car.value(car_P46WithdrawnWithoutReplacement))
          
          Call P46CaptionCol2(rep, "If you ticked this box, please fill in the boxes below,\nand then go straight to the declaration overleaf. Do not\ncomplete the other sections.", "{Arial=10,i}")
          Call P46InputBoxFullColumnLengthCol2(rep, "Date withdrawn DD MM YYYY", IIf(p46car.value(car_P46WithdrawnWithoutReplacement), .value(Car_AvailableTo_db), ""))
          Call P46CaptionCol2(rep, "Please give details of the car withdrawn.")
          Call P46InputBoxFullColumnLengthCol2(rep, "Make and model", IIf(.value(car_P46WithdrawnWithoutReplacement), .value(car_Make_db), ""))
          Call P46InputBoxFullColumnLengthCol2(rep, "", IIf(p46car.value(car_P46WithdrawnWithoutReplacement), .value(car_Model_db), ""))
          
          Call P46InputBoxFullColumnLengthCol2(rep, "Engine size", IIf(.value(car_P46WithdrawnWithoutReplacement) And .value(car_enginesize_db) <> 0, .value(car_enginesize_db) & " cc", ""))
          
          Call rep.Out("{ENDSECTION}")
          Call rep.Out("{NEWPAGE}")
          
          Call rep.Out("{BEGINSECTION}")
          Call rep.Out("{PUSHY}")
          Call P46BackgroundBoxCol1(rep, "Details of the car provided:", 82)
          Call P46InputBoxFullColumnLengthCol1(rep, "{Arial=10,b}Make and model", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_Make_db)))
          Call P46InputBoxFullColumnLengthCol1(rep, "", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_Model_db)))
          Call P46InputBoxFullColumnLengthCol1(rep, "Engine size", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_enginesize_db)))
          
          Call P46CaptionCol1(rep, "Please tick one of these boxes to show the engine size")
          
          Call P464TickOut(rep, L_P46_COL_1_X, "up to 1400cc", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46LowCC)), "2001cc or more", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46HighCC)), "1401-2000cc", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46MediumCC)), "unknown", IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_P46NoCC)))
          Call P46InputBoxFullColumnLengthCol1(rep, "Date first registered DD MM YYYY", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(car_Registrationdate_db)))

          Call P46CaptionCol1(rep, "Emissions", "{Arial=10,b}")
          Call P46CaptionCol1(rep, "Give details of the approved CO2 emissions figure at the\ndate of first registration")
          
          Call P46CaptionCol1(rep, "Grams of CO2 per kilometre" & OutLineBoxR(L_P46_COL_1_X + 24, 15, L_HMIT_STANDARDBOX_HEIGHT, IIf(.value(car_P46WithdrawnWithoutReplacement) Or .value(car_p46CarbonDioxide_db) = 0 Or Year(.value(car_Registrationdate_db)) < 1998 Or .value(car_p46NoApprovedCO2Figure_db) Or (.value(car_p46FuelType_db) = CCFT_ELECTRIC), "", .value(car_p46CarbonDioxide_db))))
          Call P46CaptionCol1(rep, "If you have not filled in a figure for approved CO2\nemissions, please show the reason:")
          Call P46TickRowCol1(rep, "car was first registered before 1998, or", IIf(.value(car_P46WithdrawnWithoutReplacement), False, IIf(Year(.value(car_Registrationdate_db)) < 1998, True, False)))
          Call P46TickRowCol1(rep, "1998 or later car, for which there is no approved\nCO2 emissions figure (for example some personal\nimports from outside the European Community)", IIf(.value(car_P46WithdrawnWithoutReplacement) Or Year(.value(car_Registrationdate_db)) < 1998, False, .value(car_p46NoApprovedCO2Figure_db)))
          
          Call P46CaptionCol1(rep, "Type of Fuel or power used", "{Arial=10,b}")
          Call P46CaptionCol1(rep, "{Arial=10,u}Key letter{Arial=10,n} - use the list of key letters below to find\nthe appropriate key letter and enter it in the box below:")
          Call P46CaptionCol1(rep, "Type{x=" & (L_P46_COL_1_X + 39) & "}Key letter", "{Arial=10,u}")

          Call P46FuelTypeLine(rep, "Diesel (all other diesel cars)", P46FuelTypeStrings(CCFT_DIESEL).Letter)
          Call P46FuelTypeLine(rep, "Diesel cars Euro Standard 6d", P46FuelTypeStrings(CCFT_RDE2_DIESEL).Letter)
          Call P46FuelTypeLine(rep, "All other cars", P46FuelTypeStrings(CCFT_PETROL).Letter)
          
          Call P46CaptionCol1(rep, "Key Letter{Arial=10,n}{x=" & (L_P46_COL_1_X + 42) & "}{WBTEXTBOXL=3,1.4, " & IIf(.value(car_P46WithdrawnWithoutReplacement), "", IIf(.value(car_p46FuelType_db) = 1, "D", .value(car_p46FuelTypeString))) & "}", "{Arial=10,b}")

          Call P46CaptionCol1(rep, "Price and employee contributions", "{Arial=10,b}")
          Call P46InputBoxFullColumnLengthCol1(rep, "Price of the car (not the price acually paid, but the price for\ntax purposes - normally the list price at the date of\nfirst registration)", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_ListPrice_db)))))
          Call P46InputBoxFullColumnLengthCol1(rep, "Price of accessories not included in the price of the car", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_Accessories)))))
          Call P46InputBoxFullColumnLengthCol1(rep, "Date the car was first made available to the employee\nDD MM YYYY", IIf(.value(car_P46WithdrawnWithoutReplacement), "", .value(Car_AvailableFrom_db)))
          
          
          If (.value(car_OPRA_Ammount_Foregone_Used_For_Value)) Then
            v = .value(car_OPRA_Ammount_Foregone_db)
          Else
            v = 0
          End If
          
          Call P46InputBoxFullColumnLengthCol1(rep, "The cash foregone in respect of the car", FormatWN(v))
          
          
          Call rep.Out("{POP}")
          Call P46BackgroundBoxCol2(rep, "Details of the car provided:", 20)
          Call P46InputBoxFullColumnLengthCol2(rep, "Capital contribution (if any) made by the employee\ntowards the cost of the car and for accessories", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_capitalcontribution_db)))))
          Call P46InputBoxFullColumnLengthCol2(rep, "Sum that the employee is required to pay (if any) for\nthe private use of the car", IIf(.value(car_P46WithdrawnWithoutReplacement), "", (FormatWN(.value(car_MadeGood_db)))))
          Call P46CaptionCol2(rep, "If so, how often?")
          Call P464TickOut(rep, L_P46_COL_2_X, "Weekly", P46PaymentFrequencyEx(p46car, P46PF_WEEKLY), "Quarterly", P46PaymentFrequencyEx(p46car, P46PF_QUARTERLY), "Monthly", P46PaymentFrequencyEx(p46car, P46PF_MONTHLY), "Yearly", P46PaymentFrequencyEx(p46car, P46PF_ANNUALLY) Or P46PaymentFrequencyEx(p46car, P46PF_ACTUAL))
          
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")
          
          Call P46BackgroundBoxCol2(rep, "Fuel for private use", 19)
          Call P46CaptionCol2(rep, "Is fuel provided for private use?")
          Call P46CaptionCol2(rep, "Tick 'Yes' if the employee is provided with any fuel at all for\nprivate use, including any combination of petrol and gas,\nor petrol for a hybrid electric car.")
          Call P46CaptionCol2(rep, "Don't 'Yes' if only electricity is provided.")
          Call P462TickOut(rep, L_P46_COL_2_X, "Yes", IIf(.value(car_P46WithdrawnWithoutReplacement) Or .value(car_p46FuelType_db) = CCFT_ELECTRIC, False, .value(car_privatefuel_db)), "No", IIf(.value(car_P46WithdrawnWithoutReplacement), False, (Not .value(car_privatefuel_db) Or (.value(car_privatefuel_db) And .value(car_p46FuelType_db) = CCFT_ELECTRIC))))
          Call P46CaptionCol2(rep, "If 'Yes', must the employee pay for all fuel used for private\nmotoring and do you expect them to continue to do so?")
          Call P462TickOut(rep, L_P46_COL_2_X, "Yes", IIf(.value(car_privatefuel_db), (IIf(.value(car_P46WithdrawnWithoutReplacement), False, .value(car_requiredmakegood_db))), False), "No", IIf(.value(car_privatefuel_db), (IIf(.value(car_P46WithdrawnWithoutReplacement), False, (Not .value(car_requiredmakegood_db)))), False))
          
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")

          Call P46BackgroundBoxCol2(rep, "Declaration", 22)
          Call P46CaptionCol2(rep, "I declare that the information I have given is correct\naccording to the best of my knowledge and belief.")
          Call P46InputBoxFullColumnLengthCol2(rep, "Signature", "", 4)
          Call RepOutCrLf(rep, "\n")
          Call RepOutCrLf(rep, "\n")

          Call P46InputBoxFullColumnLengthCol2(rep, "Capacity in which signed", "")
          Call P46InputBoxFullColumnLengthCol2(rep, "Date DD MM YYYY", "")
          
        Call rep.Out("{ENDSECTION}")
        Call rep.Out("{NEWPAGE}")
        
    End With
  Next
  
  Report_P46CarAfterApril2018 = True
Report_P46Car_end:
  Call xReturn("Report_P46Car")
  Exit Function
  
Report_P46Car_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "Report_P46Car", "P46 Car Report", "Error printing P46 Car...")
  Resume Report_P46Car_end
  Resume
      
End Function
Private Sub HMITSectionL(rep As Reporter, ee As Employee, BenArr() As BEN_CLASS)
  'same as HMITAssetsTransferredType but for value = 0 with Computer Related = true need to not include if effect of £500 deminimuns makes 0
  'came in 1999/2000
  Dim Benefit As Variant, MadeGood As Variant, value As Variant, Description As String
  Dim sSectionTitle As String
  Dim sIRDesc As String
  
  sSectionTitle = "Description of asset"
  
  Call HMITSectionHeader(rep, HMIT_L, "Assets placed at the employee's disposal")
  Call HMITColHeaders(rep, "Cost of the benefit", "or amount foregone", "Amount made good or", "from which tax deducted", "Cash equivalent", "or relevant amount")
  
  If ee.SumBenefit(Description, value, MadeGood, Benefit, BenArr()) Then
    If IsNumeric(value) Then
      If value > 0 Then
        value = FormatWNRPT(value)
        MadeGood = FormatWNRPT(MadeGood)
        Benefit = FormatWNRPT(Benefit)
        Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Description of asset", True)
      Else
        Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Description of asset", True)
      End If
    Else
      value = FormatWNRPT(value)
      MadeGood = FormatWNRPT(MadeGood)
      Benefit = FormatWNRPT(Benefit)
      Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Description of asset", True)
    End If
  Else
    Call HMITAssetsTransferredTypeNIC(rep, ee, BenArr, "{x=8}Description of asset", True)
  End If
  
End Sub
Private Sub HMITAssetsTransferredType(rep As Reporter, ee As Employee, BenArr() As BEN_CLASS, sSectionTitle As String, bAssetDescription As Boolean, Optional VALUE_ENUM As Long = ITEM_VALUE, Optional MADEGOOD_ENUM As Long = ITEM_MADEGOOD, Optional BENEFIT_ENUM As Long = ITEM_BENEFIT)
  Dim Benefit As Variant, MadeGood As Variant, value As Variant, Description As String
  Dim sIRDesc As String, tempDesc As String
  
  Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr, , , , sIRDesc)
  If Len(sIRDesc) > 0 And Not LCase(sIRDesc) = "other" Then
    tempDesc = sIRDesc
  Else
    tempDesc = Description
  End If
  Call HMITAssetsTransferredTypeOut(rep, p11d32.Rates.BenClassTo(BenArr(1), BCT_HMIT_BOX_NUMBER), sSectionTitle, tempDesc, value, MadeGood, Benefit, bAssetDescription)
  
End Sub

Private Sub HMITAssetsTransferredTypeNIC(rep As Reporter, ee As Employee, BenArr() As BEN_CLASS, sSectionTitle As String, bAssetDescription As Boolean, Optional VALUE_ENUM As Long = ITEM_VALUE, Optional MADEGOOD_ENUM As Long = ITEM_MADEGOOD, Optional BENEFIT_ENUM As Long = ITEM_BENEFIT)
  Dim Benefit As Variant, MadeGood As Variant, value As Variant, Description As String
  Dim sIRDesc As String, tempDesc As String
    
  Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr, , , , sIRDesc)
  If Len(sIRDesc) > 0 And Not LCase(sIRDesc) = "other" Then
    tempDesc = sIRDesc
  Else
    tempDesc = Description
  End If
  Call HMITAssetsTransferredTypeOutNIC(rep, p11d32.Rates.BenClassTo(BenArr(1), BCT_HMIT_BOX_NUMBER), sSectionTitle, tempDesc, value, MadeGood, Benefit, bAssetDescription, "1A")
  
End Sub


Private Function HMITAssetsTransferredTypeOut(rep As Reporter, sBoxNumber As String, sSectionTitle As String, Description As String, value As Variant, MadeGood As Variant, Benefit As Variant, bAssetDescription As Boolean)
  Description = HMITFieldTrim(Description, 25)
  
  rep.Out (IIf(bAssetDescription, OutLineBoxL(26, 22, 1.4, Description), "") & _
          OutLineBoxR(L_HMIT_COL_1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, value) & _
          OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, MadeGood) & _
          FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, sBoxNumber) & _
          OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, Benefit) & _
          HMITMinus(63) & _
          HMITEquals() & _
          LineText(sSectionTitle))
End Function

Private Function HMITAssetsTransferredTypeOutNIC(rep As Reporter, sBoxNumber As String, sSectionTitle As String, Description As String, value As Variant, MadeGood As Variant, Benefit As Variant, bAssetDescription As Boolean, NIC As String)
  Description = HMITFieldTrim(Description, 25)
  
  rep.Out (IIf(bAssetDescription, OutLineBoxL(26, 22, 1.4, Description), "") & _
          OutLineBoxR(L_HMIT_COL_1, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, value) & _
          OutLineBoxR(L_HMIT_COL_2, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, MadeGood) & _
          FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, sBoxNumber) & _
          FillBoxNIC(L_HMIT_COL_5, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
          OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, Benefit) & _
          HMITMinus(63) & _
          HMITEquals() & _
          LineText(sSectionTitle))
End Function

Private Sub HMITColHeaders(rep As Reporter, sCol1Line1 As String, sCol1Line2 As String, sCol2Line1 As String, sCol2Line2 As String, sCol4Line1 As String, sCol4Line2 As String)
  Call rep.Out(HMITColText(sCol1Line1, L_HMIT_COL_1) & HMITColText(sCol2Line1, L_HMIT_COL_2) & HMITColText(sCol4Line1, L_HMIT_COL_4) & vbCrLf & _
               HMITColText(sCol1Line2, L_HMIT_COL_1) & HMITColText(sCol2Line2, L_HMIT_COL_2) & HMITColText(sCol4Line2, L_HMIT_COL_4) & vbCrLf)
End Sub
Private Sub HMITVanType(rep As Reporter, ee As Employee, BenArr() As BEN_CLASS, sSectionTitle As String, Optional VALUE_ENUM As Long = ITEM_VALUE, Optional MADEGOOD_ENUM As Long = ITEM_MADEGOOD, Optional BENEFIT_ENUM As Long = ITEM_BENEFIT)
  Dim Benefit As Variant, MadeGood As Variant, value As Variant, Description As String
  
  Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr, VALUE_ENUM, MADEGOOD_ENUM, BENEFIT_ENUM)
  Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, Benefit) & _
               FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, p11d32.Rates.BenClassTo(BenArr(1), BCT_HMIT_BOX_NUMBER)) & LineText(sSectionTitle))
End Sub

Private Sub HMITVanTypeNIC(rep As Reporter, ee As Employee, BenArr() As BEN_CLASS, sSectionTitle As String, Optional VALUE_ENUM As Long = ITEM_VALUE, Optional MADEGOOD_ENUM As Long = ITEM_MADEGOOD, Optional BENEFIT_ENUM As Long = ITEM_BENEFIT, Optional BoxNumber As String = "")
  Dim Benefit As Variant, MadeGood As Variant, value As Variant, Description As String
  
  If Len(BoxNumber) = 0 Then
    BoxNumber = p11d32.Rates.BenClassTo(BenArr(1), BCT_HMIT_BOX_NUMBER)
  End If
  Call SumBenefitFWNRPT(ee, Description, value, MadeGood, Benefit, BenArr, VALUE_ENUM, MADEGOOD_ENUM, BENEFIT_ENUM)
  Call rep.Out(OutLineBoxR(L_HMIT_COL_4, L_HMIT_STANDARDBOX_WIDTH, L_HMIT_STANDARDBOX_HEIGHT, Benefit) & _
               FillBoxNIC(L_HMIT_COL_5, 2, L_HMIT_STANDARDBOX_HEIGHT, "1A") & _
               FillBox(L_HMIT_COL_3, 4, L_HMIT_STANDARDBOX_HEIGHT, BoxNumber) & _
               LineText(sSectionTitle))
End Sub
Private Sub HMITSectionHeader(rep, HMITS As HMIT_SECTIONS, sBulletTitle As String)
  Call BoxBulletText(rep, p11d32.Rates.HMITSectionToHMITCode(HMITS), sBulletTitle, 3)
End Sub
Private Sub BoxBulletText(rep, sBoxText As String, sBulletTitle As String, lXStart As Long)
  
  'To have consistency between screen and print-out
  If Len(HMITText(sBulletTitle, lXStart)) > 100 Then
    Call rep.Out(FillBox(lXStart, 3, L_HMIT_STANDARDBOX_HEIGHT, sBoxText) & _
                 HMITText(sBulletTitle, lXStart) & vbCrLf)
  Else
    Call rep.Out(FillBox(lXStart, 3, L_HMIT_STANDARDBOX_HEIGHT, sBoxText) & _
                 HMITText(sBulletTitle, lXStart))
  End If
End Sub

Private Function TickOut(v As Variant) As String

  If v = True Then
    TickOut = "{Wingdings=12,nb}{WBTEXTBOXL=3,1.4,ü}"
  ElseIf v = -2 Then
    TickOut = "{Arial=12,nb}{WBTEXTBOXL=3,1.4, x}"
  Else
    TickOut = "{Wingdings=12,nb}{WBTEXTBOXL=3,1.4, }"
  End If
End Function
Public Function EmployeeLettersAddress(benEE As IBenefitClass, Optional bIsEmail As Boolean = False) As String
  Dim i As EmployeeItems
  
  On Error GoTo EmployeeLettersAddress_ERR
  Call xSet("EmployeeLettersAddress")
  
  If benEE Is Nothing Then Call Err.Raise(ERR_EMPLOYEE_IS_NOTHING, "EmployeeLettersAddress", "Employee is nothing.")
  
  For i = [_EE_ADDRESS_DETAILS_FIRST_ITEM] To [_EE_ADDRESS_DETAILS_LAST_ITEM]
    If Len(benEE.value(i)) > 0 Then
      EmployeeLettersAddress = EmployeeLettersAddress & benEE.value(i)
      If i < [_EE_ADDRESS_DETAILS_LAST_ITEM] Then
        EmployeeLettersAddress = EmployeeLettersAddress & vbCrLf
      End If
    End If
  Next i
  
  If Len(EmployeeLettersAddress) > 0 Then
    If Not bIsEmail Then EmployeeLettersAddress = "{PUSHX}{STATICX}" & EmployeeLettersAddress & "{POP}"
    EmployeeLettersAddress = EmployeeLettersAddress & vbCrLf
  End If
  
EmployeeLettersAddress_END:
  Call xReturn("EmployeeLettersAddress")
  Exit Function
EmployeeLettersAddress_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeLettersAddress", "Employee Letters Address", "Error getting an employees address for the employee letter.")
  Resume EmployeeLettersAddress_END:
End Function

Public Function TimeStampReport() As String
  TimeStampReport = IIf(p11d32.ReportPrint.TimeStamp, " {DATE} {TIME} ", "")
End Function

Public Function EmployeeLetterMenuCaptions(ByVal elmc As EL_MENU_CAPTIONS) As String

  Select Case elmc
    Case ELMC_FORMAT
      EmployeeLetterMenuCaptions = "&Format"
    Case ELMC_EMPLOYER
      EmployeeLetterMenuCaptions = "&Employer"
    Case ELMC_EMPLOYEE
      EmployeeLetterMenuCaptions = "E&mployee"
    Case ELMC_DATES
      EmployeeLetterMenuCaptions = "&Dates"
    Case ELMC_SUB_REPORTS
      EmployeeLetterMenuCaptions = "Sub Reports"
    Case Else
      ECASE ("Invalid EmployeeLetterMenuCaption")
  End Select
  
End Function


Public Function EmployeeLetterCode(ByVal EMLC As EMPLOYEE_LETTER_CODE, ByVal EMLCT As EMPLOYEE_LETTER_CODE_TYPE, ByVal bIsEmail As Boolean, Optional ee As Employee = Nothing, Optional ByVal rep As Reporter) As String
  Dim ben As IBenefitClass, benEmployer As IBenefitClass
  Dim bEmployeeLetterCaption As Boolean
  On Error GoTo EmployeeLetterCode_ERR

  Call xSet("EmployeeLetterCode")
  
  If EMLCT = ELCT_REPORT_CODE Then 'report codes rely on a current employee

    If ee Is Nothing Then Call Err.Raise(ERR_NO_EMPLOYER, "EmployeeLetterCode", "The employee letter code type requested requires an employee.")
    Set ben = ee
    Set benEmployer = ben.Parent
  Else
    bEmployeeLetterCaption = (EMLCT = ELCT_CAPTION)
    
    If bEmployeeLetterCaption Then EMLCT = ELCT_MENU_CAPTION
  End If
  
  Select Case EMLC
    Case EMPLOYEE_LETTER_CODE.ELC_BOLD
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{BOLD}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Bold text"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = EmployeeLetterFont(ELC_BODY_BOLD, bIsEmail)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_FORMAT)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_COMPANY_NAME
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{COMPANYNAME}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Company name"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = benEmployer.Name
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYER)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_CONTACT_NAME
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{CONTACTNAME}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Contact name"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = benEmployer.value(employer_Contact_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYER)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_ADDRESS
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{ADDRESS}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee address"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = EmployeeLettersAddress(ben, bIsEmail)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
          
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_CONTACT_NUMBER
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{CONTACTNUM}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Contact number"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = benEmployer.value(employer_contactnumber_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYER)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_DATE_NOW
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{DATE}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Date"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = Format$(Now, "Long Date")
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_DATE_TAXYEAR
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{TAXYEAR}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Tax year"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = p11d32.Rates.value(TaxFormYear)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_DATE_NEXT_TAXYEAR
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{NEXTTAXYEAR}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Next tax year"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = p11d32.Rates.value(TaxFormYearNext)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_DATE_KEEP_DETAILS_UNTILL
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{KEEP_DETAILS_UNTIL}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Keep details until"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = Format$(p11d32.Rates.value(EmpLetP11DKeepDetailsUntill), "Long Date")
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_DATE_RESPONSE
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{RESPONSE_DATE}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Response date"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = Format$(benEmployer.value(employer_EmployeeResponseDate_db), "Long Date")
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_DATE_TAX_YEAR_START
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{TAX_YEAR_START}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Tax year start"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = Format$(p11d32.Rates.value(TaxYearStart), "Long Date")
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_DATE_TAX_YEAR_END
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{TAX_YEAR_END}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Tax year end"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = Format$(p11d32.Rates.value(TaxYearEnd), "Long Date")
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_DATES)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_EMPLOYEE_NAME_TITLE
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{TITLE}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee title"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Title_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
      
    Case EMPLOYEE_LETTER_CODE.ELC_EMPLOYEE_NAME_INITIALS
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{INITIALS}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee initials "
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Initials_db)
         Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_EMPLOYEE_NAME_SALUTATION
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{SALUTATION}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee salutation"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Salutation_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_EMPLOYEE_NAME_SURNAME
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{SURNAME}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee surname"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Surname_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_EMPLOYEE_NAME_FIRST
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{FIRSTNAME}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee first name "
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Firstname_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_EMPLOYEE_NAME_FULL
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{NAME}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Employee full name"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_FullName)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_NEWPAGE
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{NEWPAGE}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "New page / page break"
          
        Case ELCT_REPORT_CODE
          If Not bIsEmail Then EmployeeLetterCode = EmployeeLetterCode(EMLC, ELCT_LETTER_FILE_CODES, bIsEmail)
         Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_FORMAT)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
      
    Case ELC_GROUP1
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{GROUP1}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Group 1"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Group1_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_GROUP2
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{GROUP2}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Group 2"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Group2_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
      
    Case ELC_GROUP3
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{GROUP3}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Group 3"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_Group3_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_NI_NUMBER
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{NINUMBER}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "NI number"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_NINumber_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_NORMAL
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{NORMAL}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Normal text"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = EmployeeLetterFont(ELC_BODY_NORMAL, bIsEmail)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_FORMAT)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_PAYE_REF
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{PAYEREF}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "PAYE ref"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = benEmployer.value(employer_Payeref_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYER)
        Case Else
        
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_PERSONNEL_NUMBER
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{PNUMBER}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Personnel number"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = ben.value(ee_PersonnelNumber_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYEE)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_SIGNATORY
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{SIGNATORY}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Signatory"
        Case ELCT_REPORT_CODE
          
          EmployeeLetterCode = benEmployer.value(employer_Signatory_db)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_EMPLOYER)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_TABLE
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{TABLE}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Table of benefits"
        Case ELCT_REPORT_CODE
          EmployeeLetterCode = EmployeeLetterCodeTableReportCode(ee, bIsEmail)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_SUB_REPORTS)
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case EMPLOYEE_LETTER_CODE.ELC_HMIT
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{P11D}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "P11D Return"
        Case ELCT_REPORT_CODE
          Call Report_HMIT(rep, ee)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_SUB_REPORTS)
        Case ELCT_FILE_EXPORT
          EmployeeLetterCode = "P11D"
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_WORKING_PAPERS
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{WORKING_PAPERS}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "Working Papers"
          
        Case ELCT_REPORT_CODE
          Call Report_WK(rep, ee)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_SUB_REPORTS)
        Case ELCT_FILE_EXPORT
          EmployeeLetterCode = "Workings"
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case ELC_P46CAR
      Select Case EMLCT
        Case ELCT_LETTER_FILE_CODES
          EmployeeLetterCode = "{P46CAR}"
        Case ELCT_MENU_CAPTION
          EmployeeLetterCode = "P46 Car Report"
        Case ELCT_REPORT_CODE
          Call Report_P46Car(rep, ee, p11d32.ReportPrint.P46DateFrom, p11d32.ReportPrint.P46DateTo)
        Case ELCT_MENU_PARENT
          EmployeeLetterCode = EmployeeLetterMenuCaptions(ELMC_SUB_REPORTS)
        Case ELCT_FILE_EXPORT
          EmployeeLetterCode = "P46CAR"
        Case Else
          Call ECASE("Unknown employee letter code type, " & EMLCT)
      End Select
    Case Else
      
      Call ECASE("Unknown Employee letter code, " & EMLC)
  End Select
  
  If EMLCT = ELCT_MENU_CAPTION And Not bEmployeeLetterCaption Then EmployeeLetterCode = "&" & EmployeeLetterCode & " - " & EmployeeLetterCode(EMLC, ELCT_LETTER_FILE_CODES, bIsEmail)
  
EmployeeLetterCode_END:
  Call xReturn("EmployeeLetterCode")
  Exit Function
EmployeeLetterCode_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeLetterCode", "Employee Letter Code", "Error getting the employee letter code, code = " & EMLC & ", type = " & EMLCT & ".")
  Resume EmployeeLetterCode_END
  Resume
End Function
Private Function EmployeeLetterTableLine(ByVal bc As BEN_CLASS, ByVal Benefit As Variant, bIsEmail As Boolean) As String
  Dim qs As QString
  
  Set qs = New QString
  
  If Not bIsEmail Then
     Call qs.Append(EmployeeLetterColData(ELT_COL1, ELT_SPACING_REP))
     Call qs.Append(EmployeeLetterFont(ELC_TABLE_NORMAL, bIsEmail))
     Call qs.Append(p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION))
     
     Call qs.Append(EmployeeLetterColData(ELT_COL2, ELT_SPACING_REP))
     Call qs.Append(p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING))
     
     Call qs.Append(EmployeeLetterColData(ELT_COL3, ELT_SPACING_REP))
     Call qs.Append(p11d32.Rates.BenClassTo(bc, BCT_HMIT_BOX_NUMBER))
     
     Call qs.Append(EmployeeLetterColData(ELT_COL4, ELT_SPACING_REP))
     Call qs.Append(EmployeeLetterFont(ELC_TABLE_NORMAL_RIGHT, bIsEmail))
     
     Call qs.Append(FormatWN(Benefit, ""))
     Call qs.Append(vbCrLf)
  Else
    If p11d32.ReportPrint.ParseForNetscape Then
      qs.Append ("<TR>")
      
      qs.Append ("<TD>")
      qs.Append (p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION))
      qs.Append ("</TD>")
      
      qs.Append ("<TD>")
      qs.Append (p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING))
      qs.Append ("</TD>")
      
      qs.Append ("<TD>")
      qs.Append (p11d32.Rates.BenClassTo(bc, BCT_HMIT_BOX_NUMBER))
      qs.Append ("</TD>")
      
      qs.Append ("<TD>")
      qs.Append (FormatWN(Benefit, ""))
      qs.Append ("</TD>")
      
      qs.Append ("</TR>")
    Else
      qs.Append (EmployeeLetterColData(ELT_COL2, ELT_SPACING_EMAIL_TABLE_LINE))
      qs.Append (p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING))
      qs.Append (EmployeeLetterColData(ELT_COL3, ELT_SPACING_EMAIL_TABLE_LINE))
      qs.Append (p11d32.Rates.BenClassTo(bc, BCT_HMIT_BOX_NUMBER))
      qs.Append (EmployeeLetterColData(ELT_COL4, ELT_SPACING_EMAIL_TABLE_LINE))
      qs.Append (EmployeeLetterFont(ELC_TABLE_NORMAL_RIGHT, bIsEmail))
      qs.Append (FormatWN(Benefit, ""))
      qs.Append (EmployeeLetterColData(ELT_COL1, ELT_SPACING_EMAIL_TABLE_LINE))
      qs.Append (p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION))
      qs.Append (vbCrLf)
    End If
  End If
  
  EmployeeLetterTableLine = qs
End Function
Private Function StandardELColSpacing(ByVal ELTC As EL_TABLE_COL, sColData As String) As String
  StandardELColSpacing = xStrPad(sColData, " ", Len(EmployeeLetterColData(ELTC, ELT_COLUMN_HEADER))) & vbTab
End Function
Private Function EmployeeLetterColData(ByVal ELTC As EL_TABLE_COL, ByVal ELTD As EL_TABLE_DATA) As String
    
  Select Case ELTC
    Case ELT_COL1
      Select Case ELTD
        Case ELT_SPACING_REP
          EmployeeLetterColData = "{x=0}"
        Case ELT_SPACING_EMAIL_TABLE_LINE
          EmployeeLetterColData = vbTab & vbTab & vbTab
        Case ELT_SPACING_EMAIL_HEADER_LINE
          EmployeeLetterColData = vbTab
        Case ELT_COLUMN_HEADER
          EmployeeLetterColData = "Category"
        Case Else
          ECASE ("Invalid Employee letter col data.")
      End Select
    Case ELT_COL2
      Select Case ELTD
        Case ELT_SPACING_REP
          EmployeeLetterColData = "{x=40}"
        Case ELT_SPACING_EMAIL_TABLE_LINE
          EmployeeLetterColData = ""
        Case ELT_SPACING_EMAIL_HEADER_LINE
          EmployeeLetterColData = ""
        Case ELT_COLUMN_HEADER
          EmployeeLetterColData = "HMRC section"
        Case Else
          ECASE ("Invalid Employee letter col data.")
      End Select
    Case ELT_COL3
      Select Case ELTD
        Case ELT_SPACING_REP
          EmployeeLetterColData = "{x=55}"
        Case ELT_COLUMN_HEADER
          EmployeeLetterColData = "P11D Box number"
        Case ELT_SPACING_EMAIL_TABLE_LINE
          EmployeeLetterColData = vbTab & vbTab
        Case ELT_SPACING_EMAIL_HEADER_LINE
          EmployeeLetterColData = vbTab
        Case Else
          ECASE ("Invalid Employee letter col data.")
      End Select
    Case ELT_COL4
      Select Case ELTD
        Case ELT_SPACING_REP
          EmployeeLetterColData = "{x=80}"
        Case ELT_SPACING_EMAIL_TABLE_LINE
          EmployeeLetterColData = vbTab & vbTab & vbTab
        Case ELT_SPACING_EMAIL_HEADER_LINE
          EmployeeLetterColData = vbTab
        Case ELT_COLUMN_HEADER
          EmployeeLetterColData = "Cash equivalent £"
        Case Else
          ECASE ("Invalid Employee letter col data.")
      End Select
    Case Else
      Call ECASE("Invalid Col ata requect in EmployeeLetterColData, Coldata = " & ELTC)
    End Select
    
End Function
Private Function EmployeeLetterFont(ByVal ELCF As ELC_FONT, bIsEmail As Boolean)
  If Not bIsEmail Then
    Select Case ELCF
      Case ELC_BODY_NORMAL
        EmployeeLetterFont = "{FONT=" & p11d32.ReportPrint.EmployeeLetterFontName & "," & p11d32.ReportPrint.EmployeeLetterFontSize & ",n}"
      Case ELC_BODY_BOLD
        EmployeeLetterFont = "{FONT=" & p11d32.ReportPrint.EmployeeLetterFontName & "," & p11d32.ReportPrint.EmployeeLetterFontSize & ",nb}"
      Case ELC_TABLE_NORMAL
        EmployeeLetterFont = "{Arial=7,n}"
      Case ELC_TABLE_NORMAL_RIGHT
        EmployeeLetterFont = "{Arial=7,nr}"
      Case ELC_TABLE_HEADING
        EmployeeLetterFont = "{Arial=7,bu}"
      Case ELC_TABLE_HEADING_RIGHT
        EmployeeLetterFont = "{Arial=7,bur}"
      Case Else
        ECASE ("Invalid EmployeeLetter font requested in EmployeeLetterFont, ELC_FONT = " & ELCF & ".")
    End Select
  End If
End Function
Private Function EmployeeLetterCodeTableReportCode(ee As Employee, bIsEmail As Boolean) As String
  Dim Description, MadeGood, Benefit, value, benefit_other
  Dim i As Long, j As Long, k As Long, lRelevantBenClassCount As Long
  Dim qs As QString
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim RelevantBenClasses() As BEN_CLASS
  
  Dim s As String
  On Error GoTo EmployeeLetterCodeTableReportCode_ERR
  
  Call xSet("EmployeeLetterCodeTableReportCode")
    
  If ee Is Nothing Then Call Err.Raise(ERR_EMPLOYEE_IS_NOTHING, "EmployeeLetterCodeTableReportCode", "The employee passed is nothing")
  Set qs = New QString
  
  For i = BC_FIRST_ITEM To BC_UDM_BENEFITS_LAST_ITEM
    If (2 ^ p11d32.Rates.BenClassTo(i, BCT_HMIT_SECTION)) And p11d32.ReportPrint.HMITSections_PRINT Then
      BenArr(1) = i
      If BenefitIsLoan(i) Then
        If ee.AnyLoanBenefit Then
          If Not ee.SumBenefit(Description, value, MadeGood, Benefit, BenArr) Then Benefit = 0
        Else
          Benefit = 0
        End If
      Else
        If Not ee.SumBenefit(Description, value, MadeGood, Benefit, BenArr) Then
          Benefit = 0
          'If i = BC_SHARES_M Then benefit = BoolToString(False)
        Else
          'If i = BC_SHARES_M Then benefit = BoolToString(True)
        End If
      End If
      
      If i = BC_NONSHAREDVANS_G Then
        benefit_other = 0
        If (ee.AnyVanBenefit) Then
          benefit_other = ee.nonSharedVans.value(nsvans_fuel_benefit)
          Benefit = ee.nonSharedVans.value(nsvans_benefit_van_only)
          Call qs.Append(EmployeeLetterTableLine(i, Benefit, bIsEmail))
          Call qs.Append(EmployeeLetterTableLine(BC_NONSHAREDVANS_FUEL_G, benefit_other, bIsEmail))
        End If
      Else
        Call qs.Append(EmployeeLetterTableLine(i, Benefit, bIsEmail))
      End If
      
    End If
   
  Next
  
  If Len(EmployeeLetterCodeTableReportCode) > 0 Then
    Call qs.Append(EmployeeLetterCode(ELC_NORMAL, ELCT_REPORT_CODE, bIsEmail, ee))
  End If
  s = qs
  
  'welcome to bad coding but only got an hour
  If bIsEmail Then
    If p11d32.ReportPrint.ParseForNetscape Then
      EmployeeLetterCodeTableReportCode = "<TABLE border=" & """" & "1" & """" & "<TR><TD>" & EmployeeLetterColData(ELT_COL1, ELT_COLUMN_HEADER) & "</TD>" & _
                                          "<TD>" & EmployeeLetterColData(ELT_COL2, ELT_COLUMN_HEADER) & "</TD>" & _
                                          "<TD>" & EmployeeLetterColData(ELT_COL3, ELT_COLUMN_HEADER) & "</TD>" & _
                                          "<TD>" & EmployeeLetterColData(ELT_COL4, ELT_COLUMN_HEADER) & "</TD></TR>" & _
                                          IIf(Len(s) <> 0, s, "<TR><TD>NO BENEFITS</TD></TR>") & "</TABLE>"
    
    Else
      EmployeeLetterCodeTableReportCode = EmployeeLetterFont(ELC_BODY_BOLD, bIsEmail) & "P11D return of expenses, payments and benefits" & vbCrLf & vbCrLf & _
                                          EmployeeLetterColData(ELT_COL2, ELT_SPACING_EMAIL_HEADER_LINE) & EmployeeLetterColData(ELT_COL2, ELT_COLUMN_HEADER) & _
                                          EmployeeLetterColData(ELT_COL3, ELT_SPACING_EMAIL_HEADER_LINE) & EmployeeLetterColData(ELT_COL3, ELT_COLUMN_HEADER) & _
                                          EmployeeLetterColData(ELT_COL4, ELT_SPACING_EMAIL_HEADER_LINE) & EmployeeLetterColData(ELT_COL4, ELT_COLUMN_HEADER) & _
                                          EmployeeLetterColData(ELT_COL1, ELT_SPACING_EMAIL_HEADER_LINE) & EmployeeLetterFont(ELC_TABLE_HEADING, bIsEmail) & EmployeeLetterColData(ELT_COL1, ELT_COLUMN_HEADER) & _
                                          vbCrLf & vbCrLf & _
                                          IIf(Len(s) <> 0, s, EmployeeLetterFont(ELC_BODY_BOLD, bIsEmail) & "NO BENEFITS")
                                        
    End If
  
  Else
    EmployeeLetterCodeTableReportCode = EmployeeLetterFont(ELC_BODY_BOLD, bIsEmail) & "P11D return of expenses, payments and benefits" & vbCrLf & vbCrLf & _
                                        EmployeeLetterColData(ELT_COL1, ELT_SPACING_REP) & EmployeeLetterFont(ELC_TABLE_HEADING, bIsEmail) & EmployeeLetterColData(ELT_COL1, ELT_COLUMN_HEADER) & _
                                        EmployeeLetterColData(ELT_COL2, ELT_SPACING_REP) & EmployeeLetterColData(ELT_COL2, ELT_COLUMN_HEADER) & _
                                        EmployeeLetterColData(ELT_COL3, ELT_SPACING_REP) & EmployeeLetterColData(ELT_COL3, ELT_COLUMN_HEADER) & _
                                        EmployeeLetterColData(ELT_COL4, ELT_SPACING_REP) & EmployeeLetterFont(ELC_TABLE_HEADING_RIGHT, bIsEmail) & EmployeeLetterColData(ELT_COL4, ELT_COLUMN_HEADER) & vbCrLf & vbCrLf & _
                                        IIf(Len(s) <> 0, s, EmployeeLetterFont(ELC_BODY_BOLD, bIsEmail) & "NO BENEFITS")
  End If
  
EmployeeLetterCodeTableReportCode_END:
  Call xReturn("EmployeeLetterCodeTableReportCode")
  Exit Function
EmployeeLetterCodeTableReportCode_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeLetterCodeTableReportCode", "Employee Letter Code Table Report Code", "Error getting the table for the employee letter.")
  Resume EmployeeLetterCodeTableReportCode_END
End Function

Private Function BoolToString(vBool As Variant)
  If IsNumeric(vBool) Then
    If vBool <> 0 Then
      BoolToString = "Yes"
    Else
      BoolToString = "No"
    End If
  Else
    vBool = "No"
  End If
End Function



Public Function Report_MM_Submission(rep As Reporter, benEmployer As IBenefitClass)

  On Error GoTo Report_MM_SUBMISSION_err
  Call xSet("Report_MM_SUBMISSION")
  
  rep.PageFooter = HMITFooter("Expenses and Benefit Return Submission Document ")

          
  
  Call rep.Out(vbCrLf & "{x=6}{Arial=14,bn}Expenses & Benefits Return " & p11d32.Rates.value(TaxFormYear) & vbCrLf & vbCrLf & _
              "{x=6}{Arial=14,bn}SUBMISSION DOCUMENT" & vbCrLf & vbCrLf & _
              "{x=6}{Arial=14,bn}SUBMITTER" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Submitter's Name:" & "{x=40}{Arial=12,n}" & benEmployer.value(employer_SubmitterName_db) & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Submitter's Reference:" & "{x=40}{Arial=12,n}" & benEmployer.value(employer_SubmitterRef_db) & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}PAYE Reference(s):" & "{x=40}{Arial=12,n}" & benEmployer.value(employer_Payeref_db) & vbCrLf & vbCrLf & vbCrLf)
              
  Call rep.Out("{x=6}{Arial=14,bn}MEDIA RETURN INFORMATION" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Contact Name:" & "{x=40}{Arial=12,n}" & benEmployer.value(employer_Contact_db) & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Position:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Address:" & "{x=40}{Arial=12,n}" & benEmployer.Name & vbCrLf & _
              "{x=40}" & benEmployer.value(employer_AddressLine1_db) & vbCrLf & _
              "{x=40}" & benEmployer.value(employer_AddressLine2_db) & vbCrLf & _
              "{x=40}" & benEmployer.value(employer_AddressLine3_db) & vbCrLf & _
              "{x=40}" & benEmployer.value(employer_AddressLine4_db) & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Postcode:" & "{x=40}{Arial=12,n}" & benEmployer.value(employer_AddressPostCode_db) & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Telephone:" & "{x=40}{Arial=12,n}" & benEmployer.value(employer_contactnumber_db) & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Fax:" & "{x=40}{Arial=12,n}" & vbCrLf & vbCrLf & vbCrLf)

  Call rep.Out("{x=6}{Arial=14,bn}MEDIA SUBMISSION DETAILS" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Total items enclosed:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Test Submission:     Yes / No" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Live Submission:     Yes / No" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Re-submission:        Yes / No" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Sub-return No." & "{x=40}Volume No." & "{x=70}Your reference" & vbCrLf & vbCrLf)
              
  Call rep.Out(vbCrLf & vbCrLf & vbCrLf & vbCrLf)
  
  Call rep.Out("{x=6}{Arial=14,bn}TAPE INFORMATION" & "{Arial=10,nb} - This section must be completed by Tape and Cartridge Users" & "{Arial=14,nb}" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Media Type:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Encoding:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Block Size:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Block format - variable or fixed:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}No. of Headers:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Contact Name:" & "{x=40}{Arial=12,n}" & "" & vbCrLf & vbCrLf & _
              "{x=8}{Arial=12,bn}Position:" & "{x=40}{Arial=12,n}" & "" & vbCrLf)
               
Report_MM_SUBMISSION_end:
  Call xReturn("Report_MM_SUBMISSION")
  Exit Function
  
Report_MM_SUBMISSION_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "Report_MM_SUBMISSION", "Expenses and Benefits Return Submission Document", "Error printing Expenses and Benefits Return Submission Document")
  Resume Report_MM_SUBMISSION_end

                       
End Function

Public Function GetCO2DisplayFigure(CompanyCar As IBenefitClass) As String
'Return C02 figure for display purposes on reports
  On Error GoTo GetCO2DisplayFigure_Err
  Call xSet("GetCO2DisplayFigure")
  Dim CompanyCar_CO2 As String
  
  'If CompanyCar is Electric then display a value of 0
  If (CompanyCar Is Nothing) Then
    CompanyCar_CO2 = ""
  Else
    If GetBenItem(CompanyCar, car_p46FuelType_db) = CCFT_ELECTRIC Then
      CompanyCar_CO2 = "0"
    Else
      CompanyCar_CO2 = GetBenItem(CompanyCar, car_p46CarbonDioxide_db)
    End If
  End If

  GetCO2DisplayFigure = CompanyCar_CO2
  
GetCO2DisplayFigure_End:
  Call xReturn("GetCO2DisplayFigure")
  Exit Function

GetCO2DisplayFigure_Err:
  ' Call ErrorMessage(ERR_ERROR, Err, "GetCO2DisplayFigure", "Error in GetCO2DisplayFigure", "Undefined error.")
  Resume GetCO2DisplayFigure_End
End Function

Private Function GetApprovedZeroEmissionsMileage(CompanyCar As IBenefitClass) As String
  On Error GoTo err_err
  Call xSet("GetApprovedZeroEmissionsMileage")
  Dim value As String
    
  If (CompanyCar Is Nothing) Then
    value = ""
  Else
    If Not GetBenItem(CompanyCar, car_ElectricRangeMiles_NeedToInput) Then
      value = ""
    Else
      value = GetBenItem(CompanyCar, car_ElectricRangeMiles_db)
    End If
  End If

  GetApprovedZeroEmissionsMileage = value
  
err_end:
  Call xReturn("GetApprovedZeroEmissionsMileage")
  Exit Function

err_err:
  Resume err_end
End Function


Public Function IsManagementReport(ByVal pr As P11D_REPORTS) As Boolean
  
  IsManagementReport = (pr >= [RPT_FIRST_MANAGEMENT]) And (pr <= [RPT_LAST_MANAGEMENT]) And Not p11d32.ReportPrint.GroupHeader(pr)
End Function
Public Sub ManagementReportsToTree(ByVal tvwReports As TreeView)
  Dim i As Long
  
  Call ReportToTree(tvwReports, RPT_MANAGEMENT)
  For i = [RPT_FIRST_MANAGEMENT] To [RPT_LAST_MANAGEMENT]
    Call ReportToTree(tvwReports, i) 'km
    If (IsManagementReport(i)) Then
     If Not (Is83FileName(p11d32.ReportPrint.ManagementReportPathAndFile(i))) Then Call Err.Raise(ERR_FILE_INVALID, "ReportsToTree", "The management report file " & p11d32.ReportPrint.ManagementReportPathAndFile(i) & " is not 8.3 format.")
    End If
  Next i
End Sub
Public Sub QAManagementReports()
  Dim i As Long
  Dim qs As QString
  Dim sFileName
  Set qs = New QString

On Error GoTo err_err

  If (p11d32.CurrentEmployer Is Nothing) Then
    If p11d32.Employers.CountValid = 0 Then
      Call Err.Raise(ERR_EMPLOYER_INVALID, "ManagementReportFilesPresent", "At least one employer is needed to continue the test")
    End If
    Call p11d32.LoadEmployer(p11d32.Employers(1), True, True, False)
  End If
    
  For i = [RPT_FIRST_MANAGEMENT] To [RPT_LAST_MANAGEMENT]
    If (IsManagementReport(i)) Then
      sFileName = p11d32.ReportPrint.ManagementReportPathAndFile(i)
      If Not FileExists(sFileName) Then
        Call qs.Append(sFileName & vbCrLf)
      End If
    End If
  Next
  
  If (qs.Length > 0) Then
    Call MsgBox("The following report files are not present:" & vbCrLf & vbCrLf & qs.bstr)
  Else
    If MsgBox("Do you wish to test preview the management reports?", vbYesNo, "Test preview") = vbYes Then
      For i = [RPT_FIRST_MANAGEMENT] To [RPT_LAST_MANAGEMENT]
        If (IsManagementReport(i)) Then
          Call p11d32.ReportPrint.DoWizardReport(i, PREPARE_REPORT)
        End If
      Next
    End If
  End If

err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "ManagementReportFilesPresent", "QA ManagementReport", Err.Description)
End Sub
Public Sub ReportsUserToTree(ByVal tvwReports As TreeView)
  Dim sDir As String
  
  On Error Resume Next
  
  Call tvwReports.nodes.Remove(S_MAKE_KEY & p11d32.ReportPrint.Name(RPT_USER))
  
  On Error GoTo ReportsUserToTree_ERR
  
  Call xSet("ReportsUserToTree")
  
  'remove before
  Call ReportToTree(tvwReports, RPT_USER, , True)
    
  sDir = p11d32.ReportPrint.ReportPathUser
  Call EnumFiles(tvwReports, p11d32.ReportPrint.ReportPathUser, "*" & S_REPORT_FILE_EXTENSION, p11d32.ReportPrint)
  
ReportsUserToTree_END:
  Call xReturn("ReportsUserToTree")
  Exit Sub
ReportsUserToTree_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportsUserToTree", "Reports User To Tree", "Error placing the user reports to the tree, directoy = " & sDir)
  Resume ReportsUserToTree_END
End Sub
Public Sub ReportsImageListSet(ByVal tvwReports As TreeView)
  Set tvwReports.ImageList = MDIMain.imlTree
End Sub
Public Sub ReportsSelectNodeImage(ByRef nodeLastSelected As node, n As node, Optional bReset = False)
  On Error GoTo SelectNodeImage_ERR
  
  Call xSet("SelectNodeImage")
  
  If bReset Then Set nodeLastSelected = Nothing
  
  If n Is Nothing Then GoTo SelectNodeImage_END
 
  If Not nodeLastSelected Is n Then
    If Not nodeLastSelected Is Nothing Then
      nodeLastSelected.Image = IMG_UNSELECTED
    End If
  End If
  Set nodeLastSelected = n
  n.Image = IMG_SELECTED
  
SelectNodeImage_END:
  Call xReturn("SelectNodeImage")
  Exit Sub
SelectNodeImage_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SelectNodeImage", "Select Node Image", "Error setting a new nodes image.")
  Resume SelectNodeImage_END
End Sub
Public Function SelectDefaultReport(ByRef nodeLastSelected As node, ByVal tvwReports As TreeView) As Boolean
  If tvwReports.nodes.count > 1 And tvwReports.SelectedItem Is Nothing Then
    Set tvwReports.SelectedItem = tvwReports.nodes(1)
  End If
  
  If Not tvwReports.SelectedItem Is Nothing Then
    Call ReportsSelectNodeImage(nodeLastSelected, tvwReports.SelectedItem, True)
    SelectDefaultReport = True
    
  End If
  
End Function
Public Function ReportsToTreeEnd(tvwReports As TreeView, ByRef nodeLastSelected As node) As Boolean
  Call ReportsUserToTree(tvwReports)
  ReportsToTreeEnd = SelectDefaultReport(nodeLastSelected, tvwReports)
End Function

Public Sub ReportToTree(ByVal tvwReports As TreeView, ByVal rpt As P11D_REPORTS, Optional sCurrentName As String = "", Optional ByVal bExpand As Boolean = False)
  Dim n As node
  Dim bUserReport As Boolean
  Dim sParentName As String
  Dim sUserReport As String
  
  On Error GoTo ReportToTree_ERR
  
  Call xSet("ReportToTree")
  
  bUserReport = (RPTT_USER = p11d32.ReportPrint.ReportType(rpt))
  If bUserReport Then
    sParentName = p11d32.ReportPrint.Name(RPT_USER)
  Else
    sParentName = p11d32.ReportPrint.ParentName(rpt)
    sCurrentName = p11d32.ReportPrint.Name(rpt)
  End If
  If (sCurrentName = "") Then
    sCurrentName = ""
  End If
  
  
  
  If Len(sParentName) > 0 Then
    sParentName = S_MAKE_KEY & sParentName
    Set n = tvwReports.nodes.Add(sParentName, tvwChild, S_MAKE_KEY & sCurrentName, sCurrentName)
  Else
    If (p11d32.ReportPrint.EnableEmailReports = False) And (rpt = RPT_EMPLOYEE_LETTER_EMAIL) Then GoTo ReportToTree_END
    Set n = tvwReports.nodes.Add(, , S_MAKE_KEY & sCurrentName, sCurrentName)
  End If
  
  'stick image
  If bUserReport Then
    GoTo NORMAL
  Else
    If Not p11d32.ReportPrint.GroupHeader(rpt) Then
NORMAL:
      n.Tag = rpt
      n.Image = IMG_UNSELECTED
    Else
      If (bExpand) Then
        n.Expanded = True
      End If
      
      If n.Expanded Then
        n.Image = IMG_FOLDER_OPEN
      Else
        n.Image = IMG_FOLDER_CLOSED
      End If
    End If
  End If
  
  If Not tvwReports.SelectedItem Is Nothing Then GoTo ReportToTree_END
  
  If n.Tag = L_REPORT_USER_TAG Then 'user report
    If tvwReports Is F_Print.tvwReports Then
      sUserReport = p11d32.ReportPrint.UserReportFileLessExtension
    Else
      sUserReport = p11d32.ReportPrint.UserReportSelectEmployeeFileLessExtension
    End If

    For Each n In tvwReports.nodes
      If StrComp(n.Text, sUserReport) = 0 Then
        Set tvwReports.SelectedItem = n
        Exit For
      End If
    Next
  Else
    If tvwReports Is F_Print.tvwReports Then
      If n.Tag = p11d32.ReportPrint.DefaultReportIndex Then
        Set tvwReports.SelectedItem = n
      End If
    Else
      If n.Tag = p11d32.ReportPrint.DefaultSelectEmployeeReportIndex Then
        Set tvwReports.SelectedItem = n
      End If
    End If
    
  End If
  
  
ReportToTree_END:
  Call xReturn("ReportToTree")
  Exit Sub
ReportToTree_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportToTree", "Report To Tree", "Error placing a report to the tree.")
  Resume ReportToTree_END
  Resume
End Sub

Public Function ReportPercentage(dDecimalPercentage As Double) As String
  Dim s As String
  dDecimalPercentage = Round(dDecimalPercentage * 100, 1)
  s = Format$(dDecimalPercentage, "0.0")
  ReportPercentage = s & "%"
End Function

Private Sub OPRAWorkingPaperValueValues(ByVal rep As Reporter, ByVal isOPRABenefit As Boolean, ByVal opraAmountForegone As Variant, ByVal valueNonOpra, ByVal value)

   Call WKOut(rep, WK_BLANK_LINE)
  
  
  If isOPRABenefit And (opraAmountForegone > 0) Then
    Call WKOut(rep, WK_ITEM_TEXT, "Value of benefit (non OpRA)", valueNonOpra)
    Call WKOut(rep, WK_BLANK_LINE)
    Call WKOut(rep, WK_ITEM_TEXT, S_UDM_OPRA_AMOUNT_FOREGONE, opraAmountForegone)
    Call WKOut(rep, WK_BLANK_LINE)
    Call WKOut(rep, WK_ITEM_subtotal, "Value of benefit (greater of OpRA amount foregone " & FormatWNRPT(opraAmountForegone) & " and the non OpRA value " & FormatWNRPT(valueNonOpra) & ")", value)
    
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "Value of benefit", value)
    
  End If
  Call WKOut(rep, WK_BLANK_LINE)

End Sub
Public Sub OPRAWorkingPaperValue(ByVal ben As IBenefitClass, ByVal rep As Reporter)
  Call OPRAWorkingPaperValueValues(rep, IsOpRABenefitClass(ben.BenefitClass), ben.value(ITEM_OPRA_AMOUNT_FOREGONE), ben.value(ITEM_VALUE_NON_OPRA), ben.value(ITEM_VALUE))
End Sub
Public Sub EmployeeLetterAddCodes(ByVal ce As codeEditor, ByVal exportFileNameEditor As Boolean)
  Dim s As String, sParentName As String
  Dim i As Long
  Dim vbm As VBMenu
  Dim vbmi As VBMenuItem
  Dim slParentNames As StringList
  
  Set slParentNames = New StringList
  
  Set vbmi = ce.Menu.Add(S_ELMC_MASTER, "&" & S_ELMC_MASTER, "")
  
  For i = [_ELMC_FIRST_ITEM] To [_ELMC_LAST_ITEM]
    
     If exportFileNameEditor And i <> ELMC_EMPLOYER And i <> ELMC_EMPLOYEE And i <> ELMC_EMPLOYEE And i <> ELMC_DATES Then
       GoTo NEXT_TOP_LEVEL_ITEM
     End If
    
    s = EmployeeLetterMenuCaptions(i)
    Call slParentNames.Add(s)
    Call ce.Menu.Add(s, s, S_ELMC_MASTER)
NEXT_TOP_LEVEL_ITEM:
  Next
  
  For i = EMPLOYEE_LETTER_CODE.ELC_FIRST_ITEM To EMPLOYEE_LETTER_CODE.ELC_LAST_ITEM
    If (exportFileNameEditor And i = EMPLOYEE_LETTER_CODE.ELC_ADDRESS) Then
      GoTo NEXT_ITEM
    End If
    
    sParentName = EmployeeLetterCode(i, ELCT_MENU_PARENT, False)
    If (slParentNames.IsPresent(sParentName)) Then
      s = EmployeeLetterCode(i, ELCT_MENU_CAPTION, False)
      Set vbmi = ce.Menu.Add(s, s, EmployeeLetterCode(i, ELCT_MENU_PARENT, False))
      s = EmployeeLetterCode(i, ELCT_LETTER_FILE_CODES, False)
      Call ce.AddCode(s)
      vbmi.Tag = i
    End If
NEXT_ITEM:
  Next
End Sub
Public Sub Report_P11Db_Reconciliation(rep As Reporter, ey As Employer)
  Dim benEy As IBenefitClass
  Dim vTotalBenfitsExClass1 As Variant
  Dim vTotalBenfitsClass1 As Variant
  Dim vTotalBenefits As Variant
  
  
  Const S_SUB_ITEM_MARKER As String = " - "
  
  
On Error GoTo err_err
    
  Set benEy = ey
  Call benEy.Calculate
  
  Call rep.Out("{Arial=12,bi}" & S_P11D_B & " reconciliation - " & p11d32.Rates.value(TaxFormYear) & ", Employer: " & ey.Name & vbCrLf)
  
'section p11d benfits
  Call WKOut(rep, WK_SECTION_BREAK)
  
  Call P11Db_ReconciliationSubHeading(rep, UpperCaseFirstLetter(S_P11DB_SHORT_CAPTION_BOX_A))
  
  Call P11Db_ReconciliationAdjustmentsColOffsets
  
  
  Call WKTblColFormats("ib", "ib", "irb", "irb")
  Call WKTableHeadings(rep, "P11D box", "Class 1A NIC benefits:", "", "£")
  Call WKTblColFormats("n", "n", "rn", "rn")
  
  Dim HMITSection As HMIT_SECTIONS
  
  vTotalBenefits = 0
  For HMITSection = HMIT_SECTIONS.HMIT_FIRST_ITEM To HMIT_SECTIONS.HMIT_LAST_ITEM
    vTotalBenefits = vTotalBenefits + P11Db_ReconciliationTableLineP11DBox(rep, benEy, HMITSection, HMIT_SECTION_EMPLOYER_BENEFIT_ITEM_P11D_NICABLE)
  Next
  
  Call WKOut(rep, WK_ITEM_Total, "", vTotalBenefits, , False)
  
  
  Call WKOut(rep, WK_SECTION_BREAK)
  
  
  
'A
  Call WKTblColFormats("nib", "nib", "nirb", "nirb")
  Call WKTableHeadings(rep, S_P11D_B & "~box", "", "£", "£")
  Call WKTblColFormats("n", "n", "rn", "rn")
  Call WKTableRow(rep, "A", UpperCaseFirstLetter(S_P11DB_SHORT_CAPTION_BOX_A), "", FormatWNNoCurrency(benEy.value(employer_TotalBenefitsPotentiallySubjectToClass1A)))
  Call WKTableBlankRow(rep)
'B

  Call WKTableRow(rep, "B", S_P11DB_ADDITIONS_BOX_B, "", "")
  Call WKTableBlankRow(rep)
  'benefits payrolled, made good subject to PAYE
  vTotalBenfitsExClass1 = 0
  For HMITSection = HMIT_SECTIONS.HMIT_FIRST_ITEM To HMIT_SECTIONS.HMIT_LAST_ITEM
    Call P11Db_ReconciliationTableLineP11DBox(rep, benEy, HMITSection, HMIT_SECTION_EMPLOYER_BENEFIT_ITEM_P11DB_ADDITIONS, False, True, , 3)
  Next
  
  Call P11Db_ReconciliationAdjustments(ey, rep, ey.P11DbAdditions, False)
  Call WKTableBlankRow(rep)
  Call WKTableRow(rep, "", P11DbAdditionsDescription(benEy, False), "", FormatWNNoCurrency(benEy.value(employer_NIC_AdjustmentAddTotal)))
  
'C
  Call WKTableBlankRow(rep)
  Call WKTableRow(rep, "C", S_P11DB_DEDUCTIONS_BOX_C, "", "")
   
  Call WKTableBlankRow(rep)
  For HMITSection = HMIT_SECTIONS.HMIT_FIRST_ITEM To HMIT_SECTIONS.HMIT_LAST_ITEM
    Call P11Db_ReconciliationTableLineP11DBox(rep, benEy, HMITSection, HMIT_SECTION_EMPLOYER_BENEFIT_ITEM_P11DB_DEDUCTIONS, True, True, , 3)
  Next
  
  Call P11Db_ReconciliationAdjustments(ey, rep, ey.P11DbDeductions, True)
  Call WKTableBlankRow(rep)
  Call WKTableRow(rep, "", P11DbDeductionsDescription(benEy), "", FormatWNNoCurrency(-1 * benEy.value(employer_NIC_AdjustmentDeductTotal)))
  Call WKTableBlankRow(rep)
  Call WKOut(rep, WK_ITEM_Total, "Sub total", benEy.value(employer_NIC_BenefitsAndAdjustmentsSubTotal), , True)
  
  
  
'D
  Call WKTableBlankRow(rep)
  
  Call WKTableRow(rep, "D", S_P11DB_TOTAL_BOX_D & " (minimum 0)", "", FormatWNNoCurrency(benEy.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A)))
  
  Call WKTableBlankRow(rep)
'E
  Call WKTableRow(rep, "E", S_P11DB_MUTIPLY_BY_NIC_RATE_BOX_E, "", "x " & Report_P11db_NIC_Rate_Formatted())
  Call WKTableBlankRow(rep)
'F
  Call WKTableRow(rep, "F", S_P11DB_CLASS_1A_PAYABLE_BOX_F, "", FormatWNNoCurrency(benEy.value(ITEM_NIC_CLASS1A_BENEFIT), , True))
  Call WKOut(rep, WK_SECTION_BREAK)
    
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "P11DbReconciliation", "P11DbReconciliation", "Error printing the P11db reconciliation")
  Resume err_end
  Resume
End Sub
Private Sub P11Db_ReconciliationAdjustmentsColOffsets()
  Call WKTblColXOffsets(L_WK_OTHER_TABLE_COL1, 12, L_WK_OTHER_TABLE_COL3, L_WK_OTHER_TABLE_COL4)

End Sub
Private Sub P11Db_ReconciliationAdjustments(ey As Employer, rep As Reporter, Adjustments As ObjectList, negative As Boolean)
  Dim i As Long
  Dim adjustment As P11DbAdjustment
  
  For i = 1 To Adjustments.count
    Set adjustment = Adjustments(i)
    Call WKTableRow(rep, "", adjustment.caption, FormatWNNoCurrency(adjustment.value, negative), "")
  Next
End Sub
Private Sub P11Db_ReconciliationSubHeading(rep As Reporter, subHeading As String)
  Call WKTblColXOffsets(0)
  Call WKTblColFormats("lb")
  Call WKTableHeadings(rep, subHeading)
  Call P11Db_ReconciliationAdjustmentsColOffsets
End Sub

Function P11Db_ReconciliationTableLineP11DBox(rep As Reporter, benEy As IBenefitClass, HMITSection As HMIT_SECTIONS, sectionValue As HMIT_SECTION_VALUE, Optional ByVal negate As Boolean = False, Optional ByVal boxAndNameInColummn2 As Boolean = False, Optional ByVal repOut As Boolean = True, Optional ByVal valueColumn As Long = 4) As Variant

  Dim boxString As String
  Dim boxName As String
  Dim boxValue As String
  Dim eeItem As EmployeeItems
  Dim v As Variant
  Dim captionAddition As String
  
  eeItem = p11d32.Rates.HMITSectionToValue(HMITSection, sectionValue)
  
  
  If sectionValue = HMIT_SECTION_EMPLOYER_BENEFIT_ITEM_P11DB_ADDITIONS Then
    captionAddition = ":" & LowerCaseFirstLetter(S_P11DB_AMOUNTS_INFORMALLY_TAXED_THROUGH_PAYROLL)
  ElseIf sectionValue = HMIT_SECTION_EMPLOYER_BENEFIT_ITEM_P11DB_DEDUCTIONS Then
    captionAddition = ": " & LowerCaseFirstLetter(S_P11DB_EMPLOYEES_NOT_SUBJECT_TO_CLASS_1A)
  End If
  
  v = benEy.value(eeItem)
  If IsNumeric(v) And v <> 0 Then
    If (repOut) Then
      If (boxAndNameInColummn2) Then
        boxName = p11d32.Rates.HMITSectionToValue(HMITSection, HMIT_SECTION_BOX_AND_DESCRIPTION)
      Else
        boxString = p11d32.Rates.HMITSectionToValue(HMITSection, HMIT_SECTION_BOX_LETTER)
        boxName = p11d32.Rates.HMITSectionToValue(HMITSection, HMIT_SECTION_DESCRIPTION)
      End If
      boxName = boxName & captionAddition
      boxValue = FormatWNNoCurrency(benEy.value(eeItem), negate)
      If (valueColumn = 3) Then
        Call WKTableRow(rep, boxString, boxName, boxValue, "")
      Else
        Call WKTableRow(rep, boxString, boxName, "", boxValue)
      End If
    End If
    P11Db_ReconciliationTableLineP11DBox = v
  Else
    P11Db_ReconciliationTableLineP11DBox = 0
  End If
    
  
End Function

