VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IBenefitClass
Implements ISortFunction

Public Enum LoanItems
'MP DB - what is mapping between this Enum and Other Enum
'        encoured errors when removed unused enums here
  ln_item_db = ITEM_DESC
  ln_Value = ITEM_VALUE
  ln_MadeGood_Net = ITEM_MADEGOOD_NET
  ln_Benefit = ITEM_BENEFIT
  ln_InterestPaid_db = ITEM_MADEGOOD
  ln_BenefitMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  ln_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  ln_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  ln_BoxNumber = ITEM_BOX_NUMBER
  ln_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  ln_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  ln_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  ln_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  ln_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  ln_Error = ITEM_ERROR
  ln_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  ln_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  ln_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE
  
  
  'MAPPINGS TO N Other
  ln_WAIVEDitem
  ln_WAIVED_SPECIFIC_START = ln_WAIVEDitem
  ln_WAIVEDValue
  ln_WAIVEDMadeGood_NET           'MP DB
  ln_WAIVEDBenefit
  ln_WAIVEDMadeGood               'MP DB
  ln_WAIVEDActualAmountMadeGood   'MP DB (not used)
  ln_WAIVEDBenefit_Reportable
  ln_WAIVEDUDM_BENEFIT_TITLE      'MP DB (not used)
  ln_WAIVEDBoxNumber              'MP DB (not used)
  ln_WAIVEDMadeGoodIsTaxDeducted  'MP DB (not used)
  ln_WAIVEDClass1AAdjustment      'MP DB (not used)
  ln_WAIVEDNICClass1AAble         'MP DB (not used)
  ln_WAIVEDNIC_Class1A_Value      'MP DB (not used)
  ln_WAIVEDNIC_Class1A_Benefit    'MP DB (not used)
  ln_WAIVEDError
  ln_WAIVEDValue_Non_OPRA
  ln_WAIVEDOPRA_Ammount_Foregone_Used_For_Value
  ln_WAIVEDOPRA_Ammount_Foregone_db
  
  
  ln_WAIVED_IRDesc                'MP DB (not used)
  ln_WAIVEDGrossAmountPaidByEmployer
  ln_WAIVEDEmployeeReference          'MP DB (not used)
  ln_WAIVEDcategory                   'MP DB (not used)
  ln_WAIVEDHMITCode_UDBCode           'MP DB (not used)
  ln_WAIVEDCompanyDefinedCategory     'MP DB (not used)
  ln_WAIVEDCompanyDefinedCategoryKey  'MP DB (not used)
  ln_WAIVEDclass                      'MP DB (not used)
  ln_WAIVEDavailablefrom
  ln_WAIVEDavailableto
  ln_amountwaived_db
  
  ln_WAIVED_SPECIFIC_END = ln_amountwaived_db
  'end mappings
  
  ln_balkey_db
  ln_OpenOutstanding
  ln_CloseOutstanding
  ln_ZeroBalanceAtEndOfYear
  ln_MaxOutstandingAtAnyPoint
  ln_DateDischargedOrMade
  ln_DischargedDate
  ln_MadeDate
  ln_BalancesAmmended
  ln_AverageOutstanding
  ln_AverageLoanRate
  
  ln_WholeTaxMonthsOutstanding
  ln_CashEquivalentNormal
  ln_CashEquivalentDaily
  ln_InterestNormal
  ln_InterestDaily
  
  ln_DeemedInterestNormal
  ln_DeemedInterestAlternative
  

  ln_DidLoanCommenceOnFirstDayOfTaxYear_db
  
  ln_DefaultMirasRelief
  ln_Deemed               'MP DB (not used)
    
  ln_nborrowers_db 'bf    'MP DB ToDo - don't think in active use (not displayed on form) - investigate
  ln_LoanType_db 'bf      'MP DB ToDo - chk if in use
  ln_LoanCurrency_db 'bf
  ln_LoanCurrencyIndex 'bf
  ln_Lender 'bf
  ln_ACNumber 'bf
  ln_Miras 'bf            'MP DB (not used)
  
  ln_opendate
  ln_closedate
  
  ln_UseDailyCalculationOnly_db
  ln_CheapTaxableLoan_db
  
  'HMIT Specific
  ln_HMIT_NoOfBorrowers
  
  'Last Item
  ln_LASTITEM = ln_HMIT_NoOfBorrowers
End Enum

Public Enum PaymentType
  T_NULL = 0
  T_NOTHING = 1
  T_PAYMENT = 2
  T_RECEIPT = 3
End Enum

Private m_Parent As IBenefitClass
Private m_ReadFromDB As Boolean
Private m_sbookmark As String
Private m_BenItems(1 To ln_LASTITEM)

Private m_dirty  As Boolean
Private m_InvalidFields As Long
Private m_BenClass As BEN_CLASS
Private m_NeedToCalculate As Boolean

Public NOther As other
Public BalanceSheet As ObjectList
Private Sub Class_Terminate()
 'Debug.Print"Loan - terminate"
End Sub
Public Sub BalanceSheetItemsAdd(OL As ObjectList)
  Dim bi As BalanceItem, biexisting As BalanceItem
  Dim i As Long, l As Long
  Dim bFound As Boolean
  
  For l = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(l)
    If Not bi Is Nothing Then
      bFound = False
      For i = 1 To OL.Count
        Set biexisting = OL.Item(i)
        If bi.DateFrom = biexisting.DateFrom Then
          bFound = True
          Exit For
        End If
      Next
      If (Not bFound) Then
        Call OL.Add(bi)
      End If
    End If
  Next
End Sub

Private Function IBenefitClass_CalculateBody() As Variant

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Loan"
End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = m_NeedToCalculate
End Property

Private Property Let IBenefitClass_LinkBen(RHS As Boolean)
  IBenefitClass_NeedToCalculate = RHS
End Property

Private Property Get IBenefitClass_LinkBen() As Boolean

End Property

Private Function IBenefitClass_CanBringForward() As Boolean
  Dim ben As IBenefitClass
  Set ben = Me
  
  Call ben.Calculate
  IBenefitClass_CanBringForward = Not ben.value(ln_ZeroBalanceAtEndOfYear)
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim ben As IBenefitClass
  Dim Loan As Loan
  Dim ee As Employee
  Dim loans As loans
  
  On Error GoTo Copy_END
  Call xSet("Copy")
  
  Set ee = Parent
  If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Copy loan", "The employee is nothing.")
  
  Set loans = LoansAddToCurrentEmployee(ee)
  If loans Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Copy loan", "The loans collection for the destination employee is nothing.")
  
  Set ben = New Loan
  ben.BenefitClass = m_BenClass
  If CopyBenData(ben, Me) Then
    
    Set ben.Parent = loans
    ben.Dirty = True
    Call loans.Add(ben)
    Call CopyBenStandardEnd(ben)
    
    Set Loan = ben
    Call Loan.AddNOther
    Set IBenefitClass_Copy = ben
  End If
  
  
  
Copy_END:
  Call xReturn("Copy")
  Exit Function
Copy_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Copy", "Copy", "Error copying a company car benefit.")
  Resume Copy_ERR

End Function

'cad loan fix copy all
Public Sub AddNOther()
  Dim ben As IBenefitClass
  Dim benLoan As IBenefitClass
  Dim ee As Employee
  
  Static bCreatedNOther As Boolean
  
  On Error GoTo AddNOther_ERR
  
  Call xSet("AddNOther")
  
  If Not bCreatedNOther Then
    bCreatedNOther = True
    Set NOther = New other
    Set benLoan = Me
    Set NOther.Loan = benLoan
    Set ben = NOther
    ben.LinkBen = True
    Set ben.Parent = m_Parent.Parent
    ben.ReadFromDB = True
    ben.BenefitClass = BC_NON_CLASS_1A_M
    ben.value(IRDescriptionBenItem(ben.BenefitClass)) = S_IR_DESC_M_NC1A_LOANS_WRIT_WAIV
    
    Set ee = ben.Parent
    Call ee.benefits.Add(NOther)
    Call StandardReadData(NOther)
    'cad loan fix copy whole of function
    benLoan.value(ln_WAIVEDNICClass1AAble) = False
    
  End If

AddNOther_END:
  Call xReturn("AddNOther")
  Exit Sub
AddNOther_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddNOther", "Add N Other", "Error adding an N other benefit with a loan benefit.")
  Resume AddNOther_END
  Resume
End Sub

'km
Private Property Get LoanTypeString() As String
    LoanTypeString = S_BENEFICIAL_LOANS_TYPE
  
End Property



Public Function BalanceOnDate(dDate As Date, lLastIndex As Long) As Double
  Dim l As Long
  Dim bi As BalanceItem
  Dim BItemp As BalanceItem
  Dim lLastIndexTemp As Long
  'all the BalanceItems payement are assumed to be in data order
On Error GoTo BalanceOnDate_ERR
  
  Call xSet("BalanceOnDate")
  
  lLastIndexTemp = lLastIndex + 1
  For l = lLastIndexTemp To BalanceSheet.Count
    Set bi = BalanceSheet(l)
    If Not bi Is Nothing Then
      If BItemp Is Nothing Then
        If dDate >= bi.DateFrom And dDate <= bi.DateTo Then
          BalanceOnDate = bi.Balance
          Set BItemp = bi
          lLastIndex = l
        End If
      Else
        If bi.DateFrom = BItemp.DateFrom Then
          BalanceOnDate = BalanceOnDate + bi.Balance
          lLastIndex = l
        Else
          Exit For
        End If
      End If
    End If
  Next
  
BalanceOnDate_END:
  Call xReturn("BalanceOnDate")
  Exit Function
BalanceOnDate_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BalanceOnDate", "Balance On Date", "Error getting the balance on " & dDate & " for " & m_BenItems(ln_item_db) & ".")
  Resume BalanceOnDate_END
End Function

Private Sub Class_Initialize()
  Set BalanceSheet = New ObjectList
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = m_BenClass
End Property

Private Function IBenefitClass_Calculate() As Variant
  Dim ben As IBenefitClass
  
On Error GoTo Calculate_ERR

  Call xSet("Calculate")
  Set ben = Me
  Call m_Parent.Calculate
  
  IBenefitClass_Calculate = m_BenItems(ln_Benefit)
  
  'CAD2004R1 what is this ?
  ben.value(ln_HMIT_NoOfBorrowers) = ""
  
  
Calculate_END:
  Call xSet("Calculate")
  Exit Function
Calculate_ERR:
  IBenefitClass_Calculate = S_ERROR
  Resume Calculate_END
  Resume
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)

End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean

End Property




Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  On Error GoTo Loan_DeleteDB_Err
  
  Call xSet("Loan_DeleteDB")
    Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_BENLOAN)
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
    Call p11d32.CurrentEmployer.db.Execute(sql.Queries(DELETE_LOANBAL, m_BenItems(ln_balkey_db), LoanTypeString))
  End If
  
  IBenefitClass_DeleteDB = True
Loan_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("Loan_DeleteDB")
  Exit Function
  
Loan_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Loan_DeleteDB", "Loan Delete DB", "Error deleting the loan benefit.")
  Resume Loan_DeleteDB_End
  Resume
End Function

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property
Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  'other link items
  ben.value(ln_WAIVEDavailablefrom) = p11d32.Rates.value(TaxYearStart)
  ben.value(ln_WAIVEDavailableto) = p11d32.Rates.value(TaxYearEnd)
  
  ben.value(ln_WAIVEDGrossAmountPaidByEmployer) = 0
  ben.value(ln_WAIVEDValue) = 0
  ben.value(ln_WAIVEDBenefit) = 0
  
  ben.value(ln_WAIVEDMadeGood) = 0            'MP DB - ToDo - chk still saves ok after removal
  ben.value(ln_WAIVEDMadeGood_NET) = 0        'MP DB - ToDo
  ben.value(ln_WAIVEDBenefit_Reportable) = False
  ben.value(ln_WAIVEDError) = ""
  ben.value(ln_OPRA_Ammount_Foregone_Used_For_Value) = False
  
  
  'end other link items

  ben.value(ln_OpenOutstanding) = 0
  ben.value(ln_CloseOutstanding) = 0
  ben.value(ln_DateDischargedOrMade) = UNDATED
  ben.value(ln_DischargedDate) = UNDATED
  ben.value(ln_MadeDate) = UNDATED
  ben.value(ln_MaxOutstandingAtAnyPoint) = 0
  ben.value(ln_opendate) = UNDATED
  ben.value(ln_closedate) = UNDATED
  ben.value(ln_ZeroBalanceAtEndOfYear) = False
  ben.value(ln_AverageOutstanding) = 0
  ben.value(ln_InterestDaily) = 0
  ben.value(ln_AverageLoanRate) = 0
  ben.value(ln_WholeTaxMonthsOutstanding) = 0
  ben.value(ln_InterestNormal) = 0
  ben.value(ln_CashEquivalentNormal) = 0
  ben.value(ln_CashEquivalentDaily) = 0
  ben.value(ln_Benefit) = 0
  ben.value(ln_MadeGood_Net) = 0
  ben.value(ln_MadeGood_Net) = False
  ben.value(ln_Benefit_Reportable) = False
  ben.value(ln_DeemedInterestNormal) = 0
  ben.value(ln_DeemedInterestAlternative) = 0
  ben.value(ln_Error) = ""
  ben.value(ln_NIC_Class1A_Benefit) = 0             'MP DB - ToDo
  ben.value(ln_NIC_Class1A_Value) = 0            'MP DB - ToDo
  
  
End Sub

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property
Private Sub IBenefitClass_Kill()
  If Not NOther Is Nothing Then Set NOther.Loan = Nothing
  
  Set m_Parent = Nothing
End Sub

Private Property Get IBenefitClass_Name() As String
 IBenefitClass_Name = m_BenItems(ln_item_db)
End Property
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property
Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
Public Property Get IBenefitClass_PrintHeader() As String
  IBenefitClass_PrintHeader = ""
End Property
Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  IBenefitClass_PrintWk = PrintWKHelper(rep, Me)
End Function

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim i As Long, s As String
  
  Dim ben As IBenefitClass
  Dim OL As ObjectList
  Dim li As LoanInterest
  
  On Error GoTo clsLoan_PrintWK_Err

  Set ben = Me
  
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Loan reference : ", ben.value(ln_item_db))
  
  If Not ben.value(ITEM_BENEFIT_REPORTABLE) Then
    Call WKOut(rep, WK_ITEM_TEXT_BOLD, "This loan is deminimis and has not been reported on the P11D, the following is for information only.")
    Call WKOut(rep, WK_BLANK_LINE)
  End If
    
  Call WKOut(rep, WK_ITEM_TEXT, "Amount outstanding at " & DateValReadToScreen(p11d32.Rates.value(TaxYearStart)) & " or date loan was made (if later)", ben.value(ln_OpenOutstanding), , True)
  Call WKOut(rep, WK_ITEM_TEXT, "Amount outstanding at " & DateValReadToScreen(p11d32.Rates.value(TaxYearEnd)) & " or date loan was discharged (if earlier)", ben.value(ln_CloseOutstanding))
  Call WKOut(rep, WK_ITEM_TEXT, "Maximum amount outstanding at any time in the year ", ben.value(ln_MaxOutstandingAtAnyPoint))
  Call WKOut(rep, WK_ITEM_TEXT, "Total amount of interest paid by the borrower in the year to " & DateValReadToScreen(p11d32.Rates.value(TaxYearEnd)), ben.value(ln_InterestPaid_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Date loan was made in " & p11d32.Rates.value(TaxFormYear) & " (where applicable)  - " & GetBenItem(ben, ln_MadeDate))
  Call WKOut(rep, WK_ITEM_TEXT, "Date loan was discharged in " & p11d32.Rates.value(TaxFormYear) & " (where applicable)  - " & GetBenItem(ben, ln_DischargedDate))
  
  Call WKOut(rep, WK_ITEM_TEXT, "Amount waived or written off in the year to " & DateValReadToScreen(p11d32.Rates.value(TaxYearEnd)) & " (where applicable)", ben.value(ln_amountwaived_db))
  
  If Not ben.value(ITEM_BENEFIT_REPORTABLE) Then GoTo clsLoan_PrintWK_end:
  
  Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)
   
  '************************** NORMAL method of calc *********************************************
  rep.Out (vbCrLf)
  Call WKOut(rep, WK_ITEM_TEXT_BOLD, "Normal method of Loan Calculation")
  rep.Out (vbCrLf)
  
  Call WKOut(rep, WK_ITEM_TEXT, "Loan outstanding at " & DateValReadToScreen(ben.value(ln_opendate)), ben.value(ln_OpenOutstanding), , True)
  Call WKOut(rep, WK_ITEM_TEXT, "Loan outstanding at " & DateValReadToScreen(ben.value(ln_closedate)), ben.value(ln_CloseOutstanding), , False)
  Call WKOut(rep, WK_ITEM_TEXT, "Average loan over the year ", ben.value(ln_AverageOutstanding))
  Call WKOut(rep, WK_ITEM_TEXT, "Average loan rate - " & Format$(ben.value(ln_AverageLoanRate), "Percent"))
  
  Call WKOut(rep, WK_ITEM_TEXT, "Whole tax months for which loan is treated as outstanding", ben.value(ln_WholeTaxMonthsOutstanding))
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKOut(rep, WK_ITEM_TEXT, "Amount of interest treated as payable at the official rate (" & FormatWN(ben.value(ln_AverageOutstanding)) & " @ " & Format$(ben.value(ln_AverageLoanRate), "Percent") & " x " & Format$(ben.value(ln_WholeTaxMonthsOutstanding)) & " / 12 )", ben.value(ln_InterestNormal))
  Call WKOut(rep, WK_ITEM_TEXT, "Less: Interest actually paid", ben.value(ln_InterestPaid_db), , , True)
      
  Call WKOut(rep, WK_ITEM_Total, "Cash equivalent, under the normal method", ben.value(ln_CashEquivalentNormal), "A", True)
    
  
  '************************** ALTERNATIVE method of calc *********************************************
  rep.Out (vbCrLf)
  Call WKOut(rep, WK_ITEM_TEXT_BOLD, "Alternative method of loan calculation")
  rep.Out (vbCrLf)
  ' ********************************** TABLE ***************************************
  Call BalanceSheetTable(rep, ben)
    
  Call WKOut(rep, WK_ITEM_TEXT, "Amount of interest treated as payable at the official rate", ben.value(ln_InterestDaily))
  Call WKOut(rep, WK_ITEM_TEXT, "Less: Interest actually paid", ben.value(ln_InterestPaid_db), , , True)
  Call WKOut(rep, WK_ITEM_Total, "Cash equivalent, under the alternative method", ben.value(ln_CashEquivalentDaily), "B", True)
  
  Call WKOut(rep, WK_SECTION_BREAK)
  
  Call OPRAWorkingPaperValue(ben, rep)
  
  
  Call WKOut(rep, WK_ITEM_TEXT, "Cash equivalent of benefit ", ben.value(ln_Benefit), , True)
  
  Call WKOut(rep, WK_SECTION_BREAK)
  Call WKOut(rep, WK_ITEM_TEXT, "Amount of any loans written off or waived (these are reported on the P11D form as part of the 'M Other Items' section).", ben.value(ln_amountwaived_db), , True)
  
  Call WKOut(rep, WK_SECTION_BREAK)
  
  Set OL = p11d32.Rates.GetInterestRatesInYear(ben.value(ln_LoanCurrencyIndex))
  
  If Not OL Is Nothing Then
    Call WKOut(rep, WK_SECTION_HEADER, "Official interest rates:")
    Call WKTblColXOffsets(10, 25)
    Call WKTblColFormats("rn", "rn")
    Call WKTableHeadings(rep, "From", "Rate")
    For i = 1 To OL.Count
      Set li = OL(i)
      Call WKTableRow(rep, li.dtStartDate, Format$(li.dRate, "Percent"))
    Next
    Call WKOut(rep, WK_SECTION_BREAK)
  End If
    
  Call WKOut(rep, WK_SECTION_HEADER, "Notes to the employee: ")
  
  Call WKOut(rep, WK_ITEM_TEXT, "Where the alternative method is beneficial to you ('B' is less than 'A' above), an election should be made by you, in writing to the HMRC by " & DateValReadToScreen(p11d32.Rates.value(LoanDailyElectionDue)) & ".")
  Call WKOut(rep, WK_ITEM_TEXT, "If the loan was for a qualifying purpose ( types 'A' or 'C' ), you may be able to claim tax relief by entering the details into box 15.3 of the tax return. ")
  Call WKOut(rep, WK_ITEM_TEXT, "If the loan is of type 'B', tax relief is claimed by entering the details into box 15.2 of the tax return. ")
  Call WKOut(rep, WK_ITEM_TEXT, "You should refer to leaflet 'IR145 Cheap or interest free loans provided by employers - an employee's guide' should you require further guidance")
  Call WKOut(rep, WK_ITEM_TEXT, "on whether tax relief is available on the loan. ")
  
clsLoan_PrintWK_end:
  Exit Function
clsLoan_PrintWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Loan_PRINTWK", "Loan Print Wk Body", "Error printing the loan working paper")
  Resume clsLoan_PrintWK_end
  Resume
End Function
Private Sub BalanceSheetTable(rep As Reporter, ben As IBenefitClass)
  Dim bi As BalanceItem
  Dim i As Long
  
  On Error GoTo BalanceSheetTable_ERR
  
  Call WKOut(rep, WK_ITEM_TEXT, "The number of days corresponds to the number used to calculate the beneficial interest, where the maximum loan outstanding on any day must be taken")
  Call WKOut(rep, WK_ITEM_TEXT, "as the amount outstanding for that day.")
  
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKTblColXOffsets(7, 22, 37, 52, 67, 82, 97)
  Call WKTblColFormats("n", "n", "rn", "rn", "rn", "rn", "rn")
  Call WKTableHeadings(rep, "From", "To", "Days", "Payment (" & S_CURRENCY & ")", "Balance (" & S_CURRENCY & ")", "Interest rate", "Interest (" & S_CURRENCY & ")")
  
  For i = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(i)
    If Not bi Is Nothing Then Call WKTableRow(rep, bi.DateFrom, bi.DateTo, bi.days, bi.Payment, bi.Balance, Format$(bi.InterestRate(ben.value(ln_LoanCurrencyIndex)), "Percent"), FormatWN(bi.Interest(ben.value(ln_LoanCurrencyIndex)), "", , True))
  Next
  
  Call WKOut(rep, WK_ITEM_subtotal)
  
BalanceSheetTable_END:
  Exit Sub
BalanceSheetTable_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BalanceSheetTable", "Balance Sheet Table", "Error printing balance sheet table.")
  Resume BalanceSheetTable_END
End Sub
Private Function IBenefitClass_ReadDB() As Long
  Dim db As Database, rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo clsLoan_ReadDB_Err
  
  xSet "clsLoan_ReadDB"
  
  If m_ReadFromDB Then GoTo clsLoan_ReadDB_end
  
  Set ben = Me
  
    Set rs = m_Parent.Parent.Parent.rsBenTables(TBL_BENLOAN)
  
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    rs.Bookmark = m_sbookmark
    
    Call StandardReadData(ben, rs)

    ben.value(ln_item_db) = "" & rs.Fields("ITEM").value
    ben.value(ln_LoanCurrency_db) = "" & rs.Fields("CURR").value
    Call SetLoanCurrenyIndex
    ben.value(ln_amountwaived_db) = IIf(IsNull(rs.Fields("WAIVED").value), 0, rs.Fields("WAIVED").value)
    ben.value(ln_balkey_db) = IIf(IsNull(rs.Fields("BALKEY").value), 0, rs.Fields("BALKEY").value)
    
    ben.value(ln_InterestPaid_db) = IIf(IsNull(rs.Fields("Interest").value), 0, rs.Fields("Interest").value)
    ben.value(ln_UseDailyCalculationOnly_db) = IIf(IsNull(rs.Fields("DailyOnly").value), 0, rs.Fields("DailyOnly").value)
    'cadloan
    
    ben.value(ln_Lender) = "LENDER"
    ben.value(ln_ACNumber) = "DEFAULT"
    ben.value(ln_nborrowers_db) = "" & rs.Fields("BORROWERS").value
    ben.value(ln_LoanType_db) = "" & rs.Fields("LTYPE").value
  
    If Not p11d32.BringForward.Yes Then
      ben.value(ln_DidLoanCommenceOnFirstDayOfTaxYear_db) = "" & rs.Fields("DidLoanCommenceOnFirstDayOfTaxYear").value
      ben.value(ln_CheapTaxableLoan_db) = IsNullEx(rs.Fields("CheapTaxableLoan"), True)
    
    End If
    
    If ben.value(ln_nborrowers_db) = 0 Then ben.value(ln_nborrowers_db) = 1
    
    Call AddNOther
    
    Set db = m_Parent.Parent.Parent.db
    Call ReadBalances(db)
    
    m_ReadFromDB = True
  End If
  
  IBenefitClass_ReadDB = True
  
clsLoan_ReadDB_end:
  Set rs = Nothing
  Set db = Nothing
  
  Call xReturn("clsLoan_ReadDB")
  Exit Function
clsLoan_ReadDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Loan_readDB", "Loan ReadDC", "Unable to read the loan details.")
  Resume clsLoan_ReadDB_end
  Resume
End Function


Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
  bc = BC_LOAN_OTHER_H
    
    If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
    m_BenClass = bc
    Set ben = Me
    p11d32.BenDataLinkBenfitTable(bc) = ben.TABLE
    m_NeedToCalculate = True
    With p11d32
    
      Call SetStandardBenItemsDataTypes(bc)
      Call SetStandardBenItemsUDMData(bc)
      Call SetStandardBenItemsMMFieldSize(bc)
      
      .BenDataLinkDataType(bc, ln_amountwaived_db) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_balkey_db) = TYPE_LONG
      
      .BenDataLinkDataType(bc, ln_amountwaived_db) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_balkey_db) = TYPE_LONG
      
      .BenDataLinkDataType(bc, ln_DeemedInterestNormal) = TYPE_DOUBLE
      .BenDataLinkDataType(bc, ln_DeemedInterestAlternative) = TYPE_DOUBLE
      .BenDataLinkDataType(bc, ln_BalancesAmmended) = TYPE_BOOL
      .BenDataLinkDataType(bc, ln_ZeroBalanceAtEndOfYear) = TYPE_BOOL
      .BenDataLinkDataType(bc, ln_OpenOutstanding) = TYPE_LONG
      .BenDataLinkMMFieldSize(bc, ln_OpenOutstanding) = 9
      .BenDataLinkDataType(bc, ln_CloseOutstanding) = TYPE_LONG
      .BenDataLinkMMFieldSize(bc, ln_CloseOutstanding) = 9
      .BenDataLinkDataType(bc, ln_MaxOutstandingAtAnyPoint) = TYPE_LONG
      .BenDataLinkMMFieldSize(bc, ln_MaxOutstandingAtAnyPoint) = 9
      
      .BenDataLinkDataType(bc, ln_balkey_db) = TYPE_STR
      
      
      .BenDataLinkDataType(bc, ln_DateDischargedOrMade) = TYPE_DATE
      .BenDataLinkMMFieldSize(bc, ln_DateDischargedOrMade) = MM_SFS_DATE
            
      .BenDataLinkDataType(bc, ln_DidLoanCommenceOnFirstDayOfTaxYear_db) = TYPE_BOOL
       
      .BenDataLinkDataType(bc, ln_AverageOutstanding) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_AverageLoanRate) = TYPE_DOUBLE
      .BenDataLinkDataType(bc, ln_WholeTaxMonthsOutstanding) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_CashEquivalentNormal) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_CashEquivalentDaily) = TYPE_LONG
      
      .BenDataLinkMMFieldSize(bc, ln_InterestPaid_db) = 9
      
      
      .BenDataLinkDataType(bc, ln_InterestNormal) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_InterestDaily) = TYPE_DOUBLE
      
      .BenDataLinkDataType(bc, ln_DischargedDate) = TYPE_DATE
      .BenDataLinkMMFieldSize(bc, ln_DischargedDate) = MM_SFS_DATE
      .BenDataLinkDataType(bc, ln_MadeDate) = TYPE_DATE
      .BenDataLinkMMFieldSize(bc, ln_MadeDate) = MM_SFS_DATE
      
      
      .BenDataLinkDataType(bc, ln_nborrowers_db) = TYPE_LONG
      .BenDataLinkMMFieldSize(bc, ln_nborrowers_db) = 3
      
      .BenDataLinkDataType(bc, ln_DefaultMirasRelief) = TYPE_LONG
      
      .BenDataLinkDataType(bc, ln_Deemed) = TYPE_DOUBLE             'MP DB - ToDo
      
      .BenDataLinkDataType(bc, ln_LoanType_db) = TYPE_STR
      .BenDataLinkMMFieldSize(bc, ln_LoanType_db) = 1
      
      .BenDataLinkDataType(bc, ln_LoanCurrency_db) = TYPE_STR
      .BenDataLinkDataType(bc, ln_LoanCurrencyIndex) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_Lender) = TYPE_STR
      .BenDataLinkDataType(bc, ln_ACNumber) = TYPE_STR
      .BenDataLinkDataType(bc, ln_opendate) = TYPE_DATE
      .BenDataLinkDataType(bc, ln_closedate) = TYPE_DATE
      .BenDataLinkDataType(bc, ln_UseDailyCalculationOnly_db) = TYPE_BOOL
      .BenDataLinkDataType(bc, ln_CheapTaxableLoan_db) = TYPE_BOOL
      .BenDataLinkDataType(bc, ln_HMIT_NoOfBorrowers) = TYPE_STR
      
      .BenDataLinkDataType(bc, ln_WAIVEDavailablefrom) = TYPE_DATE
      .BenDataLinkDataType(bc, ln_WAIVEDavailableto) = TYPE_DATE
      .BenDataLinkDataType(bc, ln_WAIVEDBenefit) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDBenefit_Reportable) = TYPE_BOOL
      .BenDataLinkDataType(bc, ln_WAIVEDcategory) = TYPE_STR                         'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDclass) = TYPE_STR                            'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDCompanyDefinedCategory) = TYPE_STR           'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDCompanyDefinedCategoryKey) = TYPE_LONG       'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDEmployeeReference) = TYPE_STR                'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDGrossAmountPaidByEmployer) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDHMITCode_UDBCode) = TYPE_STR                 'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDitem) = TYPE_STR
      .BenDataLinkDataType(bc, ln_WAIVEDMadeGood) = TYPE_LONG                        'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDMadeGood_NET) = TYPE_LONG                    'MP DB - ToDo
      .BenDataLinkDataType(bc, ln_WAIVEDValue) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDBoxNumber) = TYPE_STR
      .BenDataLinkDataType(bc, ln_WAIVEDUDM_BENEFIT_TITLE) = TYPE_STR
      
      'cad loan fix
      .BenDataLinkDataType(bc, ln_WAIVEDNICClass1AAble) = TYPE_BOOL
      .BenDataLinkDataType(bc, ln_WAIVEDNIC_Class1A_Value) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDNIC_Class1A_Benefit) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDClass1AAdjustment) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVED_IRDesc) = TYPE_STR
      .BenDataLinkDataType(bc, ln_WAIVEDActualAmountMadeGood) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDOPRA_Ammount_Foregone_db) = TYPE_LONG
      .BenDataLinkDataType(bc, ln_WAIVEDOPRA_Ammount_Foregone_Used_For_Value) = TYPE_BOOL
        
                  
            
      .BenDataLinkUDMDescription(bc, ln_OpenOutstanding) = "Opening balance"
      .BenDataLinkUDMDescription(bc, ln_CloseOutstanding) = "Closing balance"
      .BenDataLinkUDMDescription(bc, ln_InterestPaid_db) = "Interest paid"
      .BenDataLinkUDMDescription(bc, ln_amountwaived_db) = "Waived"
      .BenDataLinkUDMDescription(bc, ln_CheapTaxableLoan_db) = S_FIELD_LOAN_CHEAP_TAXABLE
      
      .DataLinkInitialised(bc) = True
    End With
  
SetBenItemsInformation_end:
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub


Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  Select Case m_BenClass
    Case BC_LOAN_OTHER_H
      IBenefitClass_TABLE = TBL_BENLOAN
    Case Else
      Call ECASE("Invalid ben class in table for loan")
  End Select
End Property

Public Property Get BenefitClass() As BEN_CLASS
  BenefitClass = m_BenClass
End Property
Public Property Let BenefitClass(ben As BEN_CLASS)
  
End Property


Public Property Let InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Public Property Get InvalidFields() As Long
  InvalidFields = m_InvalidFields
End Property


Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
    If (Item = ITEM_BENEFIT_REPORTABLE) Then
        Item = Item
    End If
  m_BenItems(Item) = CorrectBenValue(m_BenClass, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim Loan As IBenefitClass
  Dim db As Database
  Dim ben As IBenefitClass
  
On Error GoTo clsLoan_WriteDB_Err
  
  Call xSet("clsLoan_WriteDB")
  
  Set ben = Me

  If p11d32.LoanHomeToBeneficialBringFwd Then m_BenClass = BC_LOAN_OTHER_H 'km

  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields("CURR").value = ben.value(ln_LoanCurrency_db)
  rs.Fields("ITEM").value = ben.value(ln_item_db)
  
  If Not p11d32.BringForward.Yes Then
    rs.Fields("WAIVED").value = ben.value(ln_amountwaived_db)
    rs.Fields("Interest").value = ben.value(ln_InterestPaid_db)
    rs.Fields("DailyOnly").value = ben.value(ln_UseDailyCalculationOnly_db)
  End If
  
  rs.Fields("LTYPE").value = ben.value(ln_LoanType_db)
  rs.Fields("BORROWERS").value = ben.value(ln_nborrowers_db)
  
  rs.Fields("CheapTaxableLoan") = ben.value(ln_CheapTaxableLoan_db)
  
  Call StandardWriteData(ben, rs)
    If Not p11d32.BringForward.Yes Then
      rs.Fields("DidLoanCommenceOnFirstDayOfTaxYear").value = ben.value(ln_DidLoanCommenceOnFirstDayOfTaxYear_db)
    Else
      If (p11d32.AppYear = 2004) Then
        rs.Fields("CheapTaxableLoan") = True
      Else
        rs.Fields("CheapTaxableLoan") = ben.value(ln_CheapTaxableLoan_db)
      End If
      rs.Fields("DidLoanCommenceOnFirstDayOfTaxYear").value = False
    End If
  
  Call BenefitCloseRecord(ben, rs)
  
  If IsNull(rs.Fields("BalKey").value) Or p11d32.BringForward.Yes Then
    rs.Edit
    rs.Fields("BalKey").value = GenerateGUID
    rs.Update
  End If
  
  ben.value(ln_balkey_db) = rs.Fields("BalKey").value
  
  Set db = m_Parent.Parent.Parent.db
  
  Call WriteBalances(db)
  
  IBenefitClass_WriteDB = True
  
clsLoan_WriteDB_end:
  Set rs = Nothing
  
  Set db = Nothing
  xReturn "clsLoan_WriteDB"
  Exit Function
clsLoan_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "Loan_WriteDb", "Write DB", "Error writing the loan benefit.")
  Resume clsLoan_WriteDB_end
  Resume
End Function

Public Property Get LoanBalanceSheet(Index As Long) As Long
  'Set LoanBalanceSheet = m_LoanBalanceSheet(Index)
End Property

Public Property Get Qualifying() As Boolean
  Qualifying = False
  'New rules for 2000/2001, only non-qualifying loans are reportable
'  Qualifying = (Not CBool(StrComp(S_LOANASHORT, m_BenItems(ln_LoanType))) Or Not CBool(StrComp(S_LOANBSHORT, m_BenItems(ln_LoanType))) Or Not CBool(StrComp(S_LOANCSHORT, m_BenItems(ln_LoanType))))
 End Property
Private Sub WriteBalanceItem(rs As Recordset, bi As BalanceItem, sLoanType As String, ByVal sPNum As String)
  On Error GoTo WriteBalanceItem_ERR
  

  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WriteBalanceItem", "Write Balance Item", "The recordset is nothing")
  If bi Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WriteBalanceItem", "Write Balance Item", "The Balance item is nothing.")
  
  rs.AddNew
  rs.Fields("KEY").value = m_BenItems(ln_balkey_db)
  rs.Fields("TYPE").value = sLoanType
  rs.Fields("FROM").value = bi.DateFrom
  rs.Fields("PAYMENT").value = bi.Payment
  rs.Fields("P_Num").value = sPNum
  rs.Fields("Item").value = m_BenItems(ln_item_db)
  rs.Update
    
WriteBalanceItem_END:
  Exit Sub
WriteBalanceItem_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteBalanceItem", "Write Balance Item", "Error writing balance item for a loan.")
  Resume WriteBalanceItem_END
  Resume
End Sub
Public Function WriteBalances(dbEmployer As Database) As Boolean
  Dim l As Long
  Dim rs As Recordset
  Dim bi As BalanceItem
  Dim sLoanType As String, sPNum As String
  Dim ben As IBenefitClass, benEmployee As IBenefitClass
    
  On Error GoTo WriteBalances_Err
  
  Call xSet("WriteBalances")
  
  sLoanType = LoanTypeString
  
  Set ben = Me
  Set benEmployee = m_Parent.Parent
  sPNum = benEmployee.value(ee_PersonnelNumber_db)
  
  If ben.value(ln_balkey_db) = Empty Then Call Err.Raise(ERR_IS_NOTHING, "WriteBalances", "Write Balances", "The Balance key has not been initialised.")
  
  If ben.value(ln_BalancesAmmended) Then Call dbEmployer.Execute(sql.Queries(DELETE_LOANBAL, ben.value(ln_balkey_db), sLoanType))
  
  Set rs = dbEmployer.OpenRecordset(sql.Queries(SELECT_ALLLOANBALS), dbOpenDynaset, dbFailOnError)
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WriteBalances", "Write Balances", "The Balance key has not been initialised.")
  
  If p11d32.BringForward.Yes Then
    Set bi = New BalanceItem
    bi.DateFrom = p11d32.Rates.value(TaxYearStart)
    bi.Payment = CloseOutStanding(False)
    Call WriteBalanceItem(rs, bi, sLoanType, sPNum)
    GoTo WriteBalances_End
  Else
    If ben.value(ln_BalancesAmmended) Then
      For l = 1 To BalanceSheet.Count
        Set bi = BalanceSheet(l)
        If Not bi Is Nothing Then Call WriteBalanceItem(rs, bi, sLoanType, sPNum)
      Next
      ben.value(ln_BalancesAmmended) = False
    End If
  End If

WriteBalances_End:
  Call xReturn("WriteBalances")
  Exit Function

WriteBalances_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteBalances", "Write Balances", "Error writing the loan balances to the database.")
  Resume WriteBalances_End
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property


Private Function ReadBalances(dbEmployer As Database) As Boolean
  Dim rs As Recordset
  Dim BalanceItem As BalanceItem
  Dim ben As IBenefitClass
  
On Error GoTo ReadBalances_Err
  
  Call xSet("ReadBalances")

  Set ben = Me

  Set rs = dbEmployer.OpenRecordset(sql.Queries(SELECT_LOAN_BALANCES, ben.value(ln_balkey_db), LoanTypeString), dbOpenForwardOnly, dbFailOnError)
  BalanceSheet.RemoveAll
  
  Do While Not rs.EOF
    Set BalanceItem = New BalanceItem
    With BalanceItem
      .DateFrom = IIf(IsNull(rs.Fields("From").value), UNDATED, rs.Fields("From").value)
      .Payment = IIf(IsNull(rs.Fields("Payment").value), 0, rs.Fields("Payment").value)
    End With
    Call BalanceSheet.Add(BalanceItem)
    rs.MoveNext
  Loop

  Call FinishBalances(True)

ReadBalances_End:
  Set rs = Nothing
  Set BalanceItem = Nothing
  Call xReturn("ReadBalances")
  Exit Function

ReadBalances_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadBalances", "Read Balances", "Error reading the loan balances.")
  Resume ReadBalances_End
  Resume
End Function


Public Function FinishBalances(bOrderBalancesByDate As Boolean) As Boolean
  Dim bi As BalanceItem, BINext As BalanceItem
  Dim dblLoanBalance As Double
  Dim dblWorkingBalance As Double
  Dim BICurrentIndex As Long
  Dim dLeapDate As Date
  Dim bLeapYear As Boolean
  Dim bCrossLeapDay As Boolean
  
  
On Error GoTo FinishBalances_Err
  
  Call xSet("FinishBalances")
  
  dblLoanBalance = LoanBalance
  
  If bOrderBalancesByDate Then Call BalanceSheet.Sort(Me)
  bLeapYear = p11d32.Rates.value(IsLeapyYear)
  
  If bLeapYear Then dLeapDate = p11d32.Rates.value(LeapDate)
  
  If BalanceSheet.Count Then
    BICurrentIndex = GetFirstBalanceItemFromDateFrom(bi)
    If BICurrentIndex > 0 Then
      Do While Not bi Is Nothing
        dblWorkingBalance = dblWorkingBalance + bi.Payment
        If GetNextBalanceItemFromDateFrom(BINext, bi, BICurrentIndex) Then
          If BINext Is Nothing Then
           'last one so check the balance
            If dblLoanBalance = 0 Then
              bi.DateTo = bi.DateFrom
            Else
              bi.DateTo = p11d32.Rates.value(TaxYearEnd)
            End If
          Else
            bi.DateTo = DateAdd("d", -1, BINext.DateFrom)
          End If
        Else
          Set BINext = Nothing
        End If
        If bLeapYear Then bCrossLeapDay = (bi.DateFrom <= dLeapDate) And (bi.DateTo >= dLeapDate)
        bi.days = DateDiff("d", bi.DateFrom, bi.DateTo) + Abs(Not (bCrossLeapDay))
        bi.Balance = RoundN(dblWorkingBalance)
        Set bi = BINext
        bCrossLeapDay = False
      Loop
    End If
    Call CalculateInterest
  End If

FinishBalances_End:
  Call xReturn("FinishBalances")
  Exit Function
FinishBalances_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FinishBalances", "Finish Balances", "Error finishing the balances information.")
  Resume FinishBalances_End
  Resume
End Function

  

Public Function LoanBalance() As Double
  Dim dbl As Double
  Dim l As Long
  Dim bi As BalanceItem
  
  On Error GoTo LoanBalance_Err
  Call xSet("LoanBalance")

  For l = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(l)
    If Not bi Is Nothing Then
      dbl = dbl + bi.Payment
    End If
  Next

  LoanBalance = RoundN(dbl, 2)
  
LoanBalance_End:
  Call xReturn("LoanBalance")
  Exit Function
LoanBalance_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoanBalance", "Loan Balance", "Error obtaining the overall loan balance.")
  Resume LoanBalance_End
End Function


Public Function GetNextBalanceItemFromDateFrom(BIToFind As BalanceItem, BICurrent As BalanceItem, BICurrentIndex As Long) As Boolean
  Dim l As Long
  Dim bi As BalanceItem
  Dim vCurrentDate As Date
  
  On Error GoTo GetNextBalanceItemFromDateFrom_Err
  Call xSet("GetNextBalanceItemFromDateFrom")
  
  vCurrentDate = p11d32.Rates.value(NextTaxYearStart)
  
  Set BIToFind = Nothing
  
  For l = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(l)
    If Not bi Is Nothing Then
      If (l > BICurrentIndex) Then
        If (CVDate(bi.DateFrom) >= CVDate(BICurrent.DateFrom)) And (CVDate(bi.DateFrom) < vCurrentDate) Then
          BICurrentIndex = l
          Set BIToFind = bi
          vCurrentDate = BIToFind.DateFrom
        End If
      End If
    End If
  Next

  GetNextBalanceItemFromDateFrom = True

GetNextBalanceItemFromDateFrom_End:
  Call xReturn("GetNextBalanceItemFromDateFrom")
  Exit Function

GetNextBalanceItemFromDateFrom_Err:
  GetNextBalanceItemFromDateFrom = False
  Call ErrorMessage(ERR_ERROR, Err, "GetNextBalanceItemFromDateFrom", "Get Next Balance Item From Date From", "Error obtaining the next balance item by date.")
  Resume GetNextBalanceItemFromDateFrom_End
End Function


Public Function GetFirstBalanceItemFromDateFrom(bi As BalanceItem) As Long
  Dim l As Long
  Dim BICurrent As BalanceItem
  Dim vDateFrom As Date
  
  On Error GoTo GetFirstBalanceItemFromDateFrom_Err
  Call xSet("GetFirstBalanceItemFromDateFrom")
  
  vDateFrom = p11d32.Rates.value(NextTaxYearStart)

  For l = 1 To BalanceSheet.Count
    Set BICurrent = BalanceSheet(l)
    If Not BICurrent Is Nothing Then
      
      If CVDate(BICurrent.DateFrom) < vDateFrom Then
        Set bi = BICurrent
        vDateFrom = bi.DateFrom
        GetFirstBalanceItemFromDateFrom = l
      End If
    End If
  Next

GetFirstBalanceItemFromDateFrom_End:
  Set BICurrent = Nothing
  Call xReturn("GetFirstBalanceItemFromDateFrom")
  Exit Function
GetFirstBalanceItemFromDateFrom_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetFirstBalanceItemFromDateFrom", "Get First Balance Item From Date From", "Error obtaining the first balance item by date.")
  Resume GetFirstBalanceItemFromDateFrom_End
End Function

Public Property Get MaxOutstandingAtAnyPointInYear() As Variant
  
  Dim dDate As Date
  Dim dblMax As Long, dblNext As Long
  Dim lLastIndex As Long, lLastIndexAction As Long
On Error GoTo MaxOutstandingAtAnyPointInYear_Err

  Call xSet("MaxOutstandingAtAnyPointInYear")
  
  'go through the balances and at any event date find the max outstanding
  MaxOutstandingAtAnyPointInYear = S_ERROR
  dDate = p11d32.Rates.value(LastTaxYearEnd)
  
  Do
    dblNext = RoundDownEx(BalanceOnDate(dDate, lLastIndex), 0)
    If dblNext > dblMax Then
      dblMax = dblNext
      'Exit Do
    End If
    lLastIndexAction = lLastIndexAction + 1
  Loop Until NextActionDate(dDate, dDate, lLastIndexAction) = False

  MaxOutstandingAtAnyPointInYear = dblMax

MaxOutstandingAtAnyPointInYear_End:
  Call xReturn("MaxOutstandingAtAnyPointInYear")
  Exit Property

MaxOutstandingAtAnyPointInYear_Err:
  MaxOutstandingAtAnyPointInYear = S_ERROR
  Call ErrorMessage(ERR_ERROR, Err, "MaxOutstandingAtAnyPointInYear", "Max Outstanding At Any Point In Year", "Error calculating the maximum balance outstanding at any point in the year.")
  Resume MaxOutstandingAtAnyPointInYear_End
End Property


Public Function OrderBalanceItemsByDateFrom() As Boolean
  Dim vDateFrom As Date, vDateMax As Date
  Dim BalanceSheetTemp As ObjectList
  Dim bi As BalanceItem
  Dim lObjectListIndex As Long
  Dim l As Long, m As Long
  Dim v As Variant
  
  On Error GoTo OrderBalanceItemsByDateFrom_Err
  Call xSet("OrderBalanceItemsByDateFrom")

  
  GoTo OrderBalanceItemsByDateFrom_End
  
  Set BalanceSheetTemp = New ObjectList
  
  If BalanceSheet.Count = 0 Then GoTo OrderBalanceItemsByDateFrom_End
  
  vDateMax = p11d32.Rates.value(NextTaxYearStart)
  vDateFrom = vDateMax
  
  
  For m = 1 To BalanceSheet.Count
    For l = 1 To BalanceSheet.Count
      Set bi = BalanceSheet(l)
      If Not bi Is Nothing Then
        If CVDate(bi.DateFrom) < vDateFrom Then
          vDateFrom = bi.DateFrom
          lObjectListIndex = l
        End If
      End If
    Next
    
    If lObjectListIndex <> -1 Then
      Set bi = BalanceSheet(lObjectListIndex)
      Call BalanceSheetTemp.Add(bi)
      Call BalanceSheet.Remove(lObjectListIndex)
      vDateFrom = vDateMax
      lObjectListIndex = -1
    Else
      Exit For
    End If
  Next
  
  BalanceSheet.RemoveAll
  
  For l = 1 To BalanceSheetTemp.Count
    Set bi = BalanceSheetTemp(l)
    Call BalanceSheet.Add(bi)
  Next
  
OrderBalanceItemsByDateFrom_End:
  Set bi = Nothing
  'BalanceSheetTemp.RemoveAll
  'Set BalanceSheetTemp = Nothing
  Call xReturn("OrderBalanceItemsByDateFrom")
  Exit Function

OrderBalanceItemsByDateFrom_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OrderBalanceItemsByDateFrom", "Order Balance Items By Date From", "Error ordering the balance items by date from.")
  Resume OrderBalanceItemsByDateFrom_End
  Resume
End Function


Public Function NextActionDate(dNextDate As Date, ByVal dStartDate As Date, lCurrentBalaceIndex As Long) As Boolean
  Dim l As Long
  Dim bi As BalanceItem
  Dim dMaxDate As Date
  
  On Error GoTo NextActionDate_Err
  Call xSet("NextActionDate")
  
  dNextDate = p11d32.Rates.value(NextTaxYearStart)
  For l = lCurrentBalaceIndex To BalanceSheet.Count
    Set bi = BalanceSheet(l)
    If Not bi Is Nothing Then
      If bi.DateFrom < dNextDate And bi.DateFrom > dStartDate Then
        dNextDate = bi.DateFrom
        lCurrentBalaceIndex = l
      End If
    End If
  Next

  If dNextDate <> p11d32.Rates.value(NextTaxYearStart) Then
    NextActionDate = True
  End If
  
NextActionDate_End:
  Call xReturn("NextActionDate")
  Exit Function

NextActionDate_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NextActionDate", "Next Action Date", "Error obtaining the next action date.")
  Resume NextActionDate_End
  Resume
End Function
Public Function SubCalculate() As Variant
  Dim i As Long
  Dim bi As BalanceItem
  Dim loans As loans
  Dim ben As IBenefitClass
  
  On Error GoTo SubCalculate_Err
  Call xSet("SubCalculate")

  Set ben = Me
  
  Call ben.SetCalcDefaults
    
  SubCalculate = 0

  If NumberOfBalances > 0 Or ben.value(ln_OPRA_Ammount_Foregone_db) > 0 Then
    SubCalculate = S_ERROR
    ben.value(ln_opendate) = S_ERROR
    ben.value(ln_closedate) = S_ERROR
    ben.value(ln_CashEquivalentNormal) = S_ERROR
    ben.value(ln_CashEquivalentDaily) = S_ERROR
    
    Call DateLoanDischargedOrMade
    
    ben.value(ln_MaxOutstandingAtAnyPoint) = MaxOutstandingAtAnyPointInYear
    
    Call CalculateInterest
    'CAD2004R1, also in to Loan Calculate->IsDeminimus, as it is required
    'could have not set the calc defaults to be reset for the interest values
    
    Set loans = m_Parent
    
    If loans.AllLoansDeminimus Then
      ben.value(ln_Benefit) = 0
      ben.value(ln_Benefit_Reportable) = False
    Else
      ben.value(ln_Benefit_Reportable) = True
      If ben.value(ln_UseDailyCalculationOnly_db) Then
        ben.value(ln_Value_Non_OPRA) = ben.value(ln_InterestDaily)
        
        Call CalculateOpRAValue(ben)
        If (ben.value(ln_OPRA_Ammount_Foregone_Used_For_Value)) Then
          ben.value(ln_Benefit) = ben.value(ln_OPRA_Ammount_Foregone_db) - ben.value(ln_InterestPaid_db)
          ben.value(ln_MadeGood_Net) = ben.value(ln_InterestPaid_db) + ben.value(ln_Benefit)
          ben.value(ln_Benefit) = Max(0, ben.value(ln_Benefit))
        Else
          ben.value(ln_Benefit) = ben.value(ln_CashEquivalentDaily)
          ben.value(ln_MadeGood_Net) = Min(ben.value(ln_InterestDaily), ben.value(ln_InterestPaid_db))
        End If
        
      Else
        ben.value(ln_Value_Non_OPRA) = ben.value(ln_InterestNormal)
        Call CalculateOpRAValue(ben)
        
        If (ben.value(ln_OPRA_Ammount_Foregone_Used_For_Value)) Then
          ben.value(ln_Benefit) = ben.value(ln_OPRA_Ammount_Foregone_db) - ben.value(ln_InterestPaid_db)
          ben.value(ln_MadeGood_Net) = ben.value(ln_InterestPaid_db) + ben.value(ln_Benefit)
          ben.value(ln_Benefit) = Max(0, ben.value(ln_Benefit))
        Else
          ben.value(ln_Benefit) = ben.value(ln_CashEquivalentNormal)
          ben.value(ln_MadeGood_Net) = Min(ben.value(ln_InterestNormal), ben.value(ln_InterestPaid_db))
        End If
      End If
      'cad removed oh shit.... ! 27/03/2002 .... ben.value(ln_MadeGood_Net) = ben.value(ln_InterestPaid_db)
    End If
    SubCalculate = ben.value(ln_Benefit)
  Else
    ben.value(ln_Value) = 0
    ben.value(ln_Benefit) = 0
  End If
    
  ben.value(ln_WAIVEDitem) = ben.value(ITEM_DESC)
  
  ben.value(ln_WAIVEDGrossAmountPaidByEmployer) = ben.value(ln_amountwaived_db)
  ben.value(ln_WAIVEDValue) = ben.value(ln_WAIVEDGrossAmountPaidByEmployer)
  ben.value(ln_WAIVEDBenefit) = ben.value(ln_WAIVEDValue)
  ben.value(ln_WAIVEDavailablefrom) = IIf(IsDate(ben.value(ln_opendate)), ben.value(ln_opendate), p11d32.Rates.value(TaxYearStart)) 'ek added 1/04, fix 170
  ben.value(ln_WAIVEDavailableto) = IIf(IsDate(ben.value(ln_closedate)), ben.value(ln_closedate), p11d32.Rates.value(TaxYearEnd))  'ek added 1/04, fix 170
  
  Call BenCalcNIC(ben)
  
  If ben.value(ln_WAIVEDValue) > 0 Then ben.value(ln_WAIVEDBenefit_Reportable) = True
    

SubCalculate_End:
  Set bi = Nothing
  Set loans = Nothing
  Call xReturn("SubCalculate")
  Exit Function

SubCalculate_Err:
  SubCalculate = S_ERROR
  Resume SubCalculate_End
  Resume
End Function
Private Sub DateLoanMade(ByVal dblLoanBalance As Double)
  Dim i As Long
  Dim bi As BalanceItem
  Dim ben As IBenefitClass
  
  On Error GoTo DateLoanMade_Err
  
  Call xSet("DateLoanMade")
  
  Set ben = Me
  
  ben.value(ln_MadeDate) = UNDATED
  
  For i = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(i)
    If Not bi Is Nothing Then
      If bi.Balance > 0 Then
        If (bi.DateFrom = p11d32.Rates.value(TaxYearStart)) Then
          If ben.value(ln_DidLoanCommenceOnFirstDayOfTaxYear_db) Then
            
            ben.value(ln_MadeDate) = bi.DateFrom
            GoTo DateLoanMade_End
          Else
            GoTo DateLoanMade_End
          End If
        Else
          ben.value(ln_MadeDate) = bi.DateFrom
          GoTo DateLoanMade_End
        End If
      Else
        GoTo DateLoanMade_End
      End If
    End If
  Next

DateLoanMade_End:
  Call xReturn("DateLoanMade")
  Exit Sub
DateLoanMade_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DateLoanMade", "Date Loan Made", "Error finding the date when the loan was made.")
  Resume DateLoanMade_End
End Sub
Private Sub DateLoanDischargedOrMade()
  Dim ben As IBenefitClass
  Dim dblLoanBalance As Double
  
  On Error GoTo DateLoanDischargedOrMade_Err
  
  Set ben = Me
  
  ben.value(ln_DateDischargedOrMade) = UNDATED
  dblLoanBalance = LoanBalance
  
  Call DateLoanDischanged(dblLoanBalance)
  Call DateLoanMade(dblLoanBalance)
  If ben.value(ln_DischargedDate) <> UNDATED Then
    ben.value(ln_DateDischargedOrMade) = ben.value(ln_DischargedDate)
  ElseIf ben.value(ln_MadeDate) <> UNDATED Then
    ben.value(ln_DateDischargedOrMade) = ben.value(ln_MadeDate)
  End If
  
  
DateLoanDischargedOrMade_End:
  Exit Sub
DateLoanDischargedOrMade_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DateLoanDischargedOrMade", "Date Loan Discharged Or Made", "Error setting the date loan discharged or made.")
  Resume DateLoanDischargedOrMade_End
End Sub
Private Sub DateLoanDischanged(ByVal dblLoanBalance As Double)
  Dim i As Long
  Dim bi As BalanceItem
  Dim ben As IBenefitClass
  
  On Error GoTo DateLoanDischanged_Err
  
  Call xSet("DateLoanDischanged")
  
  Set ben = Me
  
  
  ben.value(ln_DischargedDate) = UNDATED
  
  If dblLoanBalance <= 0 Then 'we have a finished loan
    'get the last balance date
    For i = BalanceSheet.Count To 1 Step -1
      Set bi = BalanceSheet(i)
      If Not bi Is Nothing Then
        'if at start then assume UNDATED CAD , rolling on to next year
        
        ben.value(ln_DischargedDate) = bi.DateFrom
        GoTo DateLoanDischanged_End
      End If
    Next
  End If


DateLoanDischanged_End:
  Call xReturn("DateLoanDischanged")
  Exit Sub
DateLoanDischanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DateLoanDischanged", "Date Loan Dischanged", "Error finding the date when the loan was discharged.")
  Resume DateLoanDischanged_End
End Sub


Public Property Get OpenOutStanding() As Variant
  Dim l As Long
  Dim bi As BalanceItem
  Dim d As Date
  
  On Error GoTo OpenOutStanding_Err
  Call xSet("OpenOutStanding")

  OpenOutStanding = S_ERROR
  d = UNDATED
  For l = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(l)
    If Not bi Is Nothing Then
      If bi.Balance > 0 Then
        If (d = UNDATED) Then
          d = bi.DateFrom
                    
          m_BenItems(ln_opendate) = bi.DateFrom 'chec with AJ that sould be 0 dp for both me and close outstading
        Else
          If bi.DateFrom <> d Then
             Exit For
          End If
        End If
        OpenOutStanding = bi.Balance

      End If
    End If
  Next

  
  If Not IsNumeric(OpenOutStanding) Then
    m_BenItems(ln_opendate) = p11d32.Rates.value(TaxYearStart)
    OpenOutStanding = 0
  Else
    OpenOutStanding = RoundDownEx(OpenOutStanding, 0)  'AM changed from roundN to RoundDown
  End If

OpenOutStanding_End:
  Call xReturn("OpenOutStanding")
  Exit Property

OpenOutStanding_Err:
  OpenOutStanding = S_ERROR
  Call ErrorMessage(ERR_ERROR, Err, "OpenOutStanding", "Open Out Standing", "Error finding the opening outstanding loan balance.")
  Resume OpenOutStanding_End
End Property


Public Function CloseOutStanding(Optional ByVal bRounding As Boolean = True) As Variant
  Dim i As Long, lCount As Long
  Dim bi As BalanceItem
  Dim ben As IBenefitClass
  
  On Error GoTo CloseOutStanding_Err
  Call xSet("CloseOutStanding")

  CloseOutStanding = S_ERROR
  Set ben = Me
  
  lCount = BalanceSheet.Count
  For i = lCount To 1 Step -1
    Set bi = BalanceSheet(i)
    If Not bi Is Nothing Then
      If i = lCount Then If bi.Balance <= 0 Then ben.value(ln_ZeroBalanceAtEndOfYear) = True
      If bi.Balance > 0 Then
        CloseOutStanding = bi.Balance
        ben.value(ln_closedate) = bi.DateTo
        Exit For
      End If
    End If
  Next
  
  If IsNumeric(CloseOutStanding) Then
    If CloseOutStanding <= 0 Then
      CloseOutStanding = 0
    Else
      If bRounding Then
        CloseOutStanding = RoundDownEx(CloseOutStanding, 0) 'AM Changed from RoundN to RoundDown
      End If
    End If
  Else
    'ie no balances
    ben.value(ln_ZeroBalanceAtEndOfYear) = True
    ben.value(ln_closedate) = p11d32.Rates.value(TaxYearStart)
    CloseOutStanding = 0
  End If
  
CloseOutStanding_End:
  Call xReturn("CloseOutStanding")
  Exit Function

CloseOutStanding_Err:
  CloseOutStanding = S_ERROR
  Call ErrorMessage(ERR_ERROR, Err, "CloseOutStanding", "Close Out Standing", "Error finding the closing outstanding loan balance.")
  Resume CloseOutStanding_End
End Function


Public Function WholeTaxMonthsOutstanding(ByVal dDateStart As Date, ByVal dDateEnd As Date) As Variant
  Dim dTaxMonthEnd As Date, dTaxMonth As Date
  Dim l As Long
  
On Error GoTo WholeTaxMonthsOutstanding_Err
  
  Call xSet("WholeTaxMonthsOutstanding")

  WholeTaxMonthsOutstanding = S_ERROR
  
  dTaxMonth = p11d32.Rates.value(TaxYearStart)
  dTaxMonthEnd = DateAdd("d", -1, DateAdd("m", 1, dTaxMonth))
  Do While dDateEnd >= dTaxMonthEnd
    If (dDateStart <= dTaxMonth) And (dDateEnd >= dTaxMonthEnd) Then
      l = l + 1
    End If
    If dDateEnd < dTaxMonthEnd Then
      Exit Do
    End If
    dTaxMonth = DateAdd("m", 1, dTaxMonth)
    dTaxMonthEnd = DateAdd("d", -1, DateAdd("m", 1, dTaxMonth))
  Loop
  
  WholeTaxMonthsOutstanding = l

WholeTaxMonthsOutstanding_End:
  Call xReturn("WholeTaxMonthsOutstanding")
  Exit Function

WholeTaxMonthsOutstanding_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WholeTaxMonthsOutstanding", "Whole Tax Months Outstanding", "Error finding the number of complete tax months the loan was outstanding.")
  Resume WholeTaxMonthsOutstanding_End
End Function


Public Property Get NumberOfBalances() As Long
  
  On Error GoTo NumberOfBalances_Err
  Call xSet("NumberOfBalances")

  
  NumberOfBalances = BalanceSheet.Count

NumberOfBalances_End:
  Call xReturn("NumberOfBalances")
  Exit Property

NumberOfBalances_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NumberOfBalances", "Number Of Balances", "Error findinh the number of balances for the loan.")
  Resume NumberOfBalances_End
End Property


Private Function SetLoanCurrenyIndex() As Boolean

  On Error GoTo SetLoanCurrenyIndex_Err
  Call xSet("SetLoanCurrenyIndex")

  Select Case UCASE$(m_BenItems(ln_LoanCurrency_db))
    Case UCASE$(S_LOANFRANC)
      m_BenItems(ln_LoanCurrencyIndex) = 1
    Case UCASE$(S_LOANYEN)
      m_BenItems(ln_LoanCurrencyIndex) = 2
    Case Else
      m_BenItems(ln_LoanCurrencyIndex) = 0
  End Select

SetLoanCurrenyIndex_End:
  Call xReturn("SetLoanCurrenyIndex")
  Exit Function

SetLoanCurrenyIndex_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetLoanCurrenyIndex", "Set Loan Curreny Index", "Error setting the loan currency index.")
  Resume SetLoanCurrenyIndex_End
End Function


Public Function CalculateInterest() As Boolean
  Dim lDays As Long
  Dim i As Long
  Dim bi As BalanceItem
  Dim ben As IBenefitClass
  
  On Error GoTo CalculateInterest_Err
  Call xSet("CalculateInterest")

  Set ben = Me
  ben.value(ln_AverageOutstanding) = S_ERROR
  ben.value(ln_InterestDaily) = 0
  ben.value(ln_AverageLoanRate) = 0
  ben.value(ln_OpenOutstanding) = OpenOutStanding 'also sets ln_OpenDate
  ben.value(ln_CloseOutstanding) = CloseOutStanding 'also set ln_CloseDate
  
  ben.value(ln_WholeTaxMonthsOutstanding) = 0
  
  For i = 1 To BalanceSheet.Count
    Set bi = BalanceSheet(i)
    If Not bi Is Nothing Then
      ben.value(ln_InterestDaily) = ben.value(ln_InterestDaily) + bi.Interest(ben.value(ln_LoanCurrencyIndex))
      If (ben.value(ln_LoanCurrencyIndex) <> INTERESTRATE_TYPE.IT_STERLING) Then
        ben.value(ln_AverageLoanRate) = ben.value(ln_AverageLoanRate) + (bi.InterestRate(ben.value(ln_LoanCurrencyIndex)) * bi.days)
      End If
      'HMRC no longer seem to actually calc the average rate
      'we must thus use that published at
      'http://hmrc.gov.uk/rates/interest-beneficial.htm, this address is not avaible at March 2011
      'and update in pd.interest rates for actual values and D_ACCOMINTRATE for the offical rate for year and in NUmbers and dates
      lDays = lDays + bi.days
    End If
  Next
  
  If (ben.value(ln_LoanCurrencyIndex) = INTERESTRATE_TYPE.IT_STERLING) Then 'sterling is special as rates caN VARY IN THE YEAR BUT OFFICIAL RATE MAY NOT EQUAL THE AVERAGE OF THE unDERLYING RATES
    ben.value(ln_AverageLoanRate) = p11d32.Rates.value(accomInterestRate)  'IRLoanRateAdjustmentDailyInterestRate(ben.value(ln_AverageLoanRate), lDays)
  Else
    ben.value(ln_AverageLoanRate) = IRLoanRateAdjustmentDailyInterestRate(ben.value(ln_AverageLoanRate), lDays)
  End If
  
  ben.value(ln_InterestDaily) = RoundDownEx(ben.value(ln_InterestDaily), 0) 'EK added RoundDown 1/04 TTP#86
  
  
  ben.value(ln_AverageOutstanding) = (ben.value(ln_OpenOutstanding) + ben.value(ln_CloseOutstanding)) / 2
  ben.value(ln_WholeTaxMonthsOutstanding) = WholeTaxMonthsOutstanding(ben.value(ln_opendate), ben.value(ln_closedate))
  ben.value(ln_InterestNormal) = RoundDownEx((ben.value(ln_AverageOutstanding) * ben.value(ln_AverageLoanRate)) * (ben.value(ln_WholeTaxMonthsOutstanding) / 12), 0) 'EK added RoundDown 1/04 TTP#86
  
  ben.value(ln_CashEquivalentNormal) = RoundDownEx(Max(0, ben.value(ln_InterestNormal) - ben.value(ln_InterestPaid_db)), 0) 'EK added RoundDown 1/04 TTP#86
  ben.value(ln_CashEquivalentDaily) = RoundDownEx(Max(0, ben.value(ln_InterestDaily) - ben.value(ln_InterestPaid_db)), 0) 'EK added RoundDown 1/04 TTP#86

  If ben.value(ln_AverageOutstanding) > ben.value(ln_DefaultMirasRelief) Then
    ben.value(ln_DeemedInterestNormal) = RoundN((ben.value(ln_DefaultMirasRelief) / ben.value(ln_AverageOutstanding)) * ben.value(ln_CashEquivalentNormal) * p11d32.Rates.value(LoanMortgageRate))
    ben.value(ln_DeemedInterestAlternative) = RoundN((ben.value(ln_DefaultMirasRelief) / ben.value(ln_AverageOutstanding)) * ben.value(ln_CashEquivalentDaily) * p11d32.Rates.value(LoanMortgageRate))
  Else
    ben.value(ln_DeemedInterestNormal) = RoundN(ben.value(ln_InterestNormal) * p11d32.Rates.value(LoanMortgageRate))
    ben.value(ln_DeemedInterestAlternative) = RoundN(ben.value(ln_InterestDaily) * p11d32.Rates.value(LoanMortgageRate))
  End If

  CalculateInterest = True

CalculateInterest_End:
  Set bi = Nothing
  Call xReturn("CalculateInterest")
  Exit Function

CalculateInterest_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CalculateInterest", "Calculate Interest", "Error calculating the loan interest.")
  Resume CalculateInterest_End
  Resume
End Function

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim BI0 As BalanceItem, BI1 As BalanceItem
  
  Set BI0 = v0
  Set BI1 = v1
  
  If BI1 Is Nothing And BI0 Is Nothing Then
    ISortFunction_CompareItems = 0
    Exit Function
  End If
  
  If BI1 Is Nothing Then
    ISortFunction_CompareItems = 1
    Exit Function
  End If
  
  If BI0 Is Nothing Then
    ISortFunction_CompareItems = -1
    Exit Function
  End If
  
  If BI0.DateFrom < BI1.DateFrom Then
    ISortFunction_CompareItems = -1
  ElseIf BI0.DateFrom > BI1.DateFrom Then
    ISortFunction_CompareItems = 1
  End If
  
End Function



