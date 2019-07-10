VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Accommodation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IBenefitClass


Public Enum AccomItems
  accom_item_db = ITEM_DESC
  accom_value = ITEM_VALUE
  accom_MadeGood_Net = ITEM_MADEGOOD_NET
  accom_Benefit = ITEM_BENEFIT
  accom_MadeGood = ITEM_MADEGOOD
  accom_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  accom_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  accom_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  accom_BoxNumber = ITEM_BOX_NUMBER
  accom_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  accom_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  accom_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  accom_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  accom_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  accom_ERROR = ITEM_ERROR
  accom_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  accom_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  accom_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE
  accom_IRDesc

 
  'MAPPINGS TO O Other
  accom_EXPitem
  accom_EXP_SPECIFIC_START = accom_EXPitem
  accom_EXPValue
  accom_EXPMadeGood_NET
  accom_EXPBenefit
  accom_EXPMadeGood
  accom_EXPActualAmountMadeGood
  accom_EXPBenefit_Reportable
  accom_EXPUDM_BENEFIT_TITLE
  accom_EXPBoxNumber
  accom_EXPMadeGoodIsTaxDeducted
  accom_EXPClass1AAdjustment
  accom_EXPNICClass1AAble
  accom_EXPNIC_Class1A_Value
  accom_EXPNIC_Class1A_Benefit
  accom_EXP_Error
  accom_EXP_Value_Non_OPRA
  accom_EXP_Ammount_Foregone_Used_For_Value
  accom_EXP_OPRA_Ammount_Foregone_db
  accom_EXP_IRDesc
  accom_EXPGrossAmountPaidByEmployer
  accom_EXPEmployerReference
  accom_EXPcategory
  accom_EXPHMITCode_UDBCode
  accom_EXPCompanyDefinedCategory
  accom_EXPCompanyDefinedCategoryKey
  accom_EXPclass
  accom_EXPAvailablefrom
  accom_EXPavailableTo
  accom_EXPamountwaived
  accom_EXP_SPECIFIC_END = accom_EXPamountwaived
  'end mappings
  
  accom_Business_db 'bf
  accom_Price_db 'bf
  accom_Improv_db
  accom_MVFirstOcc_db 'bf
  accom_availableto_db 'bf test
  accom_FirstOcc_db 'bf
  accom_JobRelated_db 'bf
  accom_SixYear_db 'bf
  accom_isrent_db 'bf
  accom_erown_db 'bf
  accom_rent_db
  accom_ConsiderationForUse_db
  
'MP DB (not used)           accom_Unavailable
'MP DB - removed accom_RelevantDays_db as only in ReadDB but not in use elsewhere
' accom_RelevantDays_db
  
  accom_CapContrib_db
  accom_Tenancy_db
  
  accom_RecentImprov_db
'MP_DB  accom_Utilities
'MP DB  accom_Repairs
'MP DB  accom_Assets
'MP DB  accom_Reimbursements
  accom_Expenses_MadeGood
'MP DB  accom_NetEmoluments
'MP DB - removed below as ReadDB+WriteDB but not in use elsewhere
'  accom_AvailDays_db
  
'MP DB  accom_Ancillary
  
'MP DB  accom_nemployees
  accom_nemployees_db
  
  accom_basiccharge
  accom_expenses
  accom_netexpenses
  accom_CostOfProvision
  accom_expensivecharge
'MP DB (not used)           accom_AllowedExpenses
  accom_ShareOfBen
  accom_NonBusinessBen
  accom_ExpensesNonBusiness
  accom_ShareOfExpenses
  

'items to temporarity store different ben values for the NIC calulation
  accom_temp_nic_calc_value
  accom_temp_nic_calc_benefit
  accom_temp_nic_calc_MadeGoodIsTaxDeducted
  accom_temp_nic_calc_accom_MadeGood_Net
  accom_temp_nic_calc_accom_NICClass1AAble
'end items to temporarity store different ben values for the NIC calulation

  accom_availablefrom_db
  Accom_LASTITEM = accom_availablefrom_db
End Enum

Public OOther As other

Private m_InvalidFields As Long
Private m_BenClass  As BEN_CLASS
Private m_Parent As Object
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_BenItems(1 To Accom_LASTITEM)

Private m_dirty As Boolean
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Accomodation - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim lDays As Long
  Dim ben As IBenefitClass
  Dim normalNICCalc As Boolean
  On Error GoTo accommodation_Error:
   
  Set ben = Me
   
   
  Call ben.SetCalcDefaults
  
  'linfs to OOther
  ben.value(accom_EXPitem) = ben.value(accom_item_db)
  
  
  'Basic Charge
  ben.value(accom_basiccharge) = ben.value(accom_rent_db)
  ' ben.value(accom_expenses) = ben.value(accom_Utilities) + ben.value(accom_Repairs) + ben.value(accom_Reimbursements) + ben.value(accom_Ancillary)
  
  ben.value(accom_ExpensesNonBusiness) = ben.value(accom_expenses) * (1# - (ben.value(accom_Business_db) / 100#))
  
  ben.value(accom_ShareOfExpenses) = ben.value(accom_ExpensesNonBusiness) / (ben.value(accom_nemployees_db) + 1)
  
'TODO MP DB - need to revisit this If condition
  If ben.value(accom_JobRelated_db) Then
    'MP DB - accom_NetEmoluments is Not used so ben.value(accom_NetEmoluments) would have always returned 0
    'MP DB      ben.value(accom_ShareOfExpenses) = Min(ben.value(accom_NetEmoluments) * p11d32.Rates.value(accomJobRelatedPercentage), ben.value(accom_ShareOfExpenses))
    ben.value(accom_ShareOfExpenses) = Min(0, ben.value(accom_ShareOfExpenses))
  End If
  
  'Expensive charge
  If ben.value(accom_erown_db) Then
    ben.value(accom_CostOfProvision) = (ben.value(accom_Price_db) + ben.value(accom_Improv_db)) - ben.value(accom_CapContrib_db) - ben.value(accom_Tenancy_db)
    If ben.value(accom_SixYear_db) Then
      If ben.value(accom_CostOfProvision) > p11d32.Rates.value(accomExpensiveChargeLimit) Then
        ben.value(accom_CostOfProvision) = ben.value(accom_MVFirstOcc_db) + ben.value(accom_RecentImprov_db) - (ben.value(accom_CapContrib_db) + ben.value(accom_Tenancy_db))
      End If
    End If
    ben.value(accom_expensivecharge) = RoundDown(Max((ben.value(accom_CostOfProvision) - p11d32.Rates.value(accomExpensiveChargeLimit)), 0) * p11d32.Rates.value(accomInterestRate), 0)
  Else
    ben.value(accom_expensivecharge) = 0&
  End If
  'Apportionments
  ben.value(accom_ShareOfBen) = (ben.value(accom_basiccharge) + ben.value(accom_expensivecharge)) / (ben.value(accom_nemployees_db) + 1)

  ben.value(accom_NonBusinessBen) = ben.value(accom_ShareOfBen) * (1 - (ben.value(accom_Business_db) / 100#))
  
  
  ben.value(accom_temp_nic_calc_value) = ben.value(accom_ShareOfBen)
  
  If ben.value(accom_isrent_db) Then
    ben.value(ITEM_VALUE_NON_OPRA) = ben.value(accom_NonBusinessBen)
    
  Else
    ben.value(ITEM_VALUE_NON_OPRA) = RoundDownEx(ben.value(accom_NonBusinessBen) * dGetDateFactor(lDays, ben.value(accom_availablefrom_db), ben.value(accom_availableto_db)), 0)
    ben.value(accom_temp_nic_calc_value) = ben.value(ITEM_VALUE_NON_OPRA)
  End If
  
  If ben.value(accom_JobRelated_db) Then
    ben.value(ITEM_VALUE_NON_OPRA) = 0
    ben.value(accom_temp_nic_calc_value) = 0
  End If
  
  
  ben.value(accom_netexpenses) = Max(ben.value(accom_ShareOfExpenses) - ben.value(accom_Expenses_MadeGood), 0)
  
  'links to O OOther
  ben.value(accom_EXPAvailablefrom) = p11d32.Rates.value(TaxYearStart)
  ben.value(accom_EXPavailableTo) = p11d32.Rates.value(TaxYearEnd)
  
  ben.value(accom_EXPGrossAmountPaidByEmployer) = ben.value(accom_ShareOfExpenses)
  ben.value(accom_EXPValue) = ben.value(accom_ShareOfExpenses)
  
  ben.value(accom_EXPMadeGood_NET) = Min(ben.value(accom_ShareOfExpenses), ben.value(accom_Expenses_MadeGood))
  ben.value(accom_EXPMadeGood) = ben.value(accom_Expenses_MadeGood)
  ben.value(accom_EXPBenefit) = ben.value(accom_netexpenses)
  
  If ben.value(accom_EXPValue) > 0 Then ben.value(accom_EXPBenefit_Reportable) = True
  'end links to O Other
  
  ben.value(accom_MadeGood) = ben.value(accom_ConsiderationForUse_db)
  
  Call CalculateOpRAValue(ben)
  ben.value(accom_MadeGood_Net) = Min(ben.value(ITEM_VALUE), ben.value(accom_MadeGood))
  
  If (ben.value(ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE)) Then
    ben.value(accom_temp_nic_calc_value) = ben.value(ITEM_VALUE)
  End If
  
  
  ben.value(accom_Benefit) = ben.value(ITEM_VALUE) - ben.value(accom_MadeGood_Net)
  
 
  ben.value(accom_temp_nic_calc_benefit) = ben.value(accom_temp_nic_calc_value) - ben.value(accom_MadeGood_Net)
  Call BenCalcNIC(ben, , , accom_temp_nic_calc_value, accom_temp_nic_calc_benefit)
  
  IBenefitClass_CalculateBody = ben.value(accom_Benefit)
  
  If ben.value(accom_Benefit) > 0 Then ben.value(accom_Benefit_Reportable) = True
  
accommodation_End:
  Exit Function
accommodation_Error:
  IBenefitClass_CalculateBody = S_ERROR
  Resume accommodation_End
 Resume
End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Accommodation"
End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
  
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = m_NeedToCalculate
End Property


Private Property Let IBenefitClass_LinkBen(RHS As Boolean)

End Property

Private Property Get IBenefitClass_LinkBen() As Boolean

End Property

Private Function IBenefitClass_CanBringForward() As Boolean
  IBenefitClass_CanBringForward = StandardCanBringForward(Me, accom_availableto_db)
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim accom As Accommodation
  
  On Error GoTo Copy_END
  Call xSet("Copy")
  
  Set accom = New Accommodation
  Set IBenefitClass_Copy = CopyBenStandard(Parent, accom, Me)
  'EK removal of accomodation expenses TTP#28
  ' Call accom.AddOOther
  
Copy_END:
  Call xReturn("Copy")
  Exit Function
Copy_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Copy", "Copy", "Error copying an accomodation benefit.")
  Resume Copy_ERR
End Function




Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_LIVING_ACCOMMODATION_D
End Property

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property


Public Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
  'No Code
End Property
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass
  Dim tempStartDate As String
  Dim tempEndDate As String
  Dim tempDateDiff As String
  
  On Error GoTo Accommodation_PrintWKBody_Err
  Call xSet("Accommodation_PrintWKBody")
  
  Set ben = Me
    
'details
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Property: ", ben.value(ITEM_DESC))
  Call WKOut(rep, WK_ITEM_TEXT, "Provided to the employee from " & S_WK_NORMAL_BOLD_FONT & ben.value(accom_availablefrom_db) & S_WK_NORMAL_FONT & " to " & S_WK_NORMAL_BOLD_FONT & ben.value(accom_availableto_db))
  'Job releated
  If ben.value(accom_JobRelated_db) Then
    Call WKOut(rep, WK_ITEM_TEXT, "This accommodation qualifies as job related under s.98-101 ITEPA 2003")
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "This accommodation does not qualify as job related under s.98-101 ITEPA 2003")
  End If
  
'calc
  Call WKOut(rep, WK_SECTION_HEADER_VALUE)
  'Is Rent
  If Not ben.value(accom_isrent_db) And ben.value(accom_erown_db) Then
    Call WKOut(rep, WK_ITEM_TEXT, "Greater of annual rent and gross rateable value - basic charge  s.105 ITEPA 2003", ben.value(accom_basiccharge))
    Call WKOut(rep, WK_ITEM_TEXT, "Additional charge s.106 ITEPA 2003 See below ", ben.value(accom_expensivecharge))
    Call WKOut(rep, WK_ITEM_Total, "Total basic and additional charge ", ben.value(accom_basiccharge) + ben.value(accom_expensivecharge))
    
    If ben.value(accom_SixYear_db) Then
      Call WKOut(rep, WK_ITEM_TEXT, "The employee first occupied the property on " & ben.value(accom_FirstOcc_db))
      Call WKOut(rep, WK_ITEM_TEXT, "The property was owned by the company through the six year period to " & p11d32.Rates.value(TaxYearStart))
      Call WKOut(rep, WK_ITEM_TEXT, "Market value when first provided to the employee", ben.value(accom_MVFirstOcc_db), , True)
      Call WKOut(rep, WK_ITEM_TEXT, "Cost of improvements prior to " & p11d32.Rates.value(TaxYearStart) & " but after " & ben.value(accom_FirstOcc_db), ben.value(accom_RecentImprov_db))
      'less
      Call WKOut(rep, WK_ITEM_TEXT, "Capital expenditure reimbursed by the employee", ben.value(accom_CapContrib_db), , , True)
      Call WKOut(rep, WK_ITEM_TEXT, "Payment by employee for the grant of the tenancy", ben.value(accom_Tenancy_db), , , True)
      
      Call WKOut(rep, WK_ITEM_Total, "Cost of providing accommodation under s.107 ITEPA 2003", ben.value(accom_CostOfProvision), , True)
    Else
      Call WKOut(rep, WK_ITEM_TEXT, "Calculation of, s.106 ITEPA 2003, additional charge:")
      Call WKOut(rep, WK_ITEM_TEXT, "Cost of acquisition", ben.value(accom_Price_db), , True)
      Call WKOut(rep, WK_ITEM_TEXT, "Cost of improvements to the property prior to " & DateValReadToScreen(p11d32.Rates.value(TaxYearStart)), ben.value(accom_Improv_db))
      Call WKOut(rep, WK_ITEM_TEXT, "Capital expenditure reimbursed by the employee", ben.value(accom_CapContrib_db), , , True)
      Call WKOut(rep, WK_ITEM_TEXT, "Payment by employee for the grant of the tenancy", ben.value(accom_Tenancy_db), , , True)
      
      Call WKOut(rep, WK_ITEM_Total, "Cost of providing accommodation under s146(4)", ben.value(accom_CostOfProvision), , True)
      Call WKOut(rep, WK_ITEM_TEXT, "")
    End If
    
    Call WKOut(rep, WK_ITEM_TEXT, "Interest rate at " & p11d32.Rates.value(TaxYearStart) & " = " & Format(p11d32.Rates.value(accomInterestRate), "PERCENT"))
    Call WKOut(rep, WK_ITEM_Total, "Additional charge, s.106 ITEPA 2003: " & Format(p11d32.Rates.value(accomInterestRate), "PERCENT") & " of " & ben.value(accom_CostOfProvision) & IIf(ben.value(accom_CostOfProvision) > 0, " -" & p11d32.Rates.value(accomExpensiveChargeLimit), "") & IIf(ben.value(accom_expensivecharge) = 0, " (Reduced to Zero)", ""), ben.value(accom_expensivecharge), , True)
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "Basic charge (excluding assets and services benefit)", ben.value(accom_basiccharge))
    Call WKOut(rep, WK_ITEM_TEXT, "Greater of annual rent and gross rateable value - basic charge s145")
    Call WKOut(rep, WK_ITEM_Total, "Basic and additional charge, s.105/106 ITEPA 2003, reduced by " & ben.value(accom_Business_db) & "% for business use of the property", ben.value(accom_basiccharge) + ben.value(accom_expensivecharge), , True)
  End If
  
  ' EK removal of accomodation expenses TTP#28
  
'  If ben.value(accom_netexpenses) > 0 Then
'    Call WKOut(rep, WK_SECTION_HEADER, "Benefit of expenses paid in connection with the accommodation")
'    Call WKOut(rep, WK_ITEM_TEXT, "Expenses incurred by the employer in respect of the accommodation")
'    Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "(a) Heating, lighting and cleaning", ben.value(accom_Utilities), , True)
'    Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "(b) Repairs, maintenance and decoration", ben.value(accom_Repairs))
'    Call WKOut(rep, WK_ITEM_TEXT, "Reimbursement of employee's expenditure in respect of (a)+(b) above", ben.value(accom_Reimbursements))
'    Call WKOut(rep, WK_ITEM_TEXT, "Other ancillary expenses not covered above", ben.value(accom_Ancillary))
    
'    If ben.value(accom_JobRelated) Then
'      Call WKOut(rep, WK_ITEM_TEXT, "", ben.value(accom_expenses), "'A'", True)
'      Call WKOut(rep, WK_ITEM_TEXT, "Net emoluments as defined in s163(4) ", ben.value(accom_NetEmoluments), , True)
'      Call WKOut(rep, WK_ITEM_TEXT, "Possible limitation to assets and services charge under s163(3) :")
'      Call WKOut(rep, WK_ITEM_Total, "10 percent of net emoluments, less expenditure made good by the employee", "", "'B'")
'      Call WKOut(rep, WK_ITEM_Total, "Lesser of 'A' and 'B'", ben.value(accom_ShareOfExpenses), "'C'")
'    Else
'      Call WKOut(rep, WK_ITEM_Total, "Total Expenses", ben.value(accom_expenses), , True)
'    End If
    
'    Call WKOut(rep, WK_ITEM_TEXT, "Less :Expenditure on services made good by the employee", ben.value(accom_Expenses_MadeGood), , True, True)
'    Call WKOut(rep, WK_ITEM_Total, "Net expenses paid by the employer", ben.value(accom_expenses) - ben.value(accom_Expenses_MadeGood), , True)
'    If (ben.value(accom_Business) > 0) Then
'      Call WKOut(rep, WK_ITEM_TEXT, "Reduced by " & ben.value(accom_Business) & "% for business use of the property", ben.value(accom_ExpensesNonBusiness))
'    End If
    
'    If (ben.value(accom_nemployees) > 0) Then
'      Call WKOut(rep, WK_ITEM_Total, "Divided by the number of employees resident in this property: " & ben.value(accom_nemployees) + 1, ben.value(accom_ShareOfExpenses), , True)
'    End If
    
    'Call WKOut(rep, WK_BLANK_LINE)
    ' Call WKOut(rep, WK_ITEM_TEXT_BOLD, "All expenses are reported as part of the '" & p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_HMIT_SECTION_STRING) & ":" & p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_FORM_CAPTION) & "' section.")
'    Call WKOut(rep, WK_BLANK_LINE)
'  Else
    Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)
'  End If
  
  If (ben.value(accom_nemployees_db) > 0) Then
    Call WKOut(rep, WK_ITEM_TEXT, "Divided by the number of employees resident in this property: " & ben.value(accom_nemployees_db) + 1, ben.value(accom_ShareOfBen), , True)
  End If
  
  If (ben.value(accom_Business_db) > 0) Then
    Call WKOut(rep, WK_ITEM_TEXT, "Reduced by " & ben.value(accom_Business_db) & "% for business use of the property", ben.value(accom_NonBusinessBen))
  End If
  
  If (ben.value(accom_nemployees_db) > 0) Or (ben.value(accom_Business_db) > 0) Then
    Call WKOut(rep, WK_BLANK_LINE)
  End If
  
  ' Overall accommodation benefit
  'EK add date calculation
  If Not ben.value(accom_availablefrom_db) = p11d32.Rates.value(TaxYearStart) Or Not ben.value(accom_availableto_db) = p11d32.Rates.value(TaxYearEnd) Then
    tempStartDate = Format$(ben.value(accom_availablefrom_db), "dd/mm/yyyy")
    tempEndDate = Format$(ben.value(accom_availableto_db), "dd/mm/yyyy")
    tempDateDiff = DateDiff("d", DateAdd("D", -1, ben.value(accom_availablefrom_db)), ben.value(accom_availableto_db))
    Call WKTblColXOffsets(L_WK_ACCOM_TABLE_COL1 - 1, L_WK_ACCOM_TABLE_COL2, L_WK_ACCOM_TABLE_COL3, L_WK_ACCOM_TABLE_COL4, L_WK_ACCOM_TABLE_COL5, L_WK_ACCOM_TABLE_COL6)
    Call WKTblColFormats("n", "n", "n", "rn", "rn", "rn")
    Call WKTableHeadings(rep, "", S_FIELD_AVAILABLE_FROM, "Available To", "Days Available", "Annual Benefit", "Benefit")
    Call WKTableRow(rep, "Basic and additional charge, s.105/106 ITEPA 2003", tempStartDate, tempEndDate, tempDateDiff, FormatWN(ben.value(accom_basiccharge) + ben.value(accom_expensivecharge)), FormatWN(ben.value(ITEM_VALUE_NON_OPRA)))  ' RDC added 12/02/02
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "Basic and additional charge, s.106 ITEPA 2003", ben.value(accom_basiccharge) + ben.value(accom_expensivecharge), , True)
  End If
  
  Call WKOut(rep, WK_BLANK_LINE)
  
  Call OPRAWorkingPaperValue(ben, rep)
  
    
  Call WKOut(rep, WK_ITEM_TEXT, "Consideration from employee for the use of the accommodation")
  Call WKOut(rep, WK_ITEM_TEXT, "or amount from which tax has been deducted under PAYE", ben.value(accom_ConsiderationForUse_db), , , True)
  
  Call WKOut(rep, WK_ITEM_Total, "Overall accommodation benefit" & IIf(ben.value(accom_JobRelated_db), " (Reduced to zero as job-related)", ""), ben.value(accom_Benefit), , True)
  
    
Accommodation_PrintWKBody_End:
  Call xReturn("Accommodation_PrintWKBody")
  Exit Function

Accommodation_PrintWKBody_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Accommodation_PrintWKBody", "Services Provided Print WK Body", "Error printing the Accommodation working paper")
  Resume Accommodation_PrintWKBody_End
  Resume

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
  Call SetCalcDefaultsStandard(Me)
  ben.value(accom_CostOfProvision) = 0
End Sub

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_LIVING_ACCOMMODATION_D, Item, RHS)
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(accom_item_db)
End Property
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property
Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
Private Property Get IBenefitClass_PrintHeader() As String
  ECASE ("IBenefitClass_PrintHeader")
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim AccomBen As IBenefitClass
  Dim accom As Accommodation
  Dim ben As IBenefitClass
  
  Dim rs As Recordset
  Dim s As String
  Dim i As Long
  
  On Error GoTo clsAccom_ReadDB_Err
  Call xSet("clsAccom_ReadDB")
  
  If m_ReadFromDB Then GoTo clsAccom_ReadDB_End
  Set ben = Me
  Set rs = m_Parent.Parent.rsBenTables(TBL_ACCOM)
  If Len(m_sbookmark) = 0 Then
    rs.FindFirst ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
    If Not rs.NoMatch Then

      m_sbookmark = rs.Bookmark
      i = i + 1
      rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set AccomBen = New Accommodation
        Set accom = AccomBen
        
        AccomBen.RSBookMark = rs.Bookmark
        Set AccomBen.Parent = m_Parent
        Call m_Parent.benefits.Add(AccomBen)
        'EK removal of accomodation expenses TTP#28
        ' Call accom.AddOOther
        Set AccomBen = Nothing
        rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If

  If Len(m_sbookmark) > 0 And (Not ben.ReadFromDB) Then
    rs.Bookmark = m_sbookmark
    
    Call StandardReadData(ben, rs)
    
    ben.value(accom_item_db) = "" & rs.Fields("ITEM").value
    
    ben.value(accom_rent_db) = IIf(IsNull(rs.Fields("Rent").value), 0, rs.Fields("Rent").value)
    ben.value(accom_Business_db) = IIf(IsNull(rs.Fields("Business").value), 0, rs.Fields("Business").value)
'MP DB - removed accom_RelevantDays_db as ReadDB but not in use elsewhere
'    ben.value(accom_RelevantDays_db) = IIf(IsNull(rs.Fields("AvailDays").value), 0, rs.Fields("AvailDays").value)
    ben.value(accom_ConsiderationForUse_db) = IIf(IsNull(rs.Fields("Consideration").value), 0, rs.Fields("Consideration").value)
    ben.value(accom_Price_db) = IIf(IsNull(rs.Fields("Price").value), 0, rs.Fields("Price").value)
    ben.value(accom_Improv_db) = IIf(IsNull(rs.Fields("Improv").value), 0, rs.Fields("Improv").value)
    ben.value(accom_CapContrib_db) = IIf(IsNull(rs.Fields("CapContrib").value), 0, rs.Fields("CapContrib").value)
    ben.value(accom_Tenancy_db) = IIf(IsNull(rs.Fields("Tenancy").value), 0, rs.Fields("Tenancy").value)
    ben.value(accom_MVFirstOcc_db) = IIf(IsNull(rs.Fields("MVFirstOcc").value), 0, rs.Fields("MVFirstOcc").value)
    ben.value(accom_RecentImprov_db) = IIf(IsNull(rs.Fields("RecentImprov").value), 0, rs.Fields("RecentImprov").value)

    ' ben.value(accom_Utilities) = IIf(IsNull(rs.Fields("Utilities").value), 0, rs.Fields("Utilities").value)
    ' ben.value(accom_Repairs) = IIf(IsNull(rs.Fields("Repairs").value), 0, rs.Fields("Repairs").value)
    ' ben.value(accom_Assets) = IIf(IsNull(rs.Fields("Assets").value), 0, rs.Fields("Assets").value)
    ' ben.value(accom_Reimbursements) = IIf(IsNull(rs.Fields("Reimbursements").value), 0, rs.Fields("Reimbursements").value)
    ' ben.value(accom_Expenses_MadeGood) = IIf(IsNull(rs.Fields("MadeGood").value), 0, rs.Fields("MadeGood").value)
    ' ben.value(accom_NetEmoluments) = IIf(IsNull(rs.Fields("NetEmoluments").value), 0, rs.Fields("NetEmoluments").value)
'MP DB ref declaration
'   ben.value(accom_AvailDays_db) = IIf(IsNull(rs.Fields("AvailDays").value), 0, rs.Fields("AvailDays").value)
    
'MP DB  ben.value(accom_Ancillary) = IIf(IsNull(rs.Fields("Ancillary").value), 0, rs.Fields("Ancillary").value)
    ben.value(accom_nemployees_db) = IIf(IsNull(rs.Fields("nemployees").value), 0, rs.Fields("nemployees").value)
    
    ben.value(accom_availablefrom_db) = IIf(IsNull(rs.Fields("AvailFrom").value), p11d32.Rates.value(TaxYearStart), rs.Fields("AvailFrom").value)
    ben.value(accom_availableto_db) = IIf(IsNull(rs.Fields("AvailTo").value), p11d32.Rates.value(TaxYearEnd), rs.Fields("AvailTo").value)
    ben.value(accom_FirstOcc_db) = IIf(IsNull(rs.Fields("FirstOcc").value), p11d32.Rates.value(TaxYearStart), rs.Fields("FirstOcc").value)
    
    ben.value(accom_JobRelated_db) = rs.Fields("JobRelated").value
    ben.value(accom_SixYear_db) = rs.Fields("SixYear").value
    ben.value(accom_isrent_db) = rs.Fields("IsRent").value
    ben.value(accom_erown_db) = rs.Fields("ErOwn").value
    'OOther Link
    
    ' Call StandardReadData(OOther)
    m_ReadFromDB = True
  End If

clsAccom_ReadDB_End:
  IBenefitClass_ReadDB = i
  Set accom = Nothing
  Call xReturn("clsAccom_ReadDB")
  Exit Function
clsAccom_ReadDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsAccom_ReadDB", "ERR_VAN_READDB", "Unable to load the accommodation details")
  Resume clsAccom_ReadDB_End
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property


Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim ben As IBenefitClass
  
  Dim bc As BEN_CLASS
  
  On Error GoTo SetBenItemsInformation_err
  
  Set ben = Me
  m_NeedToCalculate = True
  
  bc = BC_LIVING_ACCOMMODATION_D
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
    
  With p11d32
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    .BenDataLinkDataType(bc, accom_EXPitem) = TYPE_STR
    .BenDataLinkDataType(bc, accom_EXPValue) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_EXPMadeGood_NET) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_EXPBenefit) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_EXPMadeGood) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_EXPBenefit_Reportable) = TYPE_BOOL
    
    .BenDataLinkDataType(bc, accom_rent_db) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_ConsiderationForUse_db) = TYPE_LONG
    
    .BenDataLinkDataType(bc, accom_Business_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_Business_db) = "Percentage of business use"
    
    .BenDataLinkDataType(bc, accom_Price_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_Price_db) = "Purchase price"
    
    .BenDataLinkDataType(bc, accom_Improv_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_Improv_db) = "Cost of improvements"
    
    
    .BenDataLinkDataType(bc, accom_CapContrib_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_CapContrib_db) = "Capital contribution"
    
    
    
    .BenDataLinkDataType(bc, accom_Tenancy_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_Tenancy_db) = "Payments for grant of tenancy"
    
    .BenDataLinkDataType(bc, accom_MVFirstOcc_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_MVFirstOcc_db) = "Market value first date occupied"
    
    
    
    .BenDataLinkDataType(bc, accom_RecentImprov_db) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, accom_RecentImprov_db) = "Improvements in tax year"
    
    
    
    .BenDataLinkDataType(bc, accom_Expenses_MadeGood) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_nemployees_db) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_basiccharge) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_expenses) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_netexpenses) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_CostOfProvision) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_expensivecharge) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_ShareOfBen) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_NonBusinessBen) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_ExpensesNonBusiness) = TYPE_LONG
    .BenDataLinkDataType(bc, accom_ShareOfExpenses) = TYPE_LONG
    
    .BenDataLinkDataType(bc, accom_availablefrom_db) = TYPE_DATE
    .BenDataLinkUDMDescription(bc, accom_availablefrom_db) = "Date first occupied"
    
    .BenDataLinkDataType(bc, accom_availableto_db) = TYPE_DATE
    .BenDataLinkDataType(bc, accom_FirstOcc_db) = TYPE_DATE
    
    .BenDataLinkDataType(bc, accom_JobRelated_db) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, accom_JobRelated_db) = "Is accommodation job related"
    
    .BenDataLinkDataType(bc, accom_SixYear_db) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, accom_SixYear_db) = "Does accommodation meet 6 year rule"
    
    
    .BenDataLinkDataType(bc, accom_isrent_db) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, accom_isrent_db) = "Is amount paid by employee rent"
    
    .BenDataLinkDataType(bc, accom_erown_db) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, accom_erown_db) = "Property owned by employer"
    
    'UDM data
    
    .BenDataLinkUDMDescription(bc, accom_rent_db) = "Rent"
    .BenDataLinkUDMDescription(bc, accom_availablefrom_db) = S_UDM_FROM
    .BenDataLinkUDMDescription(bc, accom_availableto_db) = S_UDM_To
    
    Call SetStandardBenItemsDataTypes(bc)
    Call SetStandardBenItemsMMFieldSize(bc)
    Call SetStandardBenItemsUDMData(bc)
    
    
  End With
  
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information.")
  Resume SetBenItemsInformation_end
  Resume
End Sub

Private Sub IBenefitClass_Kill()
  If Not OOther Is Nothing Then Set OOther.Accommodation = Nothing
  Set OOther = Nothing
  Set m_Parent = Nothing
End Sub

Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  
  On Error GoTo accommodation_DeleteDB_Err
  Call xSet("accommodation_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_ACCOM)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
  
accommodation_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("accommodation_DeleteDB")
  Exit Function
accommodation_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "accommodation_DeleteDB", "accommodation Delete DB", "Error deleting the accommodation benefit.")
  Resume accommodation_DeleteDB_End
End Function


Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo accommodation_WriteDB_Err
  Call xSet("accommodation_WriteDB")
  
  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  ben.NeedToCalculate = True
  
  rs.Fields("ITEM").value = ben.value(accom_item_db)
  rs.Fields("Business").value = ben.value(accom_Business_db)
  rs.Fields("Price").value = ben.value(accom_Price_db)
  rs.Fields("Improv").value = ben.value(accom_Improv_db)
  rs.Fields("MVFirstOcc").value = ben.value(accom_MVFirstOcc_db)
  
  rs.Fields("FirstOcc").value = ben.value(accom_FirstOcc_db)
  rs.Fields("JobRelated").value = ben.value(accom_JobRelated_db)
  rs.Fields("SixYear").value = ben.value(accom_SixYear_db)
  rs.Fields("IsRent").value = ben.value(accom_isrent_db)
  rs.Fields("ErOwn").value = ben.value(accom_erown_db)
  rs.Fields("nemployees").value = ben.value(accom_nemployees_db)
  
  Call BringForwardDatesWrite(ben, accom_availablefrom_db, accom_availableto_db, rs, "AvailFrom", "AvailTo")
  
  Call StandardWriteData(ben, rs)
  
    rs.Fields("CapContrib").value = ben.value(accom_CapContrib_db)
    rs.Fields("Tenancy").value = ben.value(accom_Tenancy_db)
  
  If Not p11d32.BringForward.Yes Then
    rs.Fields("Rent").value = ben.value(accom_rent_db)
    rs.Fields("Consideration").value = ben.value(accom_ConsiderationForUse_db)
    rs.Fields("RecentImprov").value = ben.value(accom_RecentImprov_db)
    
    
'MP DB - never read to assign enum from DB
'   rs.Fields("Assets").value = ben.value(accom_Assets)
'MP DB    rs.Fields("AvailDays").value = ben.value(accom_AvailDays_db)
    
    'EK removal of accomodation expenses TTP#28
    'rs.Fields("Utilities").value = ben.value(accom_Utilities)
    'rs.Fields("Repairs").value = ben.value(accom_Repairs)
    'rs.Fields("Reimbursements").value = ben.value(accom_Reimbursements)
    'rs.Fields("MadeGood").value = ben.value(accom_Expenses_MadeGood)
    'rs.Fields("NetEmoluments").value = ben.value(accom_NetEmoluments)
    'rs.Fields("Ancillary").value = ben.value(accom_Ancillary)
  Else
    If Not ben.value(accom_isrent_db) Then rs.Fields("Rent").value = ben.value(accom_rent_db)
  End If
  
  IBenefitClass_WriteDB = BenefitCloseRecord(ben, rs)
  
accommodation_WriteDB_End:
  Set rs = Nothing
  Call xReturn("accommodation_WriteDB")
  Exit Function

accommodation_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "accommodation_WriteDB", "accommodation Write DB", "Error writing the accommodation benefit to the database.")
  Resume accommodation_WriteDB_End
  Resume
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_ACCOM
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property



