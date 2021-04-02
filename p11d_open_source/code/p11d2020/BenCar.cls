VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanyCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum CarItems
  
  car_Registration_db = ITEM_DESC 'bf
  car_carValue = ITEM_VALUE
  car_MadeGood_Net = ITEM_MADEGOOD_NET
  car_carBenefit = ITEM_BENEFIT
  car_MadeGood_db = ITEM_MADEGOOD
  car_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  car_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  car_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  car_BoxNumber = ITEM_BOX_NUMBER
  car_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  car_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  car_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  car_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  car_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  car_Error = ITEM_ERROR
  car_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  car_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  car_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE

  
  car_FuelRegistration
  car_FUEL_SPECIFIC_START = car_FuelRegistration
  car_FuelValue
  car_FuelMadeGood_NET
  car_FuelBenefit
  car_FuelMadeGood
  car_FuelActualAmountMadeGood
  car_FuelBenefit_Reportable
  car_FuelUDM_BENEFIT_TITLE
  car_FuelBoxNumber
  car_FuelMadeGoodIsTaxDeducted_db
  car_FuelClass1AAdjustment
  car_FuelNICClass1AAble
  car_FuelNIC_Class1A_Value
  car_FuelNIC_Class1A_Benefit
  car_FuelError
  car_FuelValue_Non_OPRA
  car_FuelOPRA_Ammount_Foregone_Used_For_Value
  car_FuelOPRA_Ammount_Foregone_db
  
  Car_FuelAvailableTo_db
  car_FUEL_SPECIFIC_END = Car_FuelAvailableTo_db
   
  
  car_NumberOfUsers_db
  car_FuelUnavailableDays 'used in calc only
  Car_HasFuelUnavailableDays_db
  car_fuelreinstated_db
  Car_FuelAvailableTo_calc
  car_fuelreinstated_calc
  car_FuelVATBenefit
  car_FuelVAT
  car_FuelDontUseAmountForegone
  car_Model_db 'bf
'AM To be removed  car_businessmiles 'bf ask question
  car_Make_db 'bf
  car_RegReplaced_db
  car_CarReplaced_db
  car_CarReplacedMake_db
  car_CarReplacedModel_db
  car_CarReplacedEngineSize_db
  car_ListPrice_db 'bf
  car_UnavailableDays_db
  car_capitalcontribution_db 'bf
  car_CapitalContributionRestricted
'MP DB (not used)     car_ReductionPercentageWK
  car_enginesize_db 'bf
  car_AccessoriesOriginal_db 'bf
  car_AccessoriesNew_db 'bf
  car_CheapAccessories_db 'bf
  car_carkey_db
'AM To be removed  car_SumMiles
  car_GrossPrice
  car_Price
  car_Scaled
'AM To be removed  car_BusMilereduction
'MP DB ToDo chk     car_MileScaled
'AM To be removed  car_Agereduction
  car_Accessories
'AM To be removed  car_AgeScaled
  car_GrossFuel
  car_fueldaysreduction ' EK now same as car_cardaysreduction if car_HasFuelUnavailableDays or zero
  car_cardaysreduction
  'Dates
  car_Registrationdate_db 'bf
  Car_AvailableFrom_db
  Car_AvailableTo_db 'bf test
  car_dateReplaced_db
  car_BenefitFullYear
  car_ModifiedCashEquivalent
  car_DeductionForUnavailabilityRoundUpNextWholeNumber
  car_DeductionForUnavailabilityRoundUpNextWholeNumberOpra
  car_BenefitForThePeriodTheCarWasAvailable
  car_ModifiedCashEquivalentForTheYear
  car_CapitalContributionDeductionApplicableFullYear
  car_CapitalContributionsAllowedInTheYear
  car_ProvisionalSum
  
  'Booleans
'AM To be removed  car_UseActualMiles
  car_Replacement_db
  car_Replaced_db
  car_Second_db
  car_ElectricRangeMiles_db
  car_ElectricRangeMiles_Required
'AM Car_Diesel To be removed for main 2002 release
'MP DB ToDo remove car_Diesel? chk its use in MagneticMedia.SectionF
  car_privatefuel_db 'bf
  car_requiredmakegood_db
  car_actualmadegood_db
  car_ForceP46_db
  car_FuelGrossDividedByUsers
  'p46 car print details
  car_P46PrintCarDetails
  car_P46FirstProvidedWithCar
  car_P46CarProvidedReplaced
  car_P46SecondCar
  car_P46WithdrawnWithoutReplacement
  'car_P46Withdrawn  'rh
  car_P46LowCC
  car_P46MediumCC
  car_P46HighCC
  car_P46NoCC 'so
  car_P46lowMiles
  car_P46MediumMiles
  car_P46HighMiles
  car_p46ReplacementForMultipleCarsMakeAndModel
  car_p46PaymentFrequency_db
  car_p46PaymentFrequencyString
  car_p46FuelType_db
  car_p46FuelTypeString
  car_p46FuelTypeStringLong 'RH
  car_p46CarbonDioxide_db
  car_p46NoApprovedCO2Figure_db
  car_unrestrictedPercentage 'EK
  car_OPRAModifiedCashEquivalent
  car_OPRAProvisionalSum
  
  
  'HMIT specific items
  car_HMIT_CC
  car_HMIT_petrol
  car_HMIT_diesel
  car_supplement
  car_FuelOPRAZeroField
  car_percentage
  'Lastitem
  car_lastitem = car_percentage
End Enum

Public mileage As ObjectList
Private m_ReadFromDB As Boolean
Private m_sbookmark As String
Private m_dirty As Boolean
Private Const L_DAYS_UNAVAIL_MIN As Long = 30



Private m_Parent As Employee
Private m_InvalidFields As Long
Private m_BenItems(1 To car_lastitem)
Private m_NeedToCalculate As Boolean

Public Fuel As Fuel

Private Type VAT_RATE
  From As Date
  To As Date
  Rate As Double
End Type

Private Function AdjustPercentage(dPercentageBase As Double, dPercentageAdjustment As Double) As Double
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  AdjustPercentage = dPercentageBase + dPercentageAdjustment
End Function
Private Function UseWLTPCo2Rates(ByVal ben As IBenefitClass) As Boolean
  UseWLTPCo2Rates = (ben.value(car_Registrationdate_db) >= D_CAR_USE_ELECTRIC_RANGE_MILES)
End Function

Private Function CalculateFuelPercentage(ByVal ben As IBenefitClass) As Double
  Dim dPercentageOfValue As Double
  

  If ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
    'this is purely based on the CC of the car, even if an official CO2 figure existed, special cases for electricity
    If (ben.value(car_p46FuelType_db) = CCFT_ELECTRIC) Then
      dPercentageOfValue = 0.07   'see IR480 2016
    Else
      If (ben.value(car_P46HighCC)) Then
        dPercentageOfValue = 0.37
      ElseIf ben.value(car_P46MediumCC) Then
        dPercentageOfValue = 0.27
      ElseIf ben.value(car_P46LowCC) Then
        dPercentageOfValue = 0.16
      Else 'must be 0 then we must consider as wankel engine, eg ben.value(car_P46NoCC) = true
        dPercentageOfValue = 0.37
      End If
    End If
  Else
  '2: cars registered post and including 01/01/1998
    'first get the base percentage
    If (ben.value(car_p46FuelType_db) = CCFT_ELECTRIC) Then
      dPercentageOfValue = CO2Percentage(ben)
      'dPercentageOfValue = 0#  'CAD 2010 was 0.15
    Else
      'if no approved C02 then use section 12.26
      If (ben.value(car_p46NoApprovedCO2Figure_db)) Then
        If (ben.value(car_P46HighCC)) Then
          dPercentageOfValue = 0.37
        ElseIf ben.value(car_P46MediumCC) Then
          dPercentageOfValue = 0.25
        ElseIf ben.value(car_P46LowCC) Then
          dPercentageOfValue = 0.15
        Else 'must be 0 then we must consider as wankel engine, eg ben.value(car_P46NoCC) = true
          dPercentageOfValue = 0.37
        End If
      Else
        dPercentageOfValue = CO2Percentage(ben)
      End If
    End If
    ben.value(car_unrestrictedPercentage) = dPercentageOfValue
    'apply the supplement from section 12.27 of the revenue guide
    
    
    Select Case ben.value(car_p46FuelType_db)
      Case CCFT_PETROL, CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B, CCFT_HYBRID, CCFT_E85_BIO_ENTHANOL_AND_PETROL, CCFT_BIFUEL_WITH_CO2_FOR_GAS, CCFT_GAS_ONLY, CCFT_RDE2_DIESEL
        'nothing
      Case CCFT_DIESEL, CCFT_EUROIVDIESEL
        ben.value(car_supplement) = p11d32.Rates.value(CarDIESELSUPPLEMENT)
      Case CCFT_ELECTRIC
        ben.value(car_supplement) = 0#
      Case Else
        Call Err.Raise(ERR_CAR_CO2, "Calculate", "Unknown fuel type")
    End Select

    ben.value(car_FuelDontUseAmountForegone) = CO2CalcValue(ben) <= 75
      
  End If

  dPercentageOfValue = dPercentageOfValue + ben.value(car_supplement)
  dPercentageOfValue = Max(0, Min(dPercentageOfValue, p11d32.Rates.value(carBENPERCENTMAX)))

  
  CalculateFuelPercentage = dPercentageOfValue
End Function

Private Function IBenefitClass_CalculateBody() As Variant
  'no calcualate only when required as p46 flags depend on then current dates
  Dim benFueld As IBenefitClass
  
  Dim ben As IBenefitClass
  Dim vNewLessCheapAccessories As Variant
  Dim dDateFactor As Double
  'see booklet 480 for reference for the year 2003/3004
  
On Error GoTo err_err

Call xSet("CompanyCar_Calculate")
  
  Set ben = Me
  Call ben.SetCalcDefaults
  
  Call SetP46CC
  Call SetP46FuelTypeString
  Call SetP46FrequencyPaymentString
  Call CompanyCarAvailableDatesValid(ben)
  ' ******************************** CAR CALC START ***********************************************
  If (ben.value(car_p46FuelType_db) = CCFT_HYBRID) And UseWLTPCo2Rates(ben) Then
      ben.value(car_ElectricRangeMiles_Required) = True
  End If
  
  'WS2 I/K/L, WS2B F/G/H
  ben.value(car_percentage) = CalculateFuelPercentage(ben)
  
  vNewLessCheapAccessories = ben.value(car_AccessoriesNew_db) - ben.value(car_CheapAccessories_db)
  ben.value(car_Accessories) = ben.value(car_AccessoriesOriginal_db) + IIf(vNewLessCheapAccessories > 0, vNewLessCheapAccessories, 0)
  
  'WS2/WS2B C
  ben.value(car_GrossPrice) = ben.value(car_ListPrice_db) + ben.value(car_Accessories)
  'WS2 D
  ben.value(car_CapitalContributionRestricted) = Max(Min(ben.value(car_capitalcontribution_db), p11d32.Rates.value(carMAXCAPCONTRIB)), 0)
  'WS2 E, WS2B N
  ben.value(car_Price) = Max(0, ben.value(car_GrossPrice) - ben.value(car_CapitalContributionRestricted))
  
  'WS2 M
  ben.value(car_BenefitFullYear) = ben.value(car_Price) * ben.value(car_percentage)
  'WS2B I
  ben.value(car_ModifiedCashEquivalent) = ben.value(car_GrossPrice) * ben.value(car_percentage)
    
  'WS2 N/365, WS2B J/365
  dDateFactor = dGetDateFactor(0, ben.value(Car_AvailableFrom_db), ben.value(Car_AvailableTo_db), m_BenItems(car_UnavailableDays_db), L_DAYS_UNAVAIL_MIN, False)
  
  'WS2 P
  ben.value(car_DeductionForUnavailabilityRoundUpNextWholeNumber) = RoundUp(ben.value(car_BenefitFullYear) * (1 - dDateFactor), 0)
  'WS2B K
  ben.value(car_DeductionForUnavailabilityRoundUpNextWholeNumberOpra) = RoundUp(ben.value(car_ModifiedCashEquivalent) * (1 - dDateFactor), 0)
  
  'WS2 Q
  ben.value(car_BenefitForThePeriodTheCarWasAvailable) = ben.value(car_BenefitFullYear) - ben.value(car_DeductionForUnavailabilityRoundUpNextWholeNumber)
  
  'WS2B L
  ben.value(car_ModifiedCashEquivalentForTheYear) = ben.value(car_ModifiedCashEquivalent) - ben.value(car_DeductionForUnavailabilityRoundUpNextWholeNumberOpra)
  
  
  'WS2B O
  ben.value(car_CapitalContributionDeductionApplicableFullYear) = ben.value(car_CapitalContributionRestricted) * ben.value(car_percentage)
  'WS2B P
  ben.value(car_CapitalContributionsAllowedInTheYear) = RoundDownEx(ben.value(car_CapitalContributionDeductionApplicableFullYear) * dDateFactor, 0)
  
  'WS2B Q
  ben.value(car_ProvisionalSum) = ben.value(car_OPRA_Ammount_Foregone_db) - ben.value(car_CapitalContributionsAllowedInTheYear)
  
  
  'WS2 R, WS2B R
  ben.value(car_ActualAmountMadeGood) = CalculatePaymentForPrivateUse(dDateFactor)
  
  
  '************* OPRA SPECIFIC START ***************
  
  'Section 121B ITEPA 2003:  Meaning of “modified cash equivalent”
  'ref: https://www.taxworld.co.uk/income-tax/itepa-2003/part-03-employment-income-earnings-and-benefits-etc-treated-as-earnings/chapter-06-taxable-benefits-cars-vans-and-related-benefits/section-121b-meaning-modified-cash-equivalent/
  'calcul as below ref STANDARD_CAR
  'ben.value(car_OPRAModifiedCashEquivalent) = RoundDownEx(Max(0, ben.value(car_GrossPrice)) * ben.value(car_percentage), 0) 'ignore capital contributions for modifled Opra cash equivalent
  ben.value(car_OPRAModifiedCashEquivalent) = ben.value(car_ModifiedCashEquivalentForTheYear) - ben.value(car_ActualAmountMadeGood)
  ben.value(car_OPRAModifiedCashEquivalent) = RoundDownEx(ben.value(car_OPRAModifiedCashEquivalent) / (ben.value(car_NumberOfUsers_db)), 0)
  '**** 75 CO2 thing, applies to car aswell?
  
  'Section 121A ITEPA 2003: Optional remuneration arrangements: method of calculating relevant amount
  'ref: https://www.taxworld.co.uk/income-tax/itepa-2003/part-03-employment-income-earnings-and-benefits-etc-treated-as-earnings/chapter-06-taxable-benefits-cars-vans-and-related-benefits/section-121a-optional-remuneration-arrangements-method-calculating-relevant-amount/
  ben.value(car_OPRAProvisionalSum) = ben.value(car_ProvisionalSum) '  Max(ben.value(car_OPRA_Ammount_Foregone_db) - ben.value(car_OPRACapitalContributionsTimeCarPercentage), 0)
  '*************** OPRA SPECIFIC END ***************
  
  'STANDARD_CAR START
  ben.value(car_Scaled) = RoundDownEx(ben.value(car_Price) * ben.value(car_percentage), 0)
  ben.value(ITEM_VALUE_NON_OPRA) = RoundDownEx(ben.value(car_Scaled) * dDateFactor, 0)
  ben.value(ITEM_VALUE_NON_OPRA) = RoundDownEx(ben.value(ITEM_VALUE_NON_OPRA) / (ben.value(car_NumberOfUsers_db)), 0)
  'STANDARD_CAR END
  
  ben.value(car_cardaysreduction) = ben.value(car_Scaled) - ben.value(ITEM_VALUE_NON_OPRA)  'ben.value(car_Scaled) * (1 - dDateFactor)
  
  

  Dim bOpraCalc As Boolean 'WS2B
  'Call CalculateOpRAValue(ben, ben.value(car_FuelDontUseAmountForegone),car_OPRAModifiedCashEquivalent)
  'can't use standard calc above as it's close but not quite standard
  bOpraCalc = ((ben.value(car_OPRA_Ammount_Foregone_db) > 0) And (ben.value(car_OPRA_Ammount_Foregone_db) > ben.value(car_OPRAModifiedCashEquivalent)) And (Not ben.value(car_FuelDontUseAmountForegone)))
  If bOpraCalc Then
    ben.value(ITEM_VALUE) = ben.value(car_OPRAProvisionalSum)
    ben.value(ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = True
  Else
    ben.value(ITEM_VALUE) = ben.value(ITEM_VALUE_NON_OPRA)
    ben.value(ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = False
  End If
  
  
  
  
  ben.value(car_MadeGood_Net) = Min(ben.value(car_carValue), ben.value(car_ActualAmountMadeGood))
  
  If (bOpraCalc) Then
    ben.value(car_carBenefit) = ben.value(ITEM_VALUE) - ben.value(car_MadeGood_Net)
  Else
    ben.value(car_carBenefit) = ben.value(ITEM_VALUE) - ben.value(car_MadeGood_Net)
  End If
  
  
  
  
  ' ******************************** CAR CALC END ***********************************************
  
  
  '****************************************** FUEL CALC START ********************************************
  Call CalculateFuelBen(ben.value(car_percentage))  'this is retained from original calc
    
  '****************************************** FUEL CALC END ********************************************
  Call BenCalcNIC(ben)

    
  ben.value(car_Benefit_Reportable) = True
  ben.value(car_FuelBenefit_Reportable) = True
  
  Call SetHMIT

  IBenefitClass_CalculateBody = ben.value(car_carBenefit)
   
err_end:
  Call xReturn("CompanyCar_Calculate")
  Exit Function
  
err_err:
  IBenefitClass_CalculateBody = S_ERROR
  ben.value(ITEM_ERROR) = Err.Description
  Resume err_end
  Resume

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "CompanyCar"
End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = NeedToCalculateHelper(Me, m_NeedToCalculate)
End Property

Private Property Let IBenefitClass_LinkBen(RHS As Boolean)

End Property

Private Property Get IBenefitClass_LinkBen() As Boolean

End Property

Private Function IBenefitClass_CanBringForward() As Boolean
  IBenefitClass_CanBringForward = StandardCanBringForward(Me, Car_AvailableTo_db)
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim car As CompanyCar
  
  On Error GoTo Copy_END
  Call xSet("Copy")
  
  Set car = New CompanyCar
  Set IBenefitClass_Copy = CopyBenStandard(Parent, car, Me)
  
  car.AddFuelBenefit
  
Copy_END:
  Call xReturn("Copy")
  Exit Function
Copy_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Copy", "Copy", "Error copying a company car benefit.")
  Resume Copy_ERR
End Function

Public Sub AddFuelBenefit()
  Dim ben As IBenefitClass
  Static bCreatedFuelBenefit As Boolean
'MP DB  Dim ben As IBenefitClass
  
  On Error GoTo AddFuelBenefit_ERR
  
  Call xSet("AddFuelBenefit")
  
  If Not bCreatedFuelBenefit Then
    bCreatedFuelBenefit = True
    Set Fuel = New Fuel
    Set ben = Fuel
    Set Fuel.CompanyCar = Me
    Call m_Parent.benefits.Add(Fuel)
  End If

AddFuelBenefit_END:
  Call xReturn("AddFuelBenefit")
  Exit Sub
AddFuelBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddFuelBenefit", "Add Non Qualifying Relocation", "Error adding a fuel benefit with a company car benefit.")
  Resume AddFuelBenefit_END
  Resume
End Sub

Private Sub Class_Initialize()
  Set mileage = New ObjectList
  Call IBenefitClass_SetBenItemsInformation

End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_COMPANY_CARS_F
End Property

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean

End Property
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property
Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(car_Registration_db)
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
Private Property Get IBenefitClass_Print_HMIT() As String

End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass
  Dim lRunningPercentage As Long
  Dim lFinalPercentage As Long
  Dim lFuelDeduction As Long
  Dim lAdditionalDiscount As Long
  Dim lTotalUsers As Long
  Dim bSubtotal As Boolean
  
  On Error GoTo CAR_PRINTWK_Err
  Call xSet("CAR_PRINTWK")
  
  Set ben = Me
  bSubtotal = False
  lAdditionalDiscount = 0
  lFuelDeduction = 0
  lRunningPercentage = 0
  If ben.value(car_Price) = 0 Then
    lFinalPercentage = 0
  Else
    lFinalPercentage = (ben.value(car_percentage) * 100)
    ' lFinalPercentage = (ben.value(car_Scaled) / ben.value(car_Price)) * 100 'ek change for co2 working paper
  End If
  lTotalUsers = ben.value(car_NumberOfUsers_db)
  
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Make and model: ", ben.value(car_Make_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Made available from " & S_WK_NORMAL_BOLD_FONT & ben.value(Car_AvailableFrom_db) & S_WK_NORMAL_FONT & " to " & S_WK_NORMAL_BOLD_FONT & ben.value(Car_AvailableTo_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Qualifying days unavailable: " & S_WK_NORMAL_BOLD_FONT & ben.value(car_UnavailableDays_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Number of employees using this car: " & S_WK_NORMAL_BOLD_FONT & lTotalUsers)

'EK Changes to working paper 2003
  
  Call WKOut(rep, WK_BLANK_LINE)
  If ben.value(car_p46NoApprovedCO2Figure_db) Then
    Call WKOut(rep, WK_ITEM_TEXT, "No approved CO2 figure: " & S_WK_NORMAL_BOLD_FONT & ben.value(car_p46NoApprovedCO2Figure_db))
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "CO2 emissions figure: " & S_WK_NORMAL_BOLD_FONT & CO2CalcValue(ben) & S_WK_NORMAL_FONT & " g/km")
  End If
  
  Call WKOut(rep, WK_ITEM_TEXT, "Date of registration: " & S_WK_NORMAL_BOLD_FONT & ben.value(car_Registrationdate_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Engine size: " & S_WK_NORMAL_BOLD_FONT & ben.value(car_enginesize_db) & S_WK_NORMAL_FONT & " cc")
  Call WKOut(rep, WK_ITEM_TEXT, "Type of fuel: " & S_WK_NORMAL_BOLD_FONT & ben.value(car_p46FuelTypeStringLong))
      
      
'AM To be removed  If ben.value(car_UseActualMiles) Then Call WKOut(rep, WK_ITEM_TEXT, "Total mileage " & ben.value(car_SumMiles))
' Price of car for tax purposes
  Call WKOut(rep, WK_SECTION_HEADER_VALUE)
  Call WKOut(rep, WK_ITEM_TEXT, "Car list price ", ben.value(car_ListPrice_db), , True, False)
  Call WKOut(rep, WK_ITEM_TEXT, "Optional accessories available when the car was first made available", ben.value(car_AccessoriesOriginal_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Accessories added after car was first made available", ben.value(car_AccessoriesNew_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Accessories worth less than £100", ben.value(car_CheapAccessories_db), , , True)
  
  Call WKOut(rep, WK_ITEM_Total, S_WK_LEFT_MARGIN & "Total car list price plus accessories ", ben.value(car_GrossPrice), , True)
  ' Less reductions
  Call WKOut(rep, WK_SECTION_HEADER_LESS)
  Call WKOut(rep, WK_ITEM_TEXT, "Capital contribution by the employee towards the cost of the car " & IIf(ben.value(car_capitalcontribution_db) > p11d32.Rates.value(carMAXCAPCONTRIB), "(" & FormatWN(ben.value(car_capitalcontribution_db)) & " restricted to " & FormatWN(p11d32.Rates.value(carMAXCAPCONTRIB)) & ")", ""), Min(ben.value(car_capitalcontribution_db), p11d32.Rates.value(carMAXCAPCONTRIB)), , True, True)
  
  'Call WKOut(rep, WK_ITEM_Total, "Adjusted price of the car for " & p11d32.Rates.value(TaxFormYear) & " (restricted to " & FormatWN(p11d32.Rates.value(carMAXLISTPRICE)) & ")", ben.value(car_Price), , True)
  Call WKOut(rep, WK_ITEM_Total, "Adjusted price of the car for " & p11d32.Rates.value(TaxFormYear), ben.value(car_Price), , True)
    
'AM To be removed  Call WKOut(rep, WK_ITEM_TEXT, "Gross benefit for " & p11d32.Rates.value(TaxFormYear) & " (Adjusted price x " & ben.value(car_ReductionPercentageWK) & "%)", ben.value(car_Scaled) - ben.value(car_BusMilereduction))
'AM To be removed  Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "Note: " & ben.value(car_ReductionPercentageWK) & "% reduction as car in " & ben.value(car_businessmiles) & " band.")
  Call WKOut(rep, WK_BLANK_LINE)
  
'AM To be removed  Call WKOut(rep, WK_ITEM_TEXT, "Reduction for age of car (first registered " & ben.value(car_Registrationdate) & ")", ben.value(car_Agereduction), , , True)
  Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)

' EK Added outline

  Dim dReduction As Double
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_BOLD_FONT & "Calculate appropriate percentage", "")
  
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "Base calcuation percentage ", S_WKCOL_3 & ReportPercentage(ben.value(car_unrestrictedPercentage)))
  
  
  If ben.value(car_supplement) > 0 Then
    Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "Supplement according to fuel type ", S_WKCOL_3 & ReportPercentage(ben.value(car_supplement)))
  Else
    Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "Reduction according to fuel type ", S_WKCOL_3 & ReportPercentage(ben.value(car_supplement)))
  End If
  
  Dim MAX_RATE_PERCENTAGE_STRING As String
  MAX_RATE_PERCENTAGE_STRING = RoundDownEx(p11d32.Rates.value(carBENPERCENTMAX) * 100, 0)
  
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "Final calcuation percentage restricted to " & MAX_RATE_PERCENTAGE_STRING & "%", S_WKCOL_3 & ReportPercentage(ben.value(car_percentage)))
  
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKOut(rep, WK_ITEM_TEXT, "Adjusted car list price multiplied by percentage ", RoundDownEx(ben.value(car_Price) * lFinalPercentage / 100, 0))
  Call WKOut(rep, WK_ITEM_TEXT, "Reduction for days when the car was unavailable", ben.value(car_cardaysreduction), , , True)
  
 
  If (ben.value(car_OPRA_Ammount_Foregone_db) > 0) Then
    Call WKOut(rep, WK_BLANK_LINE)
    Call WKOut(rep, WK_ITEM_TEXT, "OpRA amount foregone", ben.value(car_OPRA_Ammount_Foregone_db))
    Call WKOut(rep, WK_ITEM_TEXT, "OpRA modified cash equivalent (ITEPA 2003 - Section 121B)", ben.value(car_OPRAModifiedCashEquivalent))
    Call WKOut(rep, WK_ITEM_TEXT, "OpRA provisional sum (ITEPA 2003 - Section 121A)", ben.value(car_OPRAProvisionalSum))
    Call WKOut(rep, WK_BLANK_LINE)
    If (ben.value(car_OPRA_Ammount_Foregone_Used_For_Value)) Then
      Call WKOut(rep, WK_ITEM_TEXT, "'OpRA amount foregone' > 'OpRA modified cash equivalent', using the 'OpRA provisional sum' as the value (ITEPA 2003 - Section 120A)", ben.value(car_OPRAProvisionalSum))
    Else
      Call WKOut(rep, WK_ITEM_TEXT, "'OpRA amount foregone' <= to the 'OpRA modified cash equivalent', using the 'standard calculation' for the value (ITEPA 2003 - Section 120A)", ben.value(ITEM_VALUE_NON_OPRA))
    End If
  
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "Value of benefit", ben.value(ITEM_VALUE))
  
  End If
  Call WKOut(rep, WK_BLANK_LINE)
  
  
  'Call WKOut(rep, WK_ITEM_Total, IIf(lTotalUsers > 1, "Assessable car benefit per employee", "Assessable car benefit"), ben.value(car_carBenefit), , True)
  
  
  
  Call WKOut(rep, WK_ITEM_TEXT, "Payment from the employee for private use of the car", ben.value(car_ActualAmountMadeGood), , , True)
  
  Call WKOut(rep, WK_ITEM_Total, "Assessable car benefit", ben.value(car_carBenefit), , True)
  
  Call WKOut(rep, WK_SECTION_HEADER, "Fuel Benefit")
  
  If ben.value(Car_HasFuelUnavailableDays_db) Then
    Call WKOut(rep, WK_ITEM_TEXT, "Qualifying days unavailable: " & S_WK_NORMAL_BOLD_FONT & Min(ben.value(car_UnavailableDays_db), ben.value(car_UnavailableDays_db)))
  End If
  
  Call WKOut(rep, WK_ITEM_TEXT, "Was fuel provided for private use?", WKBool(ben, car_privatefuel_db))
  
  'if fuel benefit > 0 then display the date withdrawn if present - EK altered 2/04 TTP#212
  If ben.value(car_privatefuel_db) And (ben.value(Car_AvailableTo_db) > ben.value(Car_FuelAvailableTo_db)) Then
    'WILL THIS PRINT OUT ok in US LOCALE?
    Call WKOut(rep, WK_ITEM_TEXT, "Date car fuel withdrawn: " & S_WK_NORMAL_BOLD_FONT & ben.value(Car_FuelAvailableTo_db))
    Call WKOut(rep, WK_ITEM_TEXT, "Was fuel reinstated?", WKBool(ben, car_fuelreinstated_db))
  End If
  
  
  Call WKOut(rep, WK_ITEM_TEXT, "Was the employee required to make good the cost of all fuel so provided?", WKBool(ben, car_requiredmakegood_db))
  Call WKOut(rep, WK_ITEM_TEXT, "If so required, did the employee do so?", WKBool(ben, car_actualmadegood_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Number of users?", lTotalUsers)
  Call WKOut(rep, WK_ITEM_TEXT, "Fuel benefit (percentage multiplied by £" & FormatNumber(p11d32.Rates.value(CarFUELBENSETFIGURE), 0, , , vbTrue) & "):", (ben.value(car_GrossFuel)), , True)
  Call WKOut(rep, WK_ITEM_TEXT, "Divide by number of users '" & lTotalUsers & "'?", ben.value(car_FuelGrossDividedByUsers))
  
  'If Not (ben.value(car_FuelOPRA_Ammount_Foregone_Used_For_Value)) Then
  Call WKOut(rep, WK_ITEM_TEXT, "Reduction for days when car was unavailable", ben.value(car_fueldaysreduction), , , True)
  'End If
  
  Call WKOut(rep, WK_BLANK_LINE)
  '
  
  Call OPRAWorkingPaperValue(Fuel, rep)
  
  Call WKOut(rep, WK_ITEM_Total, IIf(lTotalUsers > 1, "Assessable fuel benefit per employee", "Assessable fuel benefit"), ben.value(car_FuelBenefit), , True)
CAR_PRINTWK_End:
  Call xReturn("CAR_PRINTWK")
  Exit Function

CAR_PRINTWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CAR_PRINTWK", "CAR PRINT WK", "Error printing the company car working paper.")
  Resume CAR_PRINTWK_End
  Resume
  
End Function
Private Function WKBool(ben As IBenefitClass, Item As Long) As String
  If p11d32.BenDataLinkDataType(ben.BenefitClass, Item) = TYPE_BOOL Then
    WKBool = IIf(ben.value(Item), "Yes", "No")
  Else
    ECASE ("Non boolean value passed to WKBool")
  End If
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
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
  bc = BC_COMPANY_CARS_F
  m_NeedToCalculate = True
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
     
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  Call SetStandardBenItemsUDMData(bc, " - Car")

  With p11d32
    Set ben = Me
    'data types first
    'data types
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    .BenDataLinkDataType(bc, car_Model_db) = TYPE_STR
    .BenDataLinkDataType(bc, car_Make_db) = TYPE_STR
    .BenDataLinkMMRequired(bc, car_Make_db) = True
'AM To be removed    .BenDataLinkDataType(bc, car_businessmiles) = TYPE_STR
    .BenDataLinkDataType(bc, car_Registration_db) = TYPE_STR
    .BenDataLinkDataType(bc, car_RegReplaced_db) = TYPE_STR
    .BenDataLinkDataType(bc, car_CarReplaced_db) = TYPE_STR
    .BenDataLinkDataType(bc, car_CarReplacedMake_db) = TYPE_STR
    .BenDataLinkDataType(bc, car_CarReplacedModel_db) = TYPE_STR
    .BenDataLinkDataType(bc, car_CarReplacedEngineSize_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelMadeGoodIsTaxDeducted_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, ITEM_BOX_NUMBER) = TYPE_STR
    .BenDataLinkDataType(bc, car_FuelBoxNumber) = TYPE_STR
    .BenDataLinkDataType(bc, car_FuelActualAmountMadeGood) = TYPE_LONG
    
    
    .BenDataLinkDataType(bc, Car_FuelAvailableTo_calc) = TYPE_DATE
    .BenDataLinkDataType(bc, Car_FuelAvailableTo_db) = TYPE_DATE
    .BenDataLinkDataType(bc, car_FuelBenefit) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelBenefit_Reportable) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_FuelClass1AAdjustment) = TYPE_LONG
    .BenDataLinkDataType(bc, car_fueldaysreduction) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelError) = TYPE_STR
    .BenDataLinkDataType(bc, car_FuelMadeGood) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelMadeGood_NET) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelNIC_Class1A_Value) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelNIC_Class1A_Benefit) = TYPE_DOUBLE
    .BenDataLinkDataType(bc, car_FuelUDM_BENEFIT_TITLE) = TYPE_STR
    .BenDataLinkDataType(bc, car_FuelUnavailableDays) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelNICClass1AAble) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_FuelOPRA_Ammount_Foregone_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelOPRA_Ammount_Foregone_Used_For_Value) = TYPE_BOOL
    
    
    
    
'MP DB    .BenDataLinkDataType(bc, car_ReductionPercentageWK) = TYPE_LONG
    .BenDataLinkDataType(bc, car_ListPrice_db) = TYPE_LONG
    .BenDataLinkMMRequired(bc, car_ListPrice_db) = True
    .BenDataLinkDataType(bc, car_UnavailableDays_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_capitalcontribution_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_enginesize_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_AccessoriesOriginal_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_AccessoriesNew_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_CheapAccessories_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_carkey_db) = TYPE_LONG
'AM To be removed    .BenDataLinkDataType(bc, car_SumMiles) = TYPE_LONG
    .BenDataLinkDataType(bc, car_GrossPrice) = TYPE_LONG
    .BenDataLinkDataType(bc, car_Price) = TYPE_LONG
    .BenDataLinkDataType(bc, car_Scaled) = TYPE_LONG
'AM To be removed    .BenDataLinkDataType(bc, car_BusMilereduction) = TYPE_LONG
'MP DB    .BenDataLinkDataType(bc, car_MileScaled) = TYPE_LONG
'AM To be removed    .BenDataLinkDataType(bc, car_Agereduction) = TYPE_LONG
    .BenDataLinkDataType(bc, car_Accessories) = TYPE_LONG
'AM    .BenDataLinkDataType(bc, car_AgeScaled) = TYPE_LONG
    .BenDataLinkDataType(bc, car_ActualAmountMadeGood) = TYPE_LONG 'so
    .BenDataLinkDataType(bc, car_p46PaymentFrequency_db) = TYPE_LONG 'so
    .BenDataLinkDataType(bc, car_p46PaymentFrequencyString) = TYPE_STR 'so
    .BenDataLinkDataType(bc, car_p46FuelType_db) = TYPE_LONG 'so
    .BenDataLinkDataType(bc, car_p46FuelTypeString) = TYPE_STR 'so
    .BenDataLinkMMFieldSize(bc, car_p46FuelTypeString) = 1  'so
    
    
    .BenDataLinkDataType(bc, car_p46FuelTypeStringLong) = TYPE_STR 'RH
    .BenDataLinkDataType(bc, car_p46CarbonDioxide_db) = TYPE_LONG 'so
    .BenDataLinkDataType(bc, car_p46NoApprovedCO2Figure_db) = TYPE_BOOL 'RH
       
    .BenDataLinkDataType(bc, car_fueldaysreduction) = TYPE_LONG
    .BenDataLinkDataType(bc, car_cardaysreduction) = TYPE_LONG
    
    .BenDataLinkDataType(bc, car_Registrationdate_db) = TYPE_DATE
    .BenDataLinkDataType(bc, Car_AvailableFrom_db) = TYPE_DATE
    .BenDataLinkDataType(bc, Car_AvailableTo_db) = TYPE_DATE
    .BenDataLinkDataType(bc, car_dateReplaced_db) = TYPE_DATE
    
'AM    .BenDataLinkDataType(bc, car_UseActualMiles) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_Replacement_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_Replaced_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_Second_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_ElectricRangeMiles_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_ElectricRangeMiles_Required) = TYPE_BOOL
    
    .BenDataLinkDataType(bc, car_privatefuel_db) = TYPE_BOOL
'AM To be removed    .BenDataLinkDataType(bc, car_Diesel) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_requiredmakegood_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_actualmadegood_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_ForceP46_db) = TYPE_BOOL
    
    'p46 car print details
    .BenDataLinkDataType(bc, car_P46PrintCarDetails) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46FirstProvidedWithCar) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46CarProvidedReplaced) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46SecondCar) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46WithdrawnWithoutReplacement) = TYPE_BOOL
    '.BenDataLinkDataType(bc, car_P46Withdrawn) = TYPE_BOOL  'rh
    .BenDataLinkDataType(bc, car_P46LowCC) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46MediumCC) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46HighCC) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46NoCC) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46lowMiles) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46MediumMiles) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_P46HighMiles) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_p46ReplacementForMultipleCarsMakeAndModel) = TYPE_BOOL
    
    'HMIT specific items
    .BenDataLinkDataType(bc, car_HMIT_petrol) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_HMIT_diesel) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_CapitalContributionRestricted) = TYPE_LONG
    
    
    .BenDataLinkDataType(bc, car_HMIT_CC) = TYPE_STR

    'Magnetic media specifics
    'CAD do fuel
'MP DB    .BenDataLinkDataType(bc, car_FuelRegistration) = TYPE_STR
    .BenDataLinkDataType(bc, car_FuelValue) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelMadeGood) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelBenefit) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelMadeGood_NET) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelBenefit_Reportable) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_FuelVAT) = TYPE_DOUBLE
    .BenDataLinkDataType(bc, car_FuelVATBenefit) = TYPE_LONG
    .BenDataLinkDataType(bc, Car_FuelAvailableTo_db) = TYPE_DATE
    .BenDataLinkDataType(bc, car_fuelreinstated_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_NumberOfUsers_db) = TYPE_LONG
    .BenDataLinkDataType(bc, car_FuelUnavailableDays) = TYPE_LONG
    .BenDataLinkDataType(bc, Car_HasFuelUnavailableDays_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, car_FuelNIC_Class1A_Benefit) = TYPE_DOUBLE
    
    .BenDataLinkDataType(bc, car_OPRAModifiedCashEquivalent) = TYPE_LONG
    .BenDataLinkDataType(bc, car_OPRAProvisionalSum) = TYPE_LONG
    
    
    
    
    'MMField sizes
    .BenDataLinkUDMDescription(bc, car_Make_db) = "Make"
    .BenDataLinkUDMDescription(bc, car_Model_db) = "Model"
    .BenDataLinkUDMDescription(bc, car_ListPrice_db) = "List Price"
    .BenDataLinkUDMDescription(bc, car_NumberOfUsers_db) = "Number of users"
    .BenDataLinkUDMDescription(bc, car_Accessories) = "Accessories"
    .BenDataLinkUDMDescription(bc, Car_AvailableFrom_db) = S_UDM_FROM
    .BenDataLinkUDMDescription(bc, Car_AvailableTo_db) = S_UDM_To
    
    .BenDataLinkUDMDescription(bc, car_UnavailableDays_db) = "Unavailable"
    
    .BenDataLinkUDMDescription(bc, car_capitalcontribution_db) = "Capital Contrib"
    .BenDataLinkUDMDescription(bc, car_Registrationdate_db) = "Reg Date"
    'AM To be removed
    '.BenDataLinkUDMDescription(bc, car_UseActualMiles) = "Use actual miles"
    '.BenDataLinkUDMDescription(bc, car_SumMiles) = "Actual miles"
    '.BenDataLinkUDMDescription(bc, car_businessmiles) = "Business miles class"
    
    .BenDataLinkUDMDescription(bc, car_FuelBenefit) = "Fuel benefit"
    .BenDataLinkUDMDescription(bc, car_FuelNIC_Class1A_Benefit) = S_UDM_NIC_CLASS1A_BENEFIT & " - Fuel"
    ' .BenDataLinkUDMDescription(bc, car_fuelavailablefrom) = "Fuel available from"
    .BenDataLinkUDMDescription(bc, car_FuelVAT) = "VAT on fuel"
    .BenDataLinkUDMDescription(bc, car_FuelVATBenefit) = "VAT on fuel benefit"
    .BenDataLinkUDMDescription(bc, car_enginesize_db) = "cc"
    'km only display "Diesel?" in 2000 version as the "Is car Diesel?" chkbx is not used in 2001
'    If p11d32.AppYear = 2000 Then .BenDataLinkUDMDescription(bc, car_Diesel) = "Diesel?"
    .BenDataLinkUDMDescription(bc, car_Second_db) = "Second car?"
    .BenDataLinkUDMDescription(bc, car_ElectricRangeMiles_db) = S_COMPANY_CAR_ELECTFRIC_RANGE_MILES_DESCRIPTION
    .BenDataLinkUDMDescription(bc, car_Replaced_db) = "Replaced?"
    .BenDataLinkUDMDescription(bc, car_Replacement_db) = "Replacement?"
    
    .BenDataLinkUDMDescription(bc, car_privatefuel_db) = "Private fuel?"
    .BenDataLinkUDMDescription(bc, car_requiredmakegood_db) = "Private fuel required to make good?"
    .BenDataLinkUDMDescription(bc, car_actualmadegood_db) = "Private fuel made good?"
    .BenDataLinkUDMDescription(bc, Car_FuelAvailableTo_db) = S_FIELD_CAR_FUEL_WITHDRAWN_DATE
    
    'ek 2/04 TTP#194
    .BenDataLinkUDMDescription(bc, car_fuelreinstated_db) = "Fuel reinstated"
    
    .BenDataLinkDataType(bc, car_fuelreinstated_calc) = TYPE_BOOL
    .BenDataLinkDataType(bc, Car_FuelAvailableTo_calc) = TYPE_DATE
    .BenDataLinkMMFieldSize(bc, car_percentage) = TYPE_DOUBLE
    .BenDataLinkMMFieldSize(bc, car_unrestrictedPercentage) = TYPE_DOUBLE
    .BenDataLinkMMFieldSize(bc, car_supplement) = TYPE_DOUBLE
        
    
    
          
        
    .BenDataLinkUDMDescription(bc, car_p46PaymentFrequencyString) = "Payment frequency"
    .BenDataLinkUDMDescription(bc, car_p46FuelTypeStringLong) = "Type of fuel" 'RH
    .BenDataLinkUDMDescription(bc, car_p46CarbonDioxide_db) = "Grams of Carbon dioxide"
    .BenDataLinkUDMDescription(bc, car_p46NoApprovedCO2Figure_db) = "No CO2 Figure available" 'RH
    
    .BenDataLinkUDMDescription(bc, Car_HasFuelUnavailableDays_db) = S_COMPANY_CAR_DAYS_UNAVAILABLE_FUEL_DESCRIPTION
    .BenDataLinkUDMDescription(bc, car_OPRAProvisionalSum) = "OpRA provisional sum"
    .BenDataLinkUDMDescription(bc, car_OPRAModifiedCashEquivalent) = "OpRA modified cash equivalent"
    
    
    .BenDataLinkMMFieldSize(bc, car_Second_db) = 2
    .BenDataLinkMMFieldSize(bc, car_Make_db) = 35
    .BenDataLinkMMFieldSize(bc, car_Model_db) = 35
    .BenDataLinkMMFieldSize(bc, car_Registrationdate_db) = MM_SFS_DATE
    .BenDataLinkMMFieldSize(bc, car_enginesize_db) = 4
    .BenDataLinkMMFieldSize(bc, car_Second_db) = MM_SFS_bool
    .BenDataLinkMMFieldSize(bc, Car_AvailableFrom_db) = MM_SFS_DATE
    .BenDataLinkMMFieldSize(bc, Car_AvailableTo_db) = MM_SFS_DATE
    .BenDataLinkMMFieldSize(bc, car_ListPrice_db) = MM_SFS_DATA
    .BenDataLinkMMFieldSize(bc, car_AccessoriesOriginal_db) = MM_SFS_DATA
    .BenDataLinkMMFieldSize(bc, car_AccessoriesNew_db) = MM_SFS_DATA
    .BenDataLinkMMFieldSize(bc, car_capitalcontribution_db) = MM_SFS_DATA
    .BenDataLinkMMFieldSize(bc, car_p46CarbonDioxide_db) = 3 'ek
    .BenDataLinkMMFieldSize(bc, car_p46NoApprovedCO2Figure_db) = MM_SFS_bool
    .BenDataLinkMMFieldSize(bc, Car_FuelAvailableTo_db) = MM_SFS_DATE
    .BenDataLinkMMFieldSize(bc, car_fuelreinstated_db) = MM_SFS_bool
    
    .BenDataLinkMMFieldSize(bc, Car_HasFuelUnavailableDays_db) = MM_SFS_bool
    
    
    
  End With

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub




Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  Call SetCalcDefaultsStandard(ben)
  
  'HMIT items
  ben.value(car_HMIT_CC) = S_ERROR
  ben.value(car_HMIT_petrol) = False
  ben.value(car_HMIT_diesel) = False
  'Set all fields to error
  ben.value(car_GrossPrice) = S_ERROR
  ben.value(car_Price) = S_ERROR
  ben.value(car_Scaled) = S_ERROR
  ben.value(car_Accessories) = S_ERROR
  ben.value(car_cardaysreduction) = 0
  ben.value(car_carValue) = S_ERROR
  ben.value(car_GrossFuel) = S_ERROR
  ben.value(car_fueldaysreduction) = S_ERROR
  ben.value(car_FuelValue) = S_ERROR
  ben.value(car_carBenefit) = S_ERROR
  ben.value(car_FuelBenefit) = S_ERROR
  ben.value(car_FuelValue) = S_ERROR
  ben.value(car_FuelMadeGood) = S_ERROR
  ben.value(car_FuelMadeGood_NET) = S_ERROR
  ben.value(car_FuelBenefit_Reportable) = False
  ben.value(car_FuelError) = ""
  ben.value(car_FuelClass1AAdjustment) = 0
  ben.value(car_unrestrictedPercentage) = 0
  ben.value(car_percentage) = 0
  ben.value(car_supplement) = 0
  ben.value(car_OPRA_Ammount_Foregone_Used_For_Value) = False
  ben.value(car_FuelOPRA_Ammount_Foregone_Used_For_Value) = False
  ben.value(car_FuelMadeGood) = 0
  
  ben.value(car_FuelVATBenefit) = 0
  ben.value(car_GrossFuel) = 0
  ben.value(car_FuelDontUseAmountForegone) = False
  ben.value(car_OPRAModifiedCashEquivalent) = 0
  ben.value(car_OPRAProvisionalSum) = 0
  ben.value(car_ElectricRangeMiles_Required) = False
  
  
  
        '.value(car_GrossFuel) = 0
        '.value(car_FuelVATBenefit) = 0
  
  
End Sub
Private Sub IBenefitClass_Kill()
  Set Fuel = Nothing
  Set m_Parent = Nothing
End Sub



Private Function CO2Percentage(ByVal ben As IBenefitClass) As Double
  Const MAX_CO2 As Long = 170
  
  Dim Rate As Long
  Dim fuelType As COMPANY_CAR_FUEL_TYPE
  Dim co2 As Long
  Dim rateDouble As Double
  Dim electricRangeMiles As Long
  
  fuelType = ben.value(car_p46FuelType_db)
  electricRangeMiles = ben.value(car_ElectricRangeMiles_db)
  co2 = CO2CalcValue(ben)
  
  'NEDC rates
  co2 = Min(co2, MAX_CO2)
  If (co2 >= 51 <= MAX_CO2) Then
    Select Case co2
      Case Is >= 170
        Rate = 37
      Case Is >= 165
        Rate = 37
      Case Is >= 160
        Rate = 37
      Case Is >= 155
        Rate = 36
      Case Is >= 150
        Rate = 35
      Case Is >= 145
        Rate = 34
      Case Is >= 140
        Rate = 33
      Case Is >= 135
        Rate = 32
      Case Is >= 130
        Rate = 31
      Case Is >= 125
        Rate = 30
      Case Is >= 120
        Rate = 29
      Case Is >= 115
        Rate = 28
      Case Is >= 110
        Rate = 27
      Case Is >= 105
        Rate = 26
      Case Is >= 100
        Rate = 25
      Case Is >= 95
        Rate = 24
      Case Is >= 90
        Rate = 23
      Case Is >= 85
        Rate = 22
      Case Is >= 80
        Rate = 21
      Case Is >= 75
        Rate = 20
      Case Is >= 70
        Rate = 19
      Case Is >= 65
        Rate = 18
      Case Is >= 60
        Rate = 17
      Case Is >= 55
        Rate = 16
      Case Is >= 51
        Rate = 15
      Case Is < 51
        Select Case fuelType
          Case CCFT_ELECTRIC
            Rate = 0
          Case CCFT_HYBRID
            Select Case (electricRangeMiles)
              Case Is >= 130
                Rate = 2
              Case Is >= 70
                Rate = 5
              Case Is >= 40
                Rate = 8
              Case Is >= 30
                Rate = 12
              Case Is < 30
                Rate = 14
            End Select
          Case Else
           'rate unknown, this is not hybid or electric vehicles < 51 assume 14
            Rate = 14
        End Select
    End Select
  End If
  
  'adjust to WLTP rates
  If (UseWLTPCo2Rates(ben)) Then
    If (fuelType <> CCFT_ELECTRIC) And co2 Then
      Select Case co2
        Case Is >= 170
        Case Is >= 165
          Rate = Rate - 1
        Case Else
          Rate = Rate - 2
      End Select
    End If
  End If
  
  
  rateDouble = CDbl(Rate) / CDbl(100)
  CO2Percentage = rateDouble
End Function

Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Private Function CO2CalcValue(ben As IBenefitClass) As Long
  If (ben.value(car_p46FuelType_db) = COMPANY_CAR_FUEL_TYPE.CCFT_ELECTRIC) Then
    CO2CalcValue = 0
  Else
    CO2CalcValue = ben.value(car_p46CarbonDioxide_db)
  End If
End Function


Private Function FuelVATAnnualScaleCharge(ben As IBenefitClass)
  Dim iCO2 As Long, iFuelVATScaleAnnual As Long
  Dim bNoCO2Type As Boolean
  
  iCO2 = CO2CalcValue(ben)
  
  Select Case iCO2
    Case Is < 125
      iFuelVATScaleAnnual = 581
    Case Is < 130
      iFuelVATScaleAnnual = 870
    Case Is < 135
      iFuelVATScaleAnnual = 930
    Case Is < 140
      iFuelVATScaleAnnual = 986
    Case Is < 145
      iFuelVATScaleAnnual = 1047
    Case Is < 150
      iFuelVATScaleAnnual = 1103
    Case Is < 155
      iFuelVATScaleAnnual = 1163
    Case Is < 160
      iFuelVATScaleAnnual = 1219
    Case Is < 165
      iFuelVATScaleAnnual = 1279
    Case Is < 170
      iFuelVATScaleAnnual = 1335
    Case Is < 175
      iFuelVATScaleAnnual = 1396
    Case Is < 180
      iFuelVATScaleAnnual = 1452
    Case Is < 185
      iFuelVATScaleAnnual = 1512
    Case Is < 190
      iFuelVATScaleAnnual = 1568
    Case Is < 195
      iFuelVATScaleAnnual = 1628
    Case Is < 200
      iFuelVATScaleAnnual = 1684
    Case Is < 205
      iFuelVATScaleAnnual = 1745
    Case Is < 210
      iFuelVATScaleAnnual = 1801
    Case Is < 215
      iFuelVATScaleAnnual = 1861
    Case Is < 220
      iFuelVATScaleAnnual = 1917
    Case Is < 225
      iFuelVATScaleAnnual = 1977
    Case Is >= 225
      iFuelVATScaleAnnual = 2033
    Case Else
      Call Err.Raise(ERR_INVALIDID, "FuelVATAnnualScaleCharge", "Invalid CO2 value" & iCO2)
  End Select

  FuelVATAnnualScaleCharge = iFuelVATScaleAnnual
End Function
Private Sub CalculateFuelBen(ByVal BenRate As Double)
  Dim ben As IBenefitClass
  Dim dFuelDateFactor As Variant
  Dim l As Long
  Dim dFuelDateTo As Date, dFuelDateFrom As Date
  Dim iCO2 As Long, iBase As Long, iFuelVATScaleAnnual As Long
  Dim vatRates(1 To 4) As VAT_RATE
  Dim iDaysAtRate As Long
  Dim iDaysAvailable As Long
  Dim dVAT As Double
  
  On Error GoTo CalculateFuelBen_Err
  Call xSet("CalculateFuelBen")
  
  Set ben = Me
  
  With ben
  
  .value(car_FuelUnavailableDays) = IIf(m_BenItems(Car_HasFuelUnavailableDays_db), m_BenItems(car_UnavailableDays_db), 0)
  'fuel date factor
  
    dFuelDateFrom = .value(Car_AvailableFrom_db)
    dFuelDateTo = .value(Car_AvailableTo_db)
    If .value(Car_FuelAvailableTo_db) < dFuelDateTo Then
      If Not .value(car_fuelreinstated_db) Then
        dFuelDateTo = .value(Car_FuelAvailableTo_db)
      End If
    End If
    dFuelDateFactor = dGetDateFactor(l, dFuelDateFrom, dFuelDateTo, .value(car_FuelUnavailableDays), L_DAYS_UNAVAIL_MIN, False, iDaysAvailable)
    
    'calculating fuel VAT benefit according to engine size and petrol type
    If .value(car_privatefuel_db) Then
      'N.B. electric are zeroed below
      If (p11d32.AppYear >= 2007) Then
        .value(car_FuelVATBenefit) = FuelVATAnnualScaleCharge(ben)
      End If

      'Calculating Fuel benefit (=  set figure (£14,400 for 2003) * benefit rate)
      .value(car_GrossFuel) = p11d32.Rates.value(CarFUELBENSETFIGURE) * BenRate  'CAD 2010 changed from 16900 to 18000
      
      'RH - No Fuel Benefit if Electricity only
      If ((.value(car_requiredmakegood_db) And .value(car_actualmadegood_db)) Or (.value(car_p46FuelType_db) = CCFT_ELECTRIC)) Then
        'electric - No taxable benefit - electricity does not sit within the meaning of fuel so the Fuel Benefit Charge does not apply. And there is no further benefit charge as s239(4) ITEPA 2003 specifically excludes a benefit connected with a taxable car.
        'currently only electrc cars have zero emissions, so no need for an additional check box on the form

        .value(car_FuelMadeGood) = (.value(car_GrossFuel) * dFuelDateFactor) / .value(car_NumberOfUsers_db)
        '.value(car_GrossFuel) = 0
        '.value(car_FuelVATBenefit) = 0
      Else
        .value(car_FuelMadeGood) = 0
      End If
    Else
      .value(car_FuelVATBenefit) = 0
      .value(car_FuelMadeGood) = 0
      .value(car_GrossFuel) = 0
    End If
        
    'cad company car change
    .value(Car_FuelAvailableTo_calc) = UNDATED
    .value(car_fuelreinstated_calc) = False
    If (.value(car_p46FuelType_db) = CCFT_ELECTRIC) Or (Not .value(car_privatefuel_db)) Then
      'leave false
    Else
      If .value(Car_FuelAvailableTo_db) >= .value(Car_AvailableTo_db) Then
        'leave false
      ElseIf .value(Car_FuelAvailableTo_db) < .value(Car_AvailableFrom_db) Then
        'leave false
      Else
        .value(Car_FuelAvailableTo_calc) = .value(Car_FuelAvailableTo_db)
        If .value(car_fuelreinstated_db) Then
          .value(car_fuelreinstated_calc) = .value(car_fuelreinstated_db)
        End If
      End If
    End If
      
      
    .value(car_fueldaysreduction) = RoundUp(.value(car_GrossFuel) * (1 - dFuelDateFactor), 0)
    .value(car_fueldaysreduction) = RoundDownEx(.value(car_fueldaysreduction) / .value(car_NumberOfUsers_db), 0)
    
    .value(car_FuelGrossDividedByUsers) = RoundDownEx((.value(car_GrossFuel)) / .value(car_NumberOfUsers_db), 0)
    .value(car_FuelValue_Non_OPRA) = Max(.value(car_FuelGrossDividedByUsers) - .value(car_fueldaysreduction), 0)
    
    Dim iFuelValueNonOPRAField As Integer
    If (.value(car_requiredmakegood_db) And .value(car_actualmadegood_db) And (.value(car_FuelOPRA_Ammount_Foregone_db) > 0)) Then
      'we need a temp field to have a  value of zero to force the value to 0 in CalculateOpRAValue
      'TODO: issue with working paper, this is an edge case for the last ever release!
      .value(car_FuelOPRAZeroField) = 0
      iFuelValueNonOPRAField = car_FuelOPRAZeroField
    Else
      iFuelValueNonOPRAField = car_FuelValue_Non_OPRA
    End If
    
    Call CalculateOpRAValue(ben, ben.value(car_FuelDontUseAmountForegone), iFuelValueNonOPRAField, car_FuelOPRA_Ammount_Foregone_db, car_FuelValue, car_FuelOPRA_Ammount_Foregone_Used_For_Value)
    
    If Not (.value(car_FuelOPRA_Ammount_Foregone_Used_For_Value)) Then
      '.value(car_FuelValue) = RoundDownEx(.value(car_FuelValue_Non_OPRA) / .value(car_NumberOfUsers_db), 0)
      
    ElseIf (.value(car_FuelMadeGood) > 0) Then    'actually made good or electric
      If (iFuelValueNonOPRAField = car_FuelOPRAZeroField) Then
        .value(car_FuelMadeGood) = 0
      Else
        .value(car_FuelMadeGood) = .value(car_FuelOPRA_Ammount_Foregone_db)
      End If
        
    
      
    End If
    
    .value(car_FuelMadeGood_NET) = Min(ben.value(car_FuelValue), ben.value(car_FuelMadeGood))
    
    .value(car_FuelBenefit) = .value(car_FuelValue) - .value(car_FuelMadeGood_NET)
    .value(car_FuelVATBenefit) = (.value(car_FuelVATBenefit) * dFuelDateFactor) / .value(car_NumberOfUsers_db)
    
    'get the vat on the number
        
    vatRates(1).From = #1/1/1990#
    vatRates(1).To = #11/30/2008#
    vatRates(1).Rate = 0.175
    
    vatRates(2).From = #12/1/2008#
    vatRates(2).To = #12/31/2009#
    vatRates(2).Rate = 0.15
    
    vatRates(3).From = #1/1/2010#
    vatRates(3).To = #1/3/2011#
    vatRates(3).Rate = 0.175
    
    vatRates(4).From = #1/4/2011#
    vatRates(4).To = #12/31/2050#
    vatRates(4).Rate = 0.2
    
    
    .value(car_FuelVAT) = 0
    '.value(car_FuelVAT) = RoundN(.value(car_FuelVATBenefit) * (7 / 47))
    For l = LBound(vatRates) To UBound(vatRates)
      iDaysAtRate = DaysAtRate(dFuelDateFrom, dFuelDateTo, vatRates(l))
      If (iDaysAtRate > 0) Then
        dVAT = (CDbl(.value(car_FuelVATBenefit)) - (CDbl(.value(car_FuelVATBenefit)) / (1 + vatRates(l).Rate))) * (CDbl(iDaysAtRate) / CDbl(iDaysAvailable))
        .value(car_FuelVAT) = .value(car_FuelVAT) + dVAT
      End If
    Next l
    .value(car_FuelVAT) = RoundN(.value(car_FuelVAT))
    '.value(car_FuelVAT) = RoundN(.value(car_FuelVATBenefit) - (.value(car_FuelVATBenefit) / (1 + p11d32.Rates.value(vatRate))))
    
    
    
    Call BenCalcNIC(Fuel)
  End With

CalculateFuelBen_End:
  Call xReturn("CalculateFuelBen_End")
  Exit Sub
  
CalculateFuelBen_Err:
  Call Err.Raise(ERR_CALCULATE_FUEL_BENEFITS, "CalculateFuelBen", "Error in Calculating Fuel Benefits")
  Resume CalculateFuelBen_End
  Resume
End Sub

Private Function DaysAtRate(availableFrom As Date, availableTo As Date, vr As VAT_RATE) As Long
  Dim iDaysAtRate As Long
  If (availableFrom >= vr.From And availableFrom <= vr.To) Then
    If (availableTo <= vr.To) Then
      'contained
      iDaysAtRate = DateDiff("d", availableFrom, availableTo) + 1
    Else
      iDaysAtRate = DateDiff("d", availableFrom, vr.To) + 1
    End If
  ElseIf (availableTo >= vr.From And availableTo <= vr.To) Then
    If (availableFrom >= vr.To) Then
      'contained
      iDaysAtRate = DateDiff("d", availableFrom, availableTo) + 1
    Else
      iDaysAtRate = DateDiff("d", vr.From, availableTo) + 1
    End If
  End If
  DaysAtRate = iDaysAtRate
End Function
'need to investigate why we have this below below CAD 05/05
Private Sub SetHMIT()
  Dim ben As IBenefitClass
  
  On Error GoTo SetHMIT_ERR
  Call xSet("SetHMIT")
  
  Set ben = Me

  If ben.value(car_FuelBenefit) > 0 Then
    ben.value(car_HMIT_CC) = ben.value(car_enginesize_db)
    If ben.value(car_p46FuelType_db) = CCFT_DIESEL Then
      ben.value(car_HMIT_diesel) = True
      ben.value(car_HMIT_petrol) = False
    Else
      ben.value(car_HMIT_diesel) = False
      ben.value(car_HMIT_petrol) = True
    End If
  Else
    ben.value(car_HMIT_CC) = ""
  End If
  
SetHMIT_END:
  Call xReturn("SetHMIT")
  Exit Sub
SetHMIT_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetHMIT", "Set HMIT", "Error setting HMIT specific items in company car benefit calculate.")
  Resume SetHMIT_END
End Sub

Private Function IBenefitClass_ReadDB() As Long
  Dim rs As Recordset
  Dim carben As IBenefitClass
  Dim car As CompanyCar
  Dim i As Long
  Dim ben As IBenefitClass
  Dim iLen As Long
  
  On Error GoTo CompanyCar_readDB_Err
  Call xSet("CompanyCar_readDB")
  
  Set ben = Me
  
  If m_ReadFromDB Then GoTo CompanyCar_readDB_End
  Set rs = p11d32.CurrentEmployer.rsBenTables(ben.TABLE)
  If Len(ben.RSBookMark) = 0 Then
    rs.FindFirst ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
    If Not rs.NoMatch Then
      Call AddFuelBenefit
      m_sbookmark = rs.Bookmark
      i = i + 1
      rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set car = New CompanyCar
        Set carben = car
        Set carben.Parent = m_Parent
        Call car.AddFuelBenefit
        carben.RSBookMark = rs.Bookmark
        Call m_Parent.benefits.Add(car)
        Set car = Nothing
        Set carben = Nothing
        rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  
  If Len(ben.RSBookMark) > 0 Then
    rs.Bookmark = ben.RSBookMark
    Call StandardReadData(ben, rs)
    Call OPRAReadDB(ben, rs, car_FuelOPRA_Ammount_Foregone_db, "Fuel")
    ben.value(car_Registration_db) = IIf(IsNull(rs.Fields(S_FIELD_CAR_REGISTRATION).value), "", rs.Fields(S_FIELD_CAR_REGISTRATION).value)
    ben.value(car_Make_db) = "" & rs.Fields("Make").value
    ben.value(car_Model_db) = IIf(IsNull(rs.Fields("Model").value), "", rs.Fields("Model").value)
    
    'P46 replacement
    ben.value(car_RegReplaced_db) = IIf(IsNull(rs.Fields("RegReplaced").value), "", rs.Fields("RegReplaced").value)
    ben.value(car_CarReplaced_db) = IIf(IsNull(rs.Fields("CarReplaced").value), "", rs.Fields("CarReplaced").value)
    ben.value(car_ListPrice_db) = IIf(IsNull(rs.Fields("Price").value), 0, rs.Fields("Price").value)
    ben.value(car_UnavailableDays_db) = IIf(IsNull(rs.Fields("Unavail").value), 0, rs.Fields("Unavail").value)
    ben.value(car_capitalcontribution_db) = IIf(IsNull(rs.Fields("CapContrib").value), 0, rs.Fields("CapContrib").value)
    ben.value(car_MadeGood_db) = IIf(IsNull(rs.Fields("UseContrib").value), 0, rs.Fields("UseContrib").value)
    ben.value(car_enginesize_db) = IIf(IsNull(rs.Fields("cc").value), 0, rs.Fields("cc").value)
    ben.value(car_AccessoriesOriginal_db) = IIf(IsNull(rs.Fields("Acc").value), 0, rs.Fields("Acc").value)
    ben.value(car_AccessoriesNew_db) = IIf(IsNull(rs.Fields("NewAcc").value), 0, rs.Fields("NewAcc").value)
    ben.value(car_CheapAccessories_db) = IIf(IsNull(rs.Fields("CheapAcc").value), 0, rs.Fields("CheapAcc").value)
    ben.value(car_carkey_db) = rs.Fields("CarKey").value
    ben.value(car_Replacement_db) = rs.Fields("Replacement").value
    ben.value(car_Replaced_db) = rs.Fields("Replaced").value
    ben.value(car_Second_db) = rs.Fields("SecondCar").value
    ben.value(car_ElectricRangeMiles_db) = IIf(IsNull(rs.Fields(S_FIELD_CAR_ELECTRIC_RANGE_MILES).value), 0, rs.Fields(S_FIELD_CAR_ELECTRIC_RANGE_MILES).value)
    ben.value(car_privatefuel_db) = rs.Fields("PvtFuel").value
    ben.value(car_requiredmakegood_db) = rs.Fields("MakeGood").value
    ben.value(car_actualmadegood_db) = rs.Fields("MadeGood").value
    ben.value(car_ForceP46_db) = rs.Fields("ForceP46").value
        
    ben.value(car_Registrationdate_db) = IIf(IsNull(rs.Fields("RegDate").value), p11d32.Rates.value(CarRegDateDef), rs.Fields("RegDate").value)
    ben.value(Car_AvailableFrom_db) = IIf(IsNull(rs.Fields("AvailFrom").value), p11d32.Rates.value(TaxYearStart), rs.Fields("AvailFrom").value)
    
    ben.value(Car_AvailableTo_db) = IIf(IsNull(rs.Fields("AvailTo").value), p11d32.Rates.value(TaxYearEnd), rs.Fields("AvailTo").value)
    ben.value(car_dateReplaced_db) = IIf(IsNull(rs.Fields("DateCarReplaced").value), UNDATED, rs.Fields("DateCarReplaced").value)
    
    ben.value(Car_FuelAvailableTo_db) = IIf(IsNull(rs.Fields("FuelAvailTo").value), ben.value(Car_AvailableTo_db), rs.Fields("FuelAvailTo").value)
    
    ben.value(car_p46FuelType_db) = IIf(IsNull(rs.Fields("P46FuelType").value), CCFT_PETROL, rs.Fields("P46FuelType").value)

    If p11d32.BringForward.Yes Then  'RH
      ben.value(car_NumberOfUsers_db) = 1
      ben.value(car_fuelreinstated_db) = False
      
      '666
      ben.value(car_FuelMadeGoodIsTaxDeducted_db) = False
      ben.value(car_p46PaymentFrequency_db) = P46_PAYMENT_FREQUENCY.P46PF_ACTUAL
      
      
        
    Else
      ben.value(car_NumberOfUsers_db) = IIf(IsNull(rs.Fields("NumOfUsers").value), 1, rs.Fields("NumOfUsers").value)
      ben.value(Car_HasFuelUnavailableDays_db) = rs.Fields("HasFuelUnavail").value
      
      'EK
      ben.value(car_fuelreinstated_db) = IIf(IsNull(rs.Fields("FuelReinstated").value), 0, rs.Fields("FuelReinstated").value)
      ben.value(car_p46PaymentFrequency_db) = IIf(IsNull(rs.Fields("P46PaymentFrequency").value), P46PF_ACTUAL, rs.Fields("P46PaymentFrequency").value)
      ben.value(car_FuelMadeGoodIsTaxDeducted_db) = rs.Fields("FuelMadeGoodIsTaxDeducted")
      
      
    End If
    
    ben.value(car_CarReplacedMake_db) = IIf(IsNull(rs.Fields("CarReplacedMake").value), "", rs.Fields("CarReplacedMake").value)
    ben.value(car_CarReplacedModel_db) = IIf(IsNull(rs.Fields("CarReplacedModel").value), "", rs.Fields("CarReplacedModel").value)
    ben.value(car_CarReplacedEngineSize_db) = IIf(IsNull(rs.Fields("CarReplacedEngineSize").value), 0, rs.Fields("CarReplacedEngineSize").value)
    
    ben.value(car_p46CarbonDioxide_db) = IIf(IsNull(rs.Fields("P46CarbonDioxide").value), 0, rs.Fields("P46CarbonDioxide").value)
    ben.value(car_p46NoApprovedCO2Figure_db) = IIf(IsNull(rs.Fields("P46NoApprovedCO2Figure").value), False, rs.Fields("P46NoApprovedCO2Figure").value)
    
    'just incase the value > 3 in length
    iLen = p11d32.BenDataLinkMMFieldSize(BC_COMPANY_CARS_F, car_p46CarbonDioxide_db)
    If (Len(CStr(ben.value(car_p46CarbonDioxide_db))) > iLen) Then
      ben.value(car_p46CarbonDioxide_db) = Right$(CStr(ben.value(car_p46CarbonDioxide_db)), iLen)
    End If
    
    m_ReadFromDB = True
    Call StandardReadData(Fuel, Nothing, False)
  End If
  
CompanyCar_readDB_End:
  IBenefitClass_ReadDB = i
  Set rs = Nothing
  Call xReturn("CompanyCar_readDB")
  Exit Function
CompanyCar_readDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CompanyCar_readDB", "ERR_CAR_READDB", "Unable to read the company car details.")
  Resume CompanyCar_readDB_End
  Resume
End Function

Public Property Get IBenefitClass_PrintHeader() As String
  IBenefitClass_PrintHeader = "{x=0}{Arial=10,n}Car and car fuel benefit" & vbCrLf
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_COMPANY_CARS
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_COMPANY_CARS_F, Item, RHS)
  
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo CompanyCar_WriteDB_Err
  Call xSet("CompanyCar_WriteDB")
  
  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields(S_FIELD_CAR_REGISTRATION).value = ben.value(car_Registration_db)
  rs.Fields("Make").value = ben.value(car_Make_db)
  rs.Fields("Model").value = ben.value(car_Model_db)
  rs.Fields("Price").value = ben.value(car_ListPrice_db)
  rs.Fields("CapContrib").value = ben.value(car_capitalcontribution_db)
  rs.Fields("cc").value = ben.value(car_enginesize_db)
  rs.Fields("Acc").value = ben.value(car_AccessoriesOriginal_db)
  rs.Fields("NewAcc").value = ben.value(car_AccessoriesNew_db)
  rs.Fields("CheapAcc").value = ben.value(car_CheapAccessories_db)
  
  
  rs.Fields("RegDate").value = IIf(TryConvertDateDMY(CStr(ben.value(car_Registrationdate_db))) = UNDATED, Null, ben.value(car_Registrationdate_db))
  rs.Fields("SecondCar").value = ben.value(car_Second_db)

  'km 12/06/02 - this field no longer exists in the 2002 db
  'rs.Fields("BusMilesActual") = ben.value(car_UseActualMiles)
  rs.Fields("P46PaymentFrequency").value = ben.value(car_p46PaymentFrequency_db) 'so
  rs.Fields("P46FuelType").value = ben.value(car_p46FuelType_db) 'so
  rs.Fields("P46CarbonDioxide").value = ben.value(car_p46CarbonDioxide_db) 'so
  rs.Fields("P46NoApprovedCO2Figure").value = ben.value(car_p46NoApprovedCO2Figure_db) 'so
  'rs.Fields("FuelUnavail").value = ben.value(car_FuelUnavailableDays)
  rs.Fields(S_FIELD_CAR_ELECTRIC_RANGE_MILES).value = ben.value(car_ElectricRangeMiles_db)
  
  Call StandardWriteData(ben, rs)
  
  Call BringForwardDatesWrite(ben, Car_AvailableFrom_db, Car_AvailableTo_db, rs, "AvailFrom", "AvailTo")
  
  If p11d32.BringForward.Yes Then
    If (ben.value(car_privatefuel_db)) And (ben.value(Car_AvailableTo_db) <= ben.value(Car_FuelAvailableTo_db) Or ben.value(car_fuelreinstated_db)) Then
      rs.Fields("PvtFuel").value = True
    End If
    'car_fuelreinstated_db
    'If ben.value(car_privatefuel_db) And (ben.value(Car_AvailableTo_db) > ben.value(Car_FuelAvailableTo_db)) Then
  
    
    rs.Fields("FuelAvailTo") = p11d32.Rates.value(TaxYearEnd)
    rs.Fields("P46PaymentFrequency").value = ben.value(car_p46PaymentFrequency_db)
    rs.Fields("NumOfUsers").value = 1 'ben.value(car_NumberOfUsers)
    rs.Fields("FuelMadeGoodIsTaxDeducted").value = False
  Else
    rs.Fields("PvtFuel").value = ben.value(car_privatefuel_db)
    rs.Fields("MakeGood").value = ben.value(car_requiredmakegood_db)
    rs.Fields("MadeGood").value = ben.value(car_actualmadegood_db)
    
    rs.Fields("FuelMadeGoodIsTaxDeducted") = ben.value(car_FuelMadeGoodIsTaxDeducted_db)
    rs.Fields("NumOfUsers").value = CLng(ben.value(car_NumberOfUsers_db))
    rs.Fields("FuelReinstated").value = ben.value(car_fuelreinstated_db)
    rs.Fields("HasFuelUnavail").value = ben.value(Car_HasFuelUnavailableDays_db)
    
    If Len(ben.value(car_RegReplaced_db)) Then rs.Fields("RegReplaced").value = ben.value(car_RegReplaced_db)
    
    If Len(ben.value(car_CarReplaced_db)) Then rs.Fields("CarReplaced").value = ben.value(car_CarReplaced_db)
    If Len(ben.value(car_CarReplacedMake_db)) Then rs.Fields("CarReplacedMake").value = ben.value(car_CarReplacedMake_db)
    If Len(ben.value(car_CarReplacedModel_db)) Then rs.Fields("CarReplacedModel").value = ben.value(car_CarReplacedModel_db)
    If Len(ben.value(car_CarReplacedEngineSize_db)) Then rs.Fields("CarReplacedEngineSize").value = ben.value(car_CarReplacedEngineSize_db)
    If Len(ben.value(car_dateReplaced_db)) Then rs.Fields("DateCarReplaced").value = ben.value(car_dateReplaced_db)

    rs.Fields("Unavail").value = ben.value(car_UnavailableDays_db)
    rs.Fields("UseContrib").value = ben.value(car_MadeGood_db)
    rs.Fields("Replacement").value = ben.value(car_Replacement_db)
    rs.Fields("Replaced").value = ben.value(car_Replaced_db)
    rs.Fields("ForceP46").value = ben.value(car_ForceP46_db)
    rs.Fields("FuelAvailTo").value = ben.value(Car_FuelAvailableTo_db)
  End If
  Call OPRAWriteDB(ben, rs, car_FuelOPRA_Ammount_Foregone_db, "Fuel")
 
  Call BenefitCloseRecord(ben, rs, True)
  ben.value(car_carkey_db) = rs.Fields("CarKey")
  
' MP DB - T_CarMiles deleted
'  If Not p11d32.BringForward.Yes Then Call WriteMileage
  
  IBenefitClass_WriteDB = True
CompanyCar_WriteDB_End:
  Set rs = Nothing
  Call xReturn("CompanyCar_WriteDB")
  Exit Function
CompanyCar_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "CompanyCar_WriteDB", "CompanyCar Write DB", "Error writing the company car to the database.")
  Resume CompanyCar_WriteDB_End
  Resume
End Function

' MP DB - T_CarMiles deleted
'Public Function WriteMileage() As Long
'  Dim Miles As MileDetail
'  Dim rs As Recordset
'  Dim l As Long
'  Dim ben As IBenefitClass, benEmployee As IBenefitClass
'
'  On Error GoTo WriteMileage_Err
'  Call xSet("WriteMileage")
'
'  Set ben = Me
'  Set benEmployee = m_Parent
'  Call m_Parent.Parent.db.Execute(sql.Queries(DELETE_CARMILES, ben.value(car_carkey)))
'
'  Set rs = m_Parent.Parent.db.OpenRecordset(sql.Queries(SELECT_CARMILES, ben.value(car_carkey)), dbOpenDynaset)
'
'  For l = 1 To mileage.Count
'    Set Miles = mileage.Item(l)
'    If Not Miles Is Nothing Then
'      rs.AddNew
'      If Len(Miles.MileItem) > 0 Then
'        rs.Fields("Item").value = Miles.MileItem
'      Else
'        rs.Fields("Item").value = Null
'      End If
'      If Miles.MileDate <> UNDATED Then
'        rs.Fields("MilesDate").value = Miles.MileDate
'      Else
'        rs.Fields("MilesDate").value = Null
'      End If
'      rs.Fields("Miles").value = Miles.MileAmount
'      rs.Fields("CarKey").value = ben.value(car_carkey)
'      rs.Fields("P_Num").value = benEmployee.value(ee_PersonnelNumber)
'      rs.Fields(S_FIELD_CAR_REGISTRATION).value = ben.value(car_Registration)
'      rs.Update
'      WriteMileage = WriteMileage + 1
'    End If
'  Next
'
'WriteMileage_End:
'  Set Miles = Nothing
'  Set rs = Nothing
'  Call xReturn("WriteMileage")
'  Exit Function
'
'WriteMileage_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "WriteMileage", "Write Mileage", "Error writing the car mileage to the database.")
'  Resume WriteMileage_End
'  Resume
'End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  On Error GoTo CompanyCar_DeleteDB_Err
  
  Call xSet("CompanyCar_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_COMPANY_CARS)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
'MP DB    Call m_Parent.Parent.db.Execute(sql.Queries(DELETE_CARMILES, m_BenItems(car_carkey)))
  End If
  
  IBenefitClass_DeleteDB = True
  
CompanyCar_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("CompanyCar_DeleteDB")
  Exit Function
CompanyCar_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CompanyCar_DeleteDB", "Company Car Delete DB", "Error deleting a company car from the database.")
  Resume CompanyCar_DeleteDB_End
  Resume
End Function
Public Function SetP46CC() As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo SetP46CC_Err
  Call xSet("SetP46CC")
  
  Set ben = Me
  
  ben.value(car_P46HighCC) = False
  ben.value(car_P46MediumCC) = False
  ben.value(car_P46LowCC) = False
  ben.value(car_P46NoCC) = False
  
  'RH - display of engine size should not be dependant on whether there is a fuel benefit
  'If ben.value(car_privatefuel_db) And Not ben.value(car_requiredmakegood_db) Then
  
    Select Case CLng(ben.value(car_enginesize_db))
      Case Is > p11d32.Rates.value(carCCHIGH)
        ben.value(car_P46HighCC) = True
      Case Is > p11d32.Rates.value(carCCLOW)
        ben.value(car_P46MediumCC) = True
      Case Is = 0
        ben.value(car_P46NoCC) = True
      Case Else
        ben.value(car_P46LowCC) = True
    End Select
    
  'End If
  

  SetP46CC = True

SetP46CC_End:
  Call xReturn("SetP46CC")
  Exit Function

SetP46CC_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetP46CC", "Set P46 CC", "Error setting the P46 car Engine size band.")
  Resume SetP46CC_End
End Function


Private Function CalculatePaymentForPrivateUse(Optional DateFactor As Variant) As Variant
  Dim l As Long
  Dim ben As IBenefitClass
  Dim vPayment
  On Error GoTo GetPayment_ERR
  Call xSet("GetPayment")
  
  Set ben = Me
  vPayment = S_ERROR
  DateFactor = dGetDateFactor(l, ben.value(Car_AvailableFrom_db), ben.value(Car_AvailableTo_db), m_BenItems(car_UnavailableDays_db), L_DAYS_UNAVAIL_MIN, False)

  With ben
    Select Case .value(car_p46PaymentFrequency_db)
      Case 0
        vPayment = (ben.value(car_MadeGood_db) * DateFactor)
      Case 1
        vPayment = (ben.value(car_MadeGood_db) * DateFactor * 4)
      Case 2
        vPayment = (ben.value(car_MadeGood_db) * DateFactor * 12)
      Case 3
        vPayment = (ben.value(car_MadeGood_db) * DateFactor * 52)
      Case 4
        vPayment = ben.value(car_MadeGood_db)
      Case Else
        Call Err.Raise(ERR_CAR_AMOUNT, "GetPayment", "Invalid payment frequency")
    End Select
  End With
    
  CalculatePaymentForPrivateUse = vPayment
  
GetPayment_END:
  Call xReturn("GetPayment")
  Exit Function
GetPayment_ERR:
  Resume GetPayment_END
  
  Resume

End Function

Public Function SetP46FrequencyPaymentString()
Dim ben As IBenefitClass

  On Error GoTo SetP46FrequencyPaymentString_ERR
  Call xSet("SetP46FrequencyPaymentString")
  
  Set ben = Me
  With ben
    Select Case .value(car_p46PaymentFrequency_db)
      Case CCPF_ANNUALLY
        .value(car_p46PaymentFrequencyString) = "Annually"
      Case CCPF_QUARTERLY
        .value(car_p46PaymentFrequencyString) = "Quarterly"
      Case CCPF_MONTHLY
        .value(car_p46PaymentFrequencyString) = "Monthly"
      Case CCPF_WEEKLY
        .value(car_p46PaymentFrequencyString) = "Weekly"
      Case CCPF_ACTUAL
        .value(car_p46PaymentFrequencyString) = "Actual"
    End Select
  End With

SetP46FrequencyPaymentString_END:
  Call xReturn("SetP46FrequencyPaymentString")
  Exit Function
SetP46FrequencyPaymentString_ERR:
  Resume SetP46FrequencyPaymentString_END
  Resume

End Function


Public Function SetP46FuelTypeString()
  Dim ben As IBenefitClass
  Dim p46s As P46_FUEL_TYPE_STRINGS
  
  On Error GoTo SetP46FuelTypeString_ERR
  Call xSet("SetP46FuelTypeString")
  
  Set ben = Me
  
  p46s = P46FuelTypeStrings(ben.value(car_p46FuelType_db))
  
  With ben
    .value(car_p46FuelTypeString) = p46s.Letter
    .value(car_p46FuelTypeStringLong) = p46s.Description
  End With

  
  

SetP46FuelTypeString_END:
  Call xReturn("SetP46FuelTypeString")
  Exit Function
SetP46FuelTypeString_ERR:
  Resume SetP46FuelTypeString_END
  Resume


End Function



