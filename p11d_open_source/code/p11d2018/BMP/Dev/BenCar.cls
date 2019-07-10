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
  car_NIC_Class1A_Value = ITEM_NIC_CLASS1A_VALUE
  car_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  
  
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
  
  Car_FuelAvailableTo_db '
  
  
  car_FUEL_SPECIFIC_END = Car_FuelAvailableTo_db
   
  
  car_NumberOfUsers_db
  car_FuelUnavailableDays 'used in calc only
  Car_HasFuelUnavailableDays_db
  car_fuelreinstated_db
  Car_FuelAvailableTo_calc
  car_fuelreinstated_calc
  car_FuelVATBenefit
  car_FuelVAT
  
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
  car_AccessoriesTotal
  car_CheapAccessories_db 'bf
  car_carkey_db
'AM To be removed  car_SumMiles
  car_GrossPrice
  car_Price
  car_Scaled
'AM To be removed  car_BusMilereduction
'MP DB ToDo chk     car_MileScaled
'AM To be removed  car_Agereduction
  car_accessories
'AM To be removed  car_AgeScaled
  car_GrossFuel
  car_fueldaysreduction ' EK now same as car_cardaysreduction if car_HasFuelUnavailableDays or zero
  car_cardaysreduction
  'Dates
  car_Registrationdate_db 'bf
  Car_AvailableFrom_db
  Car_AvailableTo_db 'bf test
  car_dateReplaced_db
  'Booleans
'AM To be removed  car_UseActualMiles
  car_Replacement_db
  car_Replaced_db
  car_Second_db
  
'AM Car_Diesel To be removed for main 2002 release
'MP DB ToDo remove car_Diesel? chk its use in MagneticMedia.SectionF
  car_privatefuel_db 'bf
  car_requiredmakegood_db
  car_actualmadegood_db
  car_ForceP46_db
  
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
  
  'HMIT specific items
  car_HMIT_CC
  car_HMIT_petrol
  car_HMIT_diesel
  
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



Private Function IBenefitClass_CalculateBody() As Variant
  'no calcualate only when required as p46 flags depend on then current dates
  Dim benFueld As IBenefitClass
  
  Dim ben As IBenefitClass
  Dim dPercentageOfValue As Double
  Dim vNewLessCheapAccessories As Variant
  Dim dDateFactor As Double
  'see booklet 480 for reference for the year 2003/3004
  
On Error GoTo err_Err

Call xSet("CompanyCar_Calculate")
  
  Set ben = Me
  Call ben.SetCalcDefaults
  
  'set the engine size flags, <=1400, 1400-2000 etc, these are required for p46
  Call SetP46CC
  'set other stuff required elswhere !!!
  Call SetP46FuelTypeString
  Call SetP46FrequencyPaymentString
  
  
  Call CompanyCarAvailableDatesValid(ben)
  ' ******************************** CAR CALC START ***********************************************
  

  'the calculation divides into 2 clear areas that get a percentage to apply to the value of the car
  '1: cars registered pre 01/01/1998
  If ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
    
    'this is purely based on the CC of the car, even if an official CO2 figure existed, special cases for electricity
    If (ben.value(car_p46FuelType_db) = CCFT_ELECTRIC) Then
      dPercentageOfValue = 0.15
    Else
      If (ben.value(car_P46HighCC)) Then
        dPercentageOfValue = 0.32
      ElseIf ben.value(car_P46MediumCC) Then
        dPercentageOfValue = 0.22
      ElseIf ben.value(car_P46LowCC) Then
        dPercentageOfValue = 0.15
      Else 'must be 0 then we must consider as wankel engine, eg ben.value(car_P46NoCC) = true
        dPercentageOfValue = 0.32
      End If
    End If
  Else
  '2: cars registered post and including 01/01/1998
    'first get the base percentage
    If (ben.value(car_p46FuelType_db) = CCFT_ELECTRIC) Then
      dPercentageOfValue = 0.15
      
    Else
      'if no approved C02 then use section 12.26
      If (ben.value(car_p46NoApprovedCO2Figure_db)) Then
        If (ben.value(car_P46HighCC)) Then
          dPercentageOfValue = 0.35
        ElseIf ben.value(car_P46MediumCC) Then
          dPercentageOfValue = 0.25
        ElseIf ben.value(car_P46LowCC) Then
          dPercentageOfValue = 0.15
        Else 'must be 0 then we must consider as wankel engine, eg ben.value(car_P46NoCC) = true
          dPercentageOfValue = 0.35
        End If
      Else
        dPercentageOfValue = ReadyReckonerPercentage()
        
      End If
    End If
    ben.value(car_unrestrictedPercentage) = dPercentageOfValue
    'apply the supplement from section 12.27 of the revenue guide
    Select Case ben.value(car_p46FuelType_db)
      Case CCFT_PETROL, CCFT_EUROIVDIESEL
        'nothing
      Case CCFT_GAS_ONLY, CCFT_BIFUEL_WITH_CO2_FOR_GAS, CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B   '(gas only?)
        dPercentageOfValue = dPercentageOfValue - 0.01
      Case CCFT_DIESEL
        dPercentageOfValue = dPercentageOfValue + 0.03
      Case CCFT_ELECTRIC
        dPercentageOfValue = dPercentageOfValue - 0.06 ' leaves 0.09%
      Case CCFT_HYBRID
        dPercentageOfValue = dPercentageOfValue - 0.02
      Case Else
        Call Err.Raise(ERR_CAR_CO2, "Calculate20032004", "Unknown fuel type")
    End Select
  End If

  
  dPercentageOfValue = Max(0, Min(dPercentageOfValue, 0.35))
  ben.value(car_percentage) = dPercentageOfValue
  
  vNewLessCheapAccessories = ben.value(car_AccessoriesNew_db) - ben.value(car_CheapAccessories_db)
  ben.value(car_accessories) = ben.value(car_AccessoriesOriginal_db) + IIf(vNewLessCheapAccessories > 0, vNewLessCheapAccessories, 0)
  ben.value(car_GrossPrice) = ben.value(car_ListPrice_db) + ben.value(car_accessories)
  ben.value(car_CapitalContributionRestricted) = Max(Min(ben.value(car_capitalcontribution_db), p11d32.Rates.value(carMAXCAPCONTRIB)), 0)
  ben.value(car_Price) = Max(0, Min(ben.value(car_GrossPrice) - ben.value(car_CapitalContributionRestricted), p11d32.Rates.value(carMAXLISTPRICE)))
  
  dDateFactor = dGetDateFactor(0, ben.value(Car_AvailableFrom_db), ben.value(Car_AvailableTo_db), m_BenItems(car_UnavailableDays_db), L_DAYS_UNAVAIL_MIN, False)
  
  ben.value(car_Scaled) = RoundDownEx(ben.value(car_Price) * ben.value(car_percentage), 0)
  
  ben.value(car_carValue) = RoundDownEx(ben.value(car_Scaled) * dDateFactor, 0)
  
  ben.value(car_cardaysreduction) = ben.value(car_Scaled) - ben.value(car_carValue)  'ben.value(car_Scaled) * (1 - dDateFactor)
  
  ben.value(car_ActualAmountMadeGood) = GetPayment(dDateFactor)
  ben.value(car_MadeGood_Net) = Min(ben.value(car_carValue), ben.value(car_ActualAmountMadeGood))
  ben.value(car_carBenefit) = ben.value(car_carValue) - ben.value(car_MadeGood_Net)
  
  'dividing benefit by number of shared users
  ben.value(car_carBenefit) = RoundDownEx(ben.value(car_carBenefit) / (ben.value(car_NumberOfUsers_db)), 0)
  ' ******************************** CAR CALC END ***********************************************
  
  
  '****************************************** FUEL CALC START ********************************************
  Call CalculateFuelBen(dPercentageOfValue)  'this is retained from original calc
    
  '****************************************** FUEL CALC END ********************************************
  Call BenCalcNIC(ben)
  
  
  
  
  ben.value(car_Benefit_Reportable) = True
  ben.value(car_FuelBenefit_Reportable) = True
  
  Call SetHMIT

  IBenefitClass_CalculateBody = ben.value(car_carBenefit)
   
err_End:
  Call xReturn("CompanyCar_Calculate")
  Exit Function
  
err_Err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume err_End
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
    Call WKOut(rep, WK_ITEM_TEXT, "CO2 emissions figure: " & S_WK_NORMAL_BOLD_FONT & ben.value(car_p46CarbonDioxide_db) & S_WK_NORMAL_FONT & " g/km")
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
  
  Call WKOut(rep, WK_ITEM_Total, "Adjusted price of the car for " & p11d32.Rates.value(TaxFormYear) & " (restricted to " & FormatWN(p11d32.Rates.value(carMAXLISTPRICE)) & ")", ben.value(car_Price), , True)
    
'AM To be removed  Call WKOut(rep, WK_ITEM_TEXT, "Gross benefit for " & p11d32.Rates.value(TaxFormYear) & " (Adjusted price x " & ben.value(car_ReductionPercentageWK) & "%)", ben.value(car_Scaled) - ben.value(car_BusMilereduction))
'AM To be removed  Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "Note: " & ben.value(car_ReductionPercentageWK) & "% reduction as car in " & ben.value(car_businessmiles) & " band.")
  Call WKOut(rep, WK_BLANK_LINE)
  
'AM To be removed  Call WKOut(rep, WK_ITEM_TEXT, "Reduction for age of car (first registered " & ben.value(car_Registrationdate) & ")", ben.value(car_Agereduction), , , True)
  Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)

' EK Added outline
  If (ben.value(car_p46FuelTypeString) = "D") Then lFuelDeduction = p11d32.Rates.value(CarDIESELSUPPLEMENT) * 100
  If ben.value(car_p46NoApprovedCO2Figure_db) Or ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
    Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Relevant percentage based on engine size (limited so total maximum is 35%)", S_WKCOL_3 & (lFinalPercentage - lFuelDeduction) & "%")
    Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Diesel Supplement [if applicable] ", S_WKCOL_3 & " " & IIf(lFuelDeduction > 0, lFuelDeduction, "0") & "%")
  
  Else
    If (ben.value(car_p46FuelTypeString) = "H" Or ben.value(car_p46FuelTypeString) = "B") Then
      If ben.value(car_p46CarbonDioxide_db) < p11d32.Rates.value(CarCO2QUALIFYINGLEVEL) Then
        lAdditionalDiscount = RoundDownEx((p11d32.Rates.value(CarCO2QUALIFYINGLEVEL) - ben.value(car_p46CarbonDioxide_db)) / 20, 0)
      End If
    End If
    If ben.value(car_p46FuelTypeString) = "E" Then lFuelDeduction = lFuelDeduction - (p11d32.Rates.value(CarELECTRICDISCOUNT) * 100)
    If ben.value(car_p46FuelTypeString) = "H" Then lFuelDeduction = lFuelDeduction - (p11d32.Rates.value(CarHYBRIDELECTRICDISCOUNT) * 100) - lAdditionalDiscount
    If ben.value(car_p46FuelTypeString) = "B" Then lFuelDeduction = lFuelDeduction - (p11d32.Rates.value(CarSTDDISCOUNTSUPPLEMENT) * 100) - lAdditionalDiscount
    If ben.value(car_p46FuelTypeString) = "C" Then lFuelDeduction = lFuelDeduction - (p11d32.Rates.value(CarSTDDISCOUNTSUPPLEMENT) * 100)
    
    Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Minimum tax charge ", S_WKCOL_3 & p11d32.Rates.value(carBENPERCENT) * 100 & "%")
    If lFuelDeduction > 0 Then Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Diesel Supplement  ", S_WKCOL_3 & " " & lFuelDeduction & "%")
    lRunningPercentage = lFinalPercentage - lFuelDeduction - (p11d32.Rates.value(carBENPERCENT) * 100)
    If (ben.value(car_unrestrictedPercentage) > p11d32.Rates.value(carBENPERCENT)) Then
      Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Additional charge as CO2 greater than qualifying level (" & p11d32.Rates.value(CarCO2QUALIFYINGLEVEL) & "g/km) ", S_WKCOL_3 & (ben.value(car_unrestrictedPercentage) - p11d32.Rates.value(carBENPERCENT)) * 100 & "%")
      Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "         (" & ben.value(car_p46CarbonDioxide_db) & " - " & p11d32.Rates.value(CarCO2QUALIFYINGLEVEL) & ") / 5 rounded to 5")
    End If
    If lFuelDeduction < 0 And ben.value(car_p46FuelTypeString) <> "D" Then
      If lFuelDeduction > 0 Or (ben.value(car_unrestrictedPercentage) > p11d32.Rates.value(carBENPERCENT)) Then
        Call WKOut(rep, WK_BLANK_LINE)
        Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_BOLD_FONT & "     Subtotal (restricted to 35% if necessary)", S_WKCOL_3 & " " & lFinalPercentage - lFuelDeduction & "%")
        Call WKOut(rep, WK_BLANK_LINE)
        bSubtotal = True
      End If
    End If
    
    If ben.value(car_p46FuelTypeString) = "E" Then Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Electric car discount ", S_WKCOL_3 & "(" & p11d32.Rates.value(CarELECTRICDISCOUNT) * 100 & "%)", , , True)
    If ben.value(car_p46FuelTypeString) = "H" Then
      Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Hybird-Electric car discount ", S_WKCOL_3 & "(" & p11d32.Rates.value(CarHYBRIDELECTRICDISCOUNT) * 100 & "%)", , , True)
      Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Additional discount for CO2 emissions below qualifying level ", S_WKCOL_3 & "(" & lAdditionalDiscount & "%)", , , True)
    End If
    If ben.value(car_p46FuelTypeString) = "B" Then
      Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Road fuel gas car discount ", S_WKCOL_3 & "(" & p11d32.Rates.value(CarSTDDISCOUNTSUPPLEMENT) * 100 & "%)", , , True)
      Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Additional discount for CO2 emissions below qualifying level ", S_WKCOL_3 & "(" & lAdditionalDiscount & "%)", , , True)
    End If
    If ben.value(car_p46FuelTypeString) = "C" Then Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_ITALIC_FONT & "     Other bi-fuel car discount ", S_WKCOL_3 & "(" & p11d32.Rates.value(CarSTDDISCOUNTSUPPLEMENT) * 100 & "%)", , , True)
    
  End If
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_NORMAL_BOLD_FONT & "     Total percentage " & IIf(bSubtotal = True, "", "(restricted to 35% if necessary)"), S_WKCOL_3 & lFinalPercentage & "%")
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKOut(rep, WK_ITEM_TEXT, "Adjusted car list price multiplied by percentage ", RoundDownEx(ben.value(car_Price) * lFinalPercentage / 100, 0))
  Call WKOut(rep, WK_ITEM_TEXT, "Reduction for days when the car was unavailable", ben.value(car_cardaysreduction), , , True)
  Call WKOut(rep, WK_ITEM_TEXT, "Payment from the employee for private use of the car", ben.value(car_ActualAmountMadeGood), , , True)
 
  'Call WKOut(rep, WK_ITEM_Total, IIf(lTotalUsers > 1, "Assessable car benefit per employee", "Assessable car benefit"), ben.value(car_carBenefit), , True)
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
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKOut(rep, WK_ITEM_TEXT, "Fuel benefit (percentage multiplied by £" & FormatNumber(p11d32.Rates.value(CarFUELBENSETFIGURE), 0, , , vbTrue) & "):", (ben.value(car_FuelBenefit) * lTotalUsers) + ben.value(car_fueldaysreduction), , True)
  Call WKOut(rep, WK_ITEM_TEXT, "Reduction for days when car was unavailable", ben.value(car_fueldaysreduction), , , True)
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
    .BenDataLinkDataType(bc, car_accessories) = TYPE_LONG
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
    
    
    'MMField sizes
    .BenDataLinkUDMDescription(bc, car_Make_db) = "Make"
    .BenDataLinkUDMDescription(bc, car_Model_db) = "Model"
    .BenDataLinkUDMDescription(bc, car_ListPrice_db) = "List Price"
    .BenDataLinkUDMDescription(bc, car_NumberOfUsers_db) = "Number of users"
    .BenDataLinkUDMDescription(bc, car_accessories) = "Accessories"
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
    
        
    
    
          
        
    .BenDataLinkUDMDescription(bc, car_p46PaymentFrequencyString) = "Payment frequency"
    .BenDataLinkUDMDescription(bc, car_p46FuelTypeStringLong) = "Type of fuel" 'RH
    .BenDataLinkUDMDescription(bc, car_p46CarbonDioxide_db) = "Grams of Carbon dioxide"
    .BenDataLinkUDMDescription(bc, car_p46NoApprovedCO2Figure_db) = "No CO2 Figure available" 'RH
    
    .BenDataLinkUDMDescription(bc, Car_HasFuelUnavailableDays_db) = S_COMPANY_CAR_DAYS_UNAVAILABLE_FUEL_DESCRIPTION
    
    
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
'AM To be removed  ben.value(car_BusMilereduction) = S_ERROR
'MP DB  ben.value(car_MileScaled) = S_ERROR
'AM To be removed  ben.value(car_Agereduction) = S_ERROR
  ben.value(car_accessories) = S_ERROR
'AM  ben.value(car_AgeScaled) = S_ERROR
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
  ben.value(car_Benefit_Reportable) = S_ERROR
  ben.value(car_FuelBenefit_Reportable) = S_ERROR
  ben.value(car_FuelClass1AAdjustment) = 0
  ben.value(car_unrestrictedPercentage) = 0
End Sub
Private Sub IBenefitClass_Kill()
  Set Fuel = Nothing
  Set m_Parent = Nothing
End Sub
Private Function ReadyReckonerPercentage() As Double
'includes support for 12.25 notes 2
  Dim ben As IBenefitClass
  Dim iCO2 As Long
  Dim d As Double, d1 As Double
  Dim iBase As Long
  Dim iSection122Notes2Level As Long
  Set ben = Me
  
  d = ben.value(car_p46CarbonDioxide_db)
  d = d / 5
  iCO2 = Int(d) * 5
  
  '20032004
  iSection122Notes2Level = 135 'don't know if this changes for further years
  
  If (p11d32.AppYear = 2003) Then
    iBase = 155
  ElseIf (p11d32.AppYear = 2004) Then
    iBase = 145
  ElseIf (p11d32.AppYear = 2005) Then
    iBase = 140
  Else
    Call Err.Raise(ERR_CAR_CO2, "ReadyReckonerPercentage", "Invlaid year")
  End If
  
  If iCO2 = 0 Then 'lowest value is considered to be the base value
    iCO2 = iBase
  End If
  
  d = (15 + ((Max(0, iCO2 - iBase)) / 5)) / 100
    
  d = Min(d, 0.35)
  'paragraph 12.27 Notes:2, or workign papers 5a stage 3
  'it seems to be only h & B fuel types according to the working papers ??
  'this is not elly part of the reckoner but it is the appropriate place
  Select Case ben.value(car_p46FuelType_db)
    Case CCFT_BIFUEL_WITH_CO2_FOR_GAS, CCFT_HYBRID, CCFT_GAS_ONLY
      iCO2 = ben.value(car_p46CarbonDioxide_db)
      If (iCO2 <= iSection122Notes2Level) Then
        iCO2 = iBase - iCO2
        d1 = iCO2 / 20
        iCO2 = Int(d1)
        d = d - (iCO2 * 0.01)
      End If
  End Select
  
  ReadyReckonerPercentage = d

End Function
Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Private Sub CalculateFuelBen(ByVal BenRate As Double)
  Dim ben As IBenefitClass
  Dim dFuelDateFactor As Variant
  Dim l As Long
  Dim dFuelDateTo As Date, dFuelDateFrom As Date
  
  On Error GoTo CalculateFuelBen_Err
  Call xSet("CalculateFuelBen")
  
  Set ben = Me
  
  With ben
  
  .value(car_FuelUnavailableDays) = IIf(m_BenItems(Car_HasFuelUnavailableDays_db), m_BenItems(car_UnavailableDays_db), 0)
  'fuel date factor
  
    dFuelDateFrom = .value(Car_AvailableFrom_db)
    dFuelDateTo = .value(Car_AvailableTo_db)
    If .value(Car_FuelAvailableTo_db) < dFuelDateTo Then
      If Not .value(car_fuelreinstated_db) Then dFuelDateTo = .value(Car_FuelAvailableTo_db)
    End If
    dFuelDateFactor = dGetDateFactor(l, dFuelDateFrom, dFuelDateTo, .value(car_FuelUnavailableDays), L_DAYS_UNAVAIL_MIN, False)
    
    'calculating fuel VAT benefit according to engine size and petrol type
    If .value(car_privatefuel_db) Then
      If (.value(car_p46FuelType_db) = CCFT_DIESEL) Or (.value(car_p46FuelType_db) = CCFT_EUROIVDIESEL) Then
        Select Case .value(car_enginesize_db)
          Case Is > p11d32.Rates.value(carCCHIGH)
            .value(car_FuelVATBenefit) = p11d32.Rates.value(CarVATSCALEDIESELHIGH)
          Case Else
            .value(car_FuelVATBenefit) = p11d32.Rates.value(CarVATSCALEDIESELLOW)
          End Select
      Else
        Select Case .value(car_enginesize_db)
          Case Is > p11d32.Rates.value(carCCHIGH)
            .value(car_FuelVATBenefit) = p11d32.Rates.value(CarVATSCALEPETROLHIGH)
          Case Is > p11d32.Rates.value(carCCLOW)
            .value(car_FuelVATBenefit) = p11d32.Rates.value(CarVATSCALEPETROLMED)
          Case Else
            .value(car_FuelVATBenefit) = p11d32.Rates.value(CarVATSCALEPETROLLOW)
        End Select
      End If
      
      'Calculating Fuel benefit (=  set figure (£14,400 for 2003) * benefit rate)
      .value(car_GrossFuel) = p11d32.Rates.value(CarFUELBENSETFIGURE) * BenRate
 
      'RH - No Fuel Benefit if Electricity only
      If .value(car_requiredmakegood_db) And .value(car_actualmadegood_db) Or .value(car_p46FuelType_db) = CCFT_ELECTRIC Then
        .value(car_FuelMadeGood) = (.value(car_GrossFuel) * dFuelDateFactor) / .value(car_NumberOfUsers_db)
        .value(car_GrossFuel) = 0
        .value(car_FuelVATBenefit) = 0
      Else
        .value(car_FuelMadeGood) = 0
      End If

    Else
      .value(car_GrossFuel) = 0
      .value(car_FuelVATBenefit) = 0
      .value(car_FuelMadeGood) = 0
      
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
      
    .value(car_fueldaysreduction) = .value(car_GrossFuel) * (1 - dFuelDateFactor)
    
    .value(car_FuelBenefit) = RoundDownEx((.value(car_GrossFuel) * dFuelDateFactor) / (.value(car_NumberOfUsers_db)), 0)
    ben.value(car_FuelValue) = .value(car_FuelBenefit)
    .value(car_FuelVATBenefit) = (.value(car_FuelVATBenefit) * dFuelDateFactor) / .value(car_NumberOfUsers_db)
    .value(car_FuelVAT) = RoundN(.value(car_FuelVATBenefit) * (7 / 47))
    
    .value(car_FuelMadeGood_NET) = Max(0, .value(car_FuelMadeGood))
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
  Dim s As String
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
'MP DB ToDo - ok? - removed StandardReadData(Fuel) as Fuel Enum does not include any Standard BenItems
'MP DB    Call StandardReadData(Fuel)
       
    ben.value(car_Registration_db) = IIf(IsNull(rs.Fields("Reg").value), "", rs.Fields("Reg").value)
    ben.value(car_Make_db) = "" & rs.Fields("Make").value
    ben.value(car_Model_db) = IIf(IsNull(rs.Fields("Model").value), "", rs.Fields("Model").value)
    
    'P46 replacement
    ben.value(car_RegReplaced_db) = IIf(IsNull(rs.Fields("RegReplaced").value), "", rs.Fields("RegReplaced").value)
    ben.value(car_CarReplaced_db) = IIf(IsNull(rs.Fields("CarReplaced").value), "", rs.Fields("CarReplaced").value)
'MP DB - car_dateReplaced read below with Undated
'MP DB    ben.value(car_dateReplaced) = IIf(IsNull(rs.Fields("DateCarReplaced").value), "", rs.Fields("DateCarReplaced").value)
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
    ben.value(car_privatefuel_db) = rs.Fields("PvtFuel").value
    ben.value(car_requiredmakegood_db) = rs.Fields("MakeGood").value
    ben.value(car_actualmadegood_db) = rs.Fields("MadeGood").value
    ben.value(car_ForceP46_db) = rs.Fields("ForceP46").value
        
    ben.value(car_Registrationdate_db) = IIf(IsNull(rs.Fields("RegDate").value), p11d32.Rates.value(CarRegDateDef), rs.Fields("RegDate").value)
    ben.value(Car_AvailableFrom_db) = IIf(IsNull(rs.Fields("AvailFrom").value), p11d32.Rates.value(TaxYearStart), rs.Fields("AvailFrom").value)
    
    ben.value(Car_AvailableTo_db) = IIf(IsNull(rs.Fields("AvailTo").value), p11d32.Rates.value(TaxYearEnd), rs.Fields("AvailTo").value)
    ben.value(car_dateReplaced_db) = IIf(IsNull(rs.Fields("DateCarReplaced").value), UNDATED, rs.Fields("DateCarReplaced").value)
    
    ben.value(Car_FuelAvailableTo_db) = IIf(IsNull(rs.Fields("FuelAvailTo").value), ben.value(Car_AvailableTo_db), rs.Fields("FuelAvailTo").value)
    
        
    If p11d32.BringForward.Yes Then  'RH
      ben.value(car_NumberOfUsers_db) = 1
      ben.value(car_fuelreinstated_db) = False
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
    
    ben.value(car_p46FuelType_db) = IIf(IsNull(rs.Fields("P46FuelType").value), 0, rs.Fields("P46FuelType").value)
    ben.value(car_p46CarbonDioxide_db) = IIf(IsNull(rs.Fields("P46CarbonDioxide").value), 0, rs.Fields("P46CarbonDioxide").value)
    ben.value(car_p46NoApprovedCO2Figure_db) = IIf(IsNull(rs.Fields("P46NoApprovedCO2Figure").value), False, rs.Fields("P46NoApprovedCO2Figure").value)
    
    'just incase the value > 3 in length
    iLen = p11d32.BenDataLinkMMFieldSize(BC_COMPANY_CARS_F, car_p46CarbonDioxide_db)
    If (Len(CStr(ben.value(car_p46CarbonDioxide_db))) > iLen) Then
      ben.value(car_p46CarbonDioxide_db) = Right$(CStr(ben.value(car_p46CarbonDioxide_db)), iLen)
    End If
    
    If (ben.value(car_p46FuelType_db) = CCFT_BIFUELOLD) Then
      ben.value(car_p46FuelType_db) = CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B
    End If
    m_ReadFromDB = True
    Call StandardReadData(Fuel)
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
  If Item = 24 Then
    Item = 24
  End If
  m_BenItems(Item) = CorrectBenValue(BC_COMPANY_CARS_F, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo CompanyCar_WriteDB_Err
  Call xSet("CompanyCar_WriteDB")
  
  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields("Reg").value = ben.value(car_Registration_db)
  rs.Fields("Make").value = ben.value(car_Make_db)
  rs.Fields("Model").value = ben.value(car_Model_db)
  rs.Fields("Price").value = ben.value(car_ListPrice_db)
  rs.Fields("CapContrib").value = ben.value(car_capitalcontribution_db)
  rs.Fields("cc").value = ben.value(car_enginesize_db)
  rs.Fields("Acc").value = ben.value(car_AccessoriesOriginal_db)
  rs.Fields("NewAcc").value = ben.value(car_AccessoriesNew_db)
  rs.Fields("CheapAcc").value = ben.value(car_CheapAccessories_db)
  
  rs.Fields("PvtFuel").value = ben.value(car_privatefuel_db)
'AM To be removed  rs.Fields("Diesel") = ben.value(car_Diesel)
  rs.Fields("MakeGood").value = ben.value(car_requiredmakegood_db)
  rs.Fields("MadeGood").value = ben.value(car_actualmadegood_db)
  rs.Fields("RegDate").value = IIf(TryConvertDateDMY(CStr(ben.value(car_Registrationdate_db))) = UNDATED, Null, ben.value(car_Registrationdate_db))
  rs.Fields("SecondCar").value = ben.value(car_Second_db)

  'km 12/06/02 - this field no longer exists in the 2002 db
  'rs.Fields("BusMilesActual") = ben.value(car_UseActualMiles)
  rs.Fields("P46PaymentFrequency").value = ben.value(car_p46PaymentFrequency_db) 'so
  rs.Fields("P46FuelType").value = ben.value(car_p46FuelType_db) 'so
  rs.Fields("P46CarbonDioxide").value = ben.value(car_p46CarbonDioxide_db) 'so
  rs.Fields("P46NoApprovedCO2Figure").value = ben.value(car_p46NoApprovedCO2Figure_db) 'so
  'rs.Fields("FuelUnavail").value = ben.value(car_FuelUnavailableDays)

  
  Call StandardWriteData(ben, rs)
  
  Call BringForwardDatesWrite(ben, Car_AvailableFrom_db, Car_AvailableTo_db, rs, "AvailFrom", "AvailTo")
  
  If p11d32.BringForward.Yes Then
    rs.Fields("FuelAvailTo") = rs.Fields("AvailTo")
    rs.Fields("P46PaymentFrequency").value = ben.value(car_p46PaymentFrequency_db)
    rs.Fields("NumOfUsers").value = 1 'ben.value(car_NumberOfUsers)
    rs.Fields("FuelMadeGoodIsTaxDeducted").value = False
  Else
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
'MP DB (repeated)   rs.Fields("Replaced").value = ben.value(car_Replaced)
    rs.Fields("ForceP46").value = ben.value(car_ForceP46_db)
    rs.Fields("FuelAvailTo").value = ben.value(Car_FuelAvailableTo_db)
  End If
 
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
'      rs.Fields("Reg").value = ben.value(car_Registration)
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


Public Function GetPayment(Optional DateFactor As Variant) As Variant
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
    
  GetPayment = vPayment
  
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

  On Error GoTo SetP46FuelTypeString_ERR
  Call xSet("SetP46FuelTypeString")
  
  Set ben = Me
  With ben
    Select Case ben.value(car_p46FuelType_db)
      Case CCFT_PETROL
        .value(car_p46FuelTypeString) = "P"
        .value(car_p46FuelTypeStringLong) = "Petrol"  'RH
      Case CCFT_DIESEL
        .value(car_p46FuelTypeString) = "D"
        .value(car_p46FuelTypeStringLong) = "Diesel"  'RH
      Case CCFT_HYBRID
        .value(car_p46FuelTypeString) = "H"
        .value(car_p46FuelTypeStringLong) = "Hybrid electric" 'RH
      Case CCFT_ELECTRIC
        .value(car_p46FuelTypeString) = "E"
        .value(car_p46FuelTypeStringLong) = "Electricity only" 'RH
      Case CCFT_BIFUEL_WITH_CO2_FOR_GAS
        .value(car_p46FuelTypeString) = "B"
        .value(car_p46FuelTypeStringLong) = "Bi-fuel" 'RH 'THIS IS BIFUEL
      'Case CCFT_BIFUELOLD 'redundant but still needs to be kept for viewing, CAD 20032004
      '  .value(car_p46FuelTypeString) = "C"
      '  .value(car_p46FuelTypeStringLong) = "Conversion or older bi-fuel"
      Case CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B
        .value(car_p46FuelTypeString) = "C"
        .value(car_p46FuelTypeStringLong) = "Bi-fuel conv/other bi-fuel not within B"
      Case CCFT_EUROIVDIESEL
        .value(car_p46FuelTypeString) = "L"
        .value(car_p46FuelTypeStringLong) = "Euro IV Diesel" 'AM
      Case CCFT_GAS_ONLY
        .value(car_p46FuelTypeString) = "B"
        .value(car_p46FuelTypeStringLong) = "Gas Only"
    End Select
  End With

SetP46FuelTypeString_END:
  Call xReturn("SetP46FuelTypeString")
  Exit Function
SetP46FuelTypeString_ERR:
  Resume SetP46FuelTypeString_END
  Resume


End Function

