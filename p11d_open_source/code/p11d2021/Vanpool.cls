VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NonSharedVans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum NonSharedVansItems
'MP DB - these enum elements not db related
  nsvans_Item = ITEM_DESC 'RW
  nsvans_Value = ITEM_VALUE
  nsvans_MadeGood_Net = ITEM_MADEGOOD_NET
  nsvans_Benefit = ITEM_BENEFIT 'W
  nsvans_MadeGood = ITEM_MADEGOOD
  nsvans_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  nsvans_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  nsvans_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  nsvans_BoxNumber = ITEM_BOX_NUMBER
  nsvans_MadeGoodIsTaxDeducted = ITEM_MADEGOOD_IS_TAXDEDUCTED
  nsvans_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  nsvans_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  nsvans_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  nsvans_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  nsvans_Error = ITEM_ERROR
  nsvans_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  nsvans_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  nsvans_OPRA_Ammount_Foregone = ITEM_OPRA_AMOUNT_FOREGONE 'not written out as we store on the employee, but the display values are used here
  nsvans_FuelOPRA_Ammount_Foregone_Used_For_Value
  nsvans_FuelOPRA_Ammount_Foregone = ITEM_OPRA_AMOUNT_FOREGONE 'not written out as we store on the employee but the display values are used here

  nsvans_MADEGOOD_NET_FOR_CLASS1A_CALC
  nsvans_benefit_van_only
  nsvans_fuel_benefit
  nsvans_LASTITEM = nsvan_fuel_benefit
  
  nsvan
End Enum

Private m_BenItems(1 To nsvans_LASTITEM)

Public Vans As ObjectList
Private m_BenClass  As BEN_CLASS
Private m_Parent As Employee
Private m_ReadFromDB As Boolean
Private m_dirty As Boolean
Private m_InvalidFields As Long
Private m_NeedToCalculate As Boolean


Private Sub Class_Terminate()
 'Debug.Print"Non shared vans - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  
End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Other"
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
  IBenefitClass_CanBringForward = True 'let nsvan deal with this
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function

Private Sub Class_Initialize()
  Set Vans = New ObjectList
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_NONSHAREDVANS_G
End Property
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Private Function IBenefitClass_Calculate() As Variant
  Dim vSharedVansBenefit As Variant
  Dim ben As IBenefitClass
  Dim benMe As IBenefitClass
  Dim i As Long
  Dim vNonSharedVansBenefit As Variant
  Dim vNonSharedVansValue As Variant
  Dim vNonSharedVansVanValue As Variant
  Dim vNonSharedVansFuelValue As Variant
  Dim vNonSharedVansClass1AAdjustment As Variant
  Dim vNonSharedVansBenefitVanOnly As Variant
  Dim vNonSharedVansBenefitFuel As Variant
  Dim vOPRAValueTotalVans As Variant
  
  Dim vTotalMadeGood As Variant

  Dim vSharedVansBenefitVanOnly As Variant
  Dim vSharedVansBenefitFuel As Variant
  Dim vOPRAMadeGood As Variant

  Dim vSharedVansMadeGood As Variant
  Dim vSharedVansValueVanOnly As Variant
  Dim vSharedVansValueFuel As Variant
  Dim nonSharedVans As VanData
  Dim allVans As VanData
  Dim SharedVans As VanData
  Dim opraVans As VanData
  Dim calcVans As VanData
  
  Set nonSharedVans = New VanData
  Set allVans = New VanData
  Set SharedVans = New VanData
  Set opraVans = New VanData
  

  On Error GoTo NonSharedVans_Calculate_ERR
  
  Call xSet("NonSharedVans_Calculate")
  
  Set benMe = Me
  
  vNonSharedVansBenefit = 0
  vNonSharedVansClass1AAdjustment = 0
  vNonSharedVansBenefitFuel = 0
  vNonSharedVansBenefitVanOnly = 0
  vNonSharedVansVanValue = 0
  vNonSharedVansFuelValue = 0
  
  vSharedVansBenefitFuel = 0
  vSharedVansBenefitVanOnly = 0
  
  IBenefitClass_Calculate = S_ERROR
  
  benMe.value(nsvans_MadeGood) = 0
  benMe.value(nsvans_MadeGood_Net) = 0
  benMe.value(nsvans_Value) = 0
  benMe.value(nsvans_Benefit_Reportable) = False
  benMe.value(nsvans_Benefit) = S_ERROR
  benMe.value(ITEM_DESC) = p11d32.Rates.BenClassTo(BC_NONSHAREDVANS_G, BCT_FORM_CAPTION)
  benMe.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) = False
  benMe.value(nsvans_MADEGOOD_NET_FOR_CLASS1A_CALC) = 0
     
  benMe.value(nsvans_fuel_benefit) = 0
  benMe.value(nsvans_benefit_van_only) = 0
  
  benMe.value(nsvans_MadeGood) = 0

  
  With p11d32.CurrentEmployer
    'NON SHARED VANS
    For i = 1 To Vans.Count
      Set ben = Vans(i)
      If Not ben Is Nothing Then
        vNonSharedVansBenefit = vNonSharedVansBenefit + ben.Calculate
        vNonSharedVansBenefitFuel = vNonSharedVansBenefitFuel + ben.value(nsvan_fuel_benefit)
        vNonSharedVansBenefitVanOnly = vNonSharedVansBenefitVanOnly + ben.value(nsvan_benefit_van_only)
                
        nonSharedVans.VanValue = nonSharedVans.VanValue + ben.value(nsvan_Value_Non_OPRA)
        nonSharedVans.FuelValue = nonSharedVans.FuelValue + ben.value(nsvan_fuel_value)
        nonSharedVans.MadeGood = nonSharedVans.MadeGood + ben.value(nsvan_madegood_db)
        
        
        benMe.value(nsvans_MadeGood) = benMe.value(nsvans_MadeGood) + ben.value(nsvan_madegood_db)
        benMe.value(nsvans_MadeGood_Net) = benMe.value(nsvans_MadeGood_Net) + ben.value(nsvan_MadeGood_Net)
        benMe.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) = benMe.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) Or ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED)
        If ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) Then
          benMe.value(nsvans_MADEGOOD_NET_FOR_CLASS1A_CALC) = benMe.value(nsvans_MADEGOOD_NET_FOR_CLASS1A_CALC) + ben.value(nsvan_MadeGood_Net)
        End If
      End If
    Next
    
    'SHARED VANS
    Set ben = .SharedVans
    vSharedVansBenefit = ben.Calculate
    vSharedVansBenefitFuel = ben.value(shvan_fuel_benefit)
    
    SharedVans.VanValue = ben.value(shvan_Value_Non_OPRA)
    SharedVans.FuelValue = ben.value(shvan_fuel_value)
    
        
    vSharedVansValueVanOnly = ben.value(shvan_Value_Non_OPRA)
    vSharedVansValueFuel = ben.value(shvan_fuel_value)
    
    
    vSharedVansBenefitVanOnly = ben.value(shvan_benefit_van_only)
    
    'TOTAL SHARED VANS VALUE (VAN AND FUEL) - see SharedVans Class
    Set ben = m_Parent
    
    SharedVans.MadeGood = ben.value(ee_PaymentsForPrivateUseOfSharedVans_db)
    
    allVans.FuelValue = SharedVans.FuelValue + nonSharedVans.FuelValue
    allVans.VanValue = SharedVans.VanValue + nonSharedVans.VanValue
    allVans.MadeGood = SharedVans.MadeGood + nonSharedVans.MadeGood
    opraVans.VanValue = ben.value(ee_VansOPRAAmountForegone_db)
    opraVans.FuelValue = ben.value(ee_VansFuelOPRAAmountForegone_db)
    
    
    
    
    
    
    
    If vSharedVansMadeGood > 0 Then
      vSharedVansBenefit = Max(vSharedVansBenefit - vSharedVansMadeGood, 0)
      vSharedVansBenefitVanOnly = vSharedVansBenefitVanOnly - vSharedVansMadeGood
      If (vSharedVansBenefitVanOnly < 0) Then
        vSharedVansBenefitFuel = vSharedVansBenefitFuel + vSharedVansBenefitVanOnly
      End If
      vSharedVansBenefitFuel = Max(0, vSharedVansBenefitFuel)
      vSharedVansBenefitVanOnly = Max(0, vSharedVansBenefitVanOnly)
    End If
    
    'think we will need to store for the working paper!
    
    'THE FINAL CALC
    If ben.value(ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db) Then
      If ben.value(ee_OneOrMoreSharedVanAvailable_db) Then
        opraVans.MadeGood = SharedVans.MadeGood + nonSharedVans.MadeGood
        Call AssignVanValues(benMe, opraVans, allVans, vNonSharedVansBenefit + vSharedVansBenefit, vNonSharedVansBenefitFuel + vSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly + vSharedVansBenefitVanOnly)
        'Call AssignVanValuesVan(benMe, opraVans, allVans, vNonSharedVansBenefit + vSharedVansBenefit, vNonSharedVansBenefitFuel + vSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly + vSharedVansBenefitVanOnly)
        'opraVans.VanValue = SharedVans.VanValue + nonSharedVans.VanValue
        'Call AssignVanValuesFuel(benMe, opraVans, allVans, vNonSharedVansBenefit + vSharedVansBenefit, vNonSharedVansBenefitFuel + vSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly + vSharedVansBenefitVanOnly)
         
      Else
        opraVans.MadeGood = nonSharedVans.MadeGood
        Call AssignVanValues(benMe, opraVans, allVans, vNonSharedVansBenefit, vNonSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly)
      End If
    Else
      If ben.value(ee_OneOrMoreSharedVanAvailable_db) Then
        opraVans.MadeGood = SharedVans.MadeGood + nonSharedVans.MadeGood
        If vNonSharedVansBenefit > vSharedVansBenefit Then
          Call AssignVanValues(benMe, opraVans, nonSharedVans, vNonSharedVansBenefit, vNonSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly)
          
        Else
          Call AssignVanValues(benMe, opraVans, SharedVans, vSharedVansBenefit, vSharedVansBenefitFuel, vSharedVansBenefitVanOnly)
        End If
      Else
        opraVans.MadeGood = nonSharedVans.MadeGood
        
        
        Call AssignVanValues(benMe, opraVans, nonSharedVans, vNonSharedVansBenefit, vNonSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly)
        'Call AssignVanValuesVan(benMe, opraVans, allVans, vNonSharedVansBenefit + vSharedVansBenefit, vNonSharedVansBenefitFuel + vSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly + vSharedVansBenefitVanOnly)
        'opraVans.VanValue = nonSharedVans.VanValue
        'Call AssignVanValuesFuel(benMe, opraVans, allVans, vNonSharedVansBenefit + vSharedVansBenefit, vNonSharedVansBenefitFuel + vSharedVansBenefitFuel, vNonSharedVansBenefitVanOnly + vSharedVansBenefitVanOnly)
        
      End If
    End If
    
    benMe.value(nsvans_Value) = benMe.value(nsvans_Benefit) + benMe.value(nsvans_MadeGood_Net)

    IBenefitClass_Calculate = benMe.value(nsvans_Benefit)
    Call BenCalcNIC(benMe, , , , , , , nsvans_MADEGOOD_NET_FOR_CLASS1A_CALC)
    
    If benMe.value(nsvans_Benefit) > 0 Then benMe.value(nsvans_Benefit_Reportable) = True
  End With
  
NonSharedVans_Calculate_END:
  Call xReturn("NonSharedVans_Calculate")
  Exit Function
NonSharedVans_Calculate_ERR:
  benMe.value(nsvans_Benefit_Reportable) = True
  
  IBenefitClass_Calculate = S_ERROR
  Resume NonSharedVans_Calculate_END
  Resume
End Function

Private Sub AssignVanValuesFuel(benMe As IBenefitClass, opraVans As VanData, vansCompare As VanData, ByVal vBenefitStandardAll As Variant, ByVal vBenefitStandardFuel As Variant, ByVal vBenefitStandardVanOnly As Variant)
  Dim vTotalBenefit As Variant
  
  If (opraVans.TotalFuelValue() > vansCompare.TotalFuelValue()) Then
    benMe.value(nsvans_FuelOPRA_Ammount_Foregone_Used_For_Value) = True 'need another value here
    benMe.value(nsvans_fuel_benefit) = opraVans.FuelBenefit()
    
  Else
    benMe.value(nsvans_fuel_benefit) = vBenefitStandardFuel
  End If
  
  'total benfit, always do
  benMe.value(nsvans_Benefit) = benMe.value(nsvans_benefit_van_only) + benMe.value(nsvans_fuel_benefit)


  
End Sub
Private Sub AssignVanValuesVan(benMe As IBenefitClass, opraVans As VanData, vansCompare As VanData, ByVal vBenefitStandardAll As Variant, ByVal vBenefitStandardFuel As Variant, ByVal vBenefitStandardVanOnly As Variant)
  Dim vTotalBenefit As Variant
  
  If (opraVans.TotalVanValue() > vansCompare.TotalVanValue()) Then
    benMe.value(nsvans_OPRA_Ammount_Foregone_Used_For_Value) = True
    benMe.value(nsvans_benefit_van_only) = opraVans.VanBenefit()
  Else
    benMe.value(nsvans_benefit_van_only) = vBenefitStandardVanOnly
  End If


  'total benfit
  benMe.value(nsvans_Benefit) = benMe.value(nsvans_benefit_van_only) + benMe.value(nsvans_fuel_benefit)


  
End Sub



Private Sub AssignVanValues(benMe As IBenefitClass, opraVans As VanData, vansCompare As VanData, ByVal vBenefitStandardAll As Variant, ByVal vBenefitStandardFuel As Variant, ByVal vBenefitStandardVanOnly As Variant)
  Dim vTotalBenefit As Variant
  
  If (opraVans.TotalVanValue() > vansCompare.TotalVanValue()) Then
    benMe.value(nsvans_OPRA_Ammount_Foregone_Used_For_Value) = True
    benMe.value(nsvans_benefit_van_only) = opraVans.VanBenefit()
  Else
    benMe.value(nsvans_benefit_van_only) = vBenefitStandardVanOnly
  End If

  If (opraVans.TotalFuelValue() > vansCompare.TotalFuelValue()) Then
    benMe.value(nsvans_FuelOPRA_Ammount_Foregone_Used_For_Value) = True 'need another value here
    benMe.value(nsvans_fuel_benefit) = opraVans.FuelBenefit()
    
  Else
    benMe.value(nsvans_fuel_benefit) = vBenefitStandardFuel
  End If
  
  'total benfit
  benMe.value(nsvans_Benefit) = benMe.value(nsvans_benefit_van_only) + benMe.value(nsvans_fuel_benefit)


  'If (opraVans.TotalValue() > vansCompare.TotalValue()) Then
  '  benMe.value(nsvans_Benefit) = opraVans.TotalBenefit()
  '  benMe.value(nsvans_fuel_benefit) = opraVans.FuelBenefit()
  '  benMe.value(nsvans_benefit_van_only) = opraVans.VanBenefit()
  '  benMe.value(nsvans_FuelOPRA_Ammount_Foregone_Used_For_Value) = True
  '  benMe.value(nsvans_OPRA_Ammount_Foregone_Used_For_Value) = True
  'Else
  '  benMe.value(nsvans_Benefit) = vBenefitStandardAll
  '  benMe.value(nsvans_fuel_benefit) = vBenefitStandardFuel
  '  benMe.value(nsvans_benefit_van_only) = vBenefitStandardVanOnly
  'End If
End Sub

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "clsVanPool_CompanyDefined"
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
End Property
Private Function IBenefitClass_DeleteDB() As Boolean
  Dim i As Long
  Dim ben As IBenefitClass
  
  For i = 1 To Vans.Count
    Set ben = Vans(i)
    If Not ben Is Nothing Then
      Call ben.DeleteDB
      Set ben = Nothing
    End If
  Next
  
  Set ben = Nothing
End Function

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  Dim benParent As IBenefitClass
  m_dirty = DirtyHelper(Me, NewValue)
  
  Set benParent = m_Parent
  

  benParent.Dirty = benParent.Dirty Or m_dirty
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass, benNsVan As IBenefitClass, ee As IBenefitClass
  Dim ey As IBenefitClass, ey2 As Employer, nsvanBen As IBenefitClass, ee2 As Employee
  Dim sharedvansben As IBenefitClass
  Dim v As Variant
  Dim i As Long
  
  
  On Error GoTo NonSharedVans_PrintWKBody_Err
  Call xSet("NonSharedVans_PrintWKBody")
  
  Set ben = Me
  Set ee = m_Parent
  Set ee2 = ee
  Set ey = ee.Parent
  Set ey2 = ey
  
  'CAD all this needs a re-work of vans, calcs??
  
  '************************* shared vans *****************************************
  If ee.value(ee_OneOrMoreSharedVanAvailable_db) Then
    Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
    Call WKOut(rep, WK_ITEM_TEXT, "Was more than one van available for private use at any time in the tax year? " & S_WK_NORMAL_BOLD_FONT & p11d32.MagneticMedia.MMBool(ee.value(ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db)))
    Call WKOut(rep, WK_ITEM_TEXT, "Number of employees with shared vans: " & S_WK_NORMAL_BOLD_FONT & ey.value(employer_EmployeesWithSharedVan))
    Call WKOut(rep, WK_ITEM_TEXT, "Number shared vans: " & S_WK_NORMAL_BOLD_FONT & ey2.SharedVans.Vans.CountValid)
    Call WKOut(rep, WK_BLANK_LINE)
    'Call WKOut(rep, WK_SECTION_HEADER, "Shared van benefit 'normal method':") 'CAD 2008 removed daily calc
    Call WKOut(rep, WK_SECTION_HEADER, "Shared van benefit:")
    Call WKOut(rep, WK_BLANK_LINE)
    Set sharedvansben = ey2.SharedVans
    v = sharedvansben.Calculate
    Call WKOut(rep, WK_ITEM_TEXT, "Van benefit", sharedvansben.value(shvan_benefit_van_only), , True)
    Call WKOut(rep, WK_ITEM_TEXT, "Fuel benefit", sharedvansben.value(shvan_fuel_benefit), , True)
    Call WKOut(rep, WK_ITEM_Total, "Total van benefit", v, , True)
    Call WKOut(rep, WK_ITEM_TEXT, "Payments for private use of shared vans ", ee.value(ee_PaymentsForPrivateUseOfSharedVans_db), , True, True)
    Call WKOut(rep, WK_ITEM_Total, "Benefit ", Max(sharedvansben.Calculate - ee.value(ee_PaymentsForPrivateUseOfSharedVans_db), 0), , True)
      
    'CAD 2008 removed daily calc
    If False And ee.value(ee_ReportyDailyCalculationOfSharedVans_db) Then
      Call WKOut(rep, WK_BLANK_LINE)
      Call WKOut(rep, WK_SECTION_HEADER, "Shared van benefit 'alternative daily method (sch.6A para.8)':")
      
      Call WKOut(rep, WK_BLANK_LINE)
      Call WKOut(rep, WK_ITEM_TEXT, "Relvant number of days - " & ee.value(ee_RelevantDaysForDailySharedVanCalc_db))
      v = ee.value(ee_RelevantDaysForDailySharedVanCalc_db) * p11d32.Rates.value(vanDailyCharge)
      Call WKOut(rep, WK_ITEM_TEXT, "Daily calculation  - " & ee.value(ee_RelevantDaysForDailySharedVanCalc_db) & " x " & FormatWN(p11d32.Rates.value(vanDailyCharge)), v)
      Call WKOut(rep, WK_ITEM_TEXT, "Payments for private use of shared vans ", ee.value(ee_PaymentsForPrivateUseOfSharedVans_db), , , True)
      Call WKOut(rep, WK_ITEM_Total, "Benefit ", Max(0, v - ee.value(ee_PaymentsForPrivateUseOfSharedVans_db)), , True)
      Call WKOut(rep, WK_BLANK_LINE)
      Call WKOut(rep, WK_ITEM_TEXT, "Note: where the alternative method is beneficial, a claim should be made on your tax return under s.164 ITEPA 2003")
    End If
    
    Call WKOut(rep, WK_SECTION_BREAK)
  End If
  
  
  ' ************************************* non shared vans ****************************************
  If Vans.CountValid Then
    Call WKOut(rep, WK_SECTION_HEADER, "Private van benefit:")
    For i = 1 To Vans.Count
      Set nsvanBen = Vans(i)
      Call WKNonSharedVan(rep, nsvanBen)
    Next
  End If
  
  
  If (ben.value(nsvans_OPRA_Ammount_Foregone_Used_For_Value)) Then
    Call WKOut(rep, WK_ITEM_TEXT, "OpRA value is greater than the standard calculation's value, using OpRA value for calculation")
    Call WKOut(rep, WK_BLANK_LINE)
  End If
  
  
  
  Call WKOut(rep, WK_ITEM_TEXT, "Overall van benefit", ben.value(ITEM_BENEFIT), , True)
  
  'cad removed 2017
  If ee2.AnyVanBenefit And False Then
    Call WKOut(rep, WK_SECTION_BREAK)
    Call WKOut(rep, WK_ITEM_TEXT, "It is possible to elect for a daily method of calculation of the van benefit where this gives a lower figure.")
    Call WKOut(rep, WK_ITEM_TEXT, "This method is calculated by applying a daily value of " & FormatWN(p11d32.Rates.value(vanDailyCharge)) & " to the number of days the van is used privately.")
  End If
    
NonSharedVans_PrintWKBody_End:
  Call xReturn("NonSharedVans_PrintWKBody")
  Exit Function

NonSharedVans_PrintWKBody_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVans_PrintWKBody", "Non Shared Vans Print WK Body", "Error printing the non shared vans working paper")
  Resume NonSharedVans_PrintWKBody_End
  Resume
End Function
Private Sub WKNonSharedVan(rep As Reporter, van As IBenefitClass)
  If Not van Is Nothing Then
    Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
    
    Call WKOut(rep, WK_ITEM_TEXT, "Description: " & S_WK_NORMAL_BOLD_FONT & van.value(ITEM_DESC))
    Call WKOut(rep, WK_ITEM_TEXT, "Date of registration: " & S_WK_NORMAL_BOLD_FONT & van.value(nsvan_RegistrationDate_db))
    
    'Call WKOut(rep, WK_ITEM_TEXT, "First registered on or after " & p11d32.Rates.value(vanOldDate) & ": " & S_WK_NORMAL_BOLD_FONT & IIf(p11d32.Rates.value(vanOldDate) < van.value(nsvan_RegistrationDate_db), "True", "False"))
'    Call WKOut(rep, WK_ITEM_TEXT, "First registered on or after " & p11d32.Rates.value(vanOldDate) & " - " & van.value(nsvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear))
    Call WKOut(rep, WK_ITEM_TEXT, "Available from " & S_WK_NORMAL_BOLD_FONT & van.value(nsvan_AvailableFrom_db))
    Call WKOut(rep, WK_ITEM_TEXT, "Available to " & S_WK_NORMAL_BOLD_FONT & van.value(nsvan_Availableto_db))
    Call WKOut(rep, WK_ITEM_TEXT, "Qualifying days unavailable or shared: " & S_WK_NORMAL_BOLD_FONT & van.value(nsvan_DaysUnavailable_db))
    
    Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)
    
    Call WKOut(rep, WK_ITEM_TEXT, "Van benefit", van.value(nsvan_benefit_van_only), , True)
    Call WKOut(rep, WK_ITEM_TEXT, "Fuel benefit", van.value(nsvan_fuel_benefit), , True)
    
    Call WKOut(rep, WK_ITEM_Total, "Value", van.value(ITEM_VALUE), , True)
    Call WKOut(rep, WK_ITEM_TEXT, "Made good", van.value(ITEM_MADEGOOD_NET), , , True)
    Call WKOut(rep, WK_ITEM_Total, "Benefit", van.value(ITEM_BENEFIT), , True)
    Call WKOut(rep, WK_SECTION_BREAK)
  End If
End Sub

Private Sub IBenefitClass_SetCalcDefaults()
  
  If p11d32.DataLinkInitialised(BC_NONSHAREDVANS_G) Then Exit Sub
  
  Call SetStandardBenItemsDataTypes(BC_NONSHAREDVANS_G)
  Call SetStandardBenItemsMMFieldSize(BC_NONSHAREDVANS_G)
  m_BenItems(nsvans_FuelOPRA_Ammount_Foregone_Used_For_Value) = False
  
End Sub


Private Property Get IBenefitClass_HasBookMark() As Boolean
'cd  IBenefitClass_HasBookMark = (Len(m_sbookmark) > 0)
End Property

Private Sub IBenefitClass_Kill()
  'ECASE "clsVanPool_Kill"
End Sub
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(nsvans_Item)
End Property

Public Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Public Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Property Get IBenefitClass_PrintHeader() As String
  
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  IBenefitClass_PrintWk = PrintWKHelper(rep, Me)
End Function

Public Function IBenefitClass_ReadDB() As Long
  Dim ibc As IBenefitClass
  Dim van As IBenefitClass
  Dim rs As Recordset
  Dim i As Long
  
  On Error GoTo NonSharedVans_ReadDb_err
  xSet "NonSharedVans_ReadDb"
  
  Set ibc = Me
   
  If ibc.ReadFromDB Then GoTo NonSharedVans_ReadDb_end
  Call StandardReadData(ibc)
  
  Set rs = m_Parent.Parent.rsBenTables(TBL_VANS)
  rs.FindFirst ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
  Do While Not rs.NoMatch
    i = i + 1
    
    Set van = New NonSharedVan
    van.RSBookMark = rs.Bookmark
    Set van.Parent = Me
    Call van.ReadDB
    Call Me.Vans.Add(van)
    ibc.value(nsvans_Item) = van.value(ITEM_DESC)
    Set van = Nothing
    rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
  Loop
    
  ibc.ReadFromDB = True
  
  
NonSharedVans_ReadDb_end:
  IBenefitClass_ReadDB = True
  Set ibc = Nothing
  Set van = Nothing
  Set rs = Nothing
  Call xReturn("NonSharedVans_ReadDb")
  Exit Function
NonSharedVans_ReadDb_err:
  Call ErrorMessage(ERR_ERROR, Err, "clsVanPool_ReadDB", "ERR_VANPOOL_READDB", "Unable to read the van pool.")
  Resume NonSharedVans_ReadDb_end
  Resume
End Function


Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  ECASE "clsVanPool_sBookMark"
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  ECASE "clsVanPool_sBookMark"
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
    
  On Error GoTo SetBenItemsInformation_err
  bc = BC_NONSHAREDVANS_G
  m_NeedToCalculate = True
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end

  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  
  With p11d32
    Set ben = Me
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    
    .BenDataLinkDataType(bc, nsvans_benefit_van_only) = TYPE_LONG
    .BenDataLinkDataType(bc, nsvans_fuel_benefit) = TYPE_LONG
    .BenDataLinkDataType(bc, nsvans_MADEGOOD_NET_FOR_CLASS1A_CALC) = TYPE_LONG

    .BenDataLinkUDMDescription(bc, ITEM_DESC) = S_UDM_DESCRIPTION
    .BenDataLinkUDMDescription(bc, ITEM_UDM_BENEFIT_TITLE) = S_UDM_BENEFIT_TITLE
    .BenDataLinkUDMDescription(bc, ITEM_BOX_NUMBER) = S_UDM_BOX_NUMBER
    
    .BenDataLinkUDMDescription(bc, ITEM_NIC_CLASS1A_BENEFIT) = S_UDM_NIC_CLASS1A_BENEFIT
    
    .BenDataLinkUDMDescription(bc, ITEM_BENEFIT) = S_UDM_BENEFIT
    .BenDataLinkUDMDescription(bc, ITEM_MADEGOOD_NET) = S_UDM_MADE_GOOD_NET
    .BenDataLinkUDMDescription(bc, ITEM_VALUE) = S_UDM_VALUE

    .BenDataLinkUDMDescription(bc, nsvans_benefit_van_only) = "Van only benefit"
    .BenDataLinkUDMDescription(bc, nsvans_fuel_benefit) = "Fuel benefit"
 
  End With
  Call AbacusUDMData(bc)

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub


Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_NONSHAREDVANS_G, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo VanPool_WriteDB_Err
  
  Call xSet("WriteDB")
  
  Set ben = Me
  
  IBenefitClass_WriteDB = True
  
  If ben.Dirty Or p11d32.BringForward.Yes Then
    For i = 1 To Vans.Count
      Set ben = Vans(i)
      If Not (ben Is Nothing) Then
        If ben.InvalidFields = 0 Then
          If (ben.CanBringForward And p11d32.BringForward.Yes) Or ben.Dirty Then
            Call ben.WriteDB
          End If
        Else
          IBenefitClass_WriteDB = False
        End If
      End If
    Next i
  End If
  
  
  
VanPool_WriteDB_End:
  Call xReturn("WriteDB")
  Exit Function
VanPool_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ErrorMessage(ERR_ERROR, Err, "VanPool_WriteDB", "ERR_WRITEDB", "Error writing the company van pool to the database.")
  Resume VanPool_WriteDB_End
  Resume
End Function


Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_LASTITEM
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  'not used
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  Dim ben As IBenefitClass
  Dim i As Long
  
  For i = 1 To Vans.Count
    Set ben = Vans(i)
    If Not ben Is Nothing Then
      IBenefitClass_InvalidFields = IBenefitClass_InvalidFields + ben.InvalidFields
    End If
  Next
  
  Set ben = Nothing
  
End Property


Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property



