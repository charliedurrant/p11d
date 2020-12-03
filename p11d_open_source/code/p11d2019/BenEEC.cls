VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum EECarItems
  eecar_Item_db = ITEM_DESC
  eecar_Value = ITEM_VALUE
  eecar_MadeGood_Net = ITEM_MADEGOOD_NET
  eecar_Benefit = ITEM_BENEFIT
  eecar_madegood = ITEM_MADEGOOD
  eecar_actualamountmadegood = ITEM_ACTUALAMOUNTMADEGOOD
  eecar_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  eecar_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  eecar_BoxNumber = ITEM_BOX_NUMBER
  eecar_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  eecar_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  eecar_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  eecar_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  eecar_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  eecar_ERROR = ITEM_ERROR
  eecar_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  eecar_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE 'added so all benefits are the same evein though mileage doesn't have this
  eecar_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE 'added so all benefits are the same evein though mileage doesn't have this
  
  eecar_CarMadeGood_db = [_ITEM_LAST_ITEM] + 1
  'Strings
  
  'combo boxes
  eecar_FPCS_db 'bf - Now MARORS Index of FPCS scheme
'MP DB ToDo confirm not used in rpt as stated
'  eecar_FPCSName_db 'Display name for report wizard
  eecar_EngineSize_db 'bf
'MP DB  eecar_AverageIRRate_db 'bf
  eecar_UseCompanyCarScheme_db 'bf
  
  eecar_AmountReceived_db
  eecar_WKAlternativeSection
  
  'EK No longer needed 2003 eecar_AlternativeMethod
  
  eecar_WKValue
  'extras dialog /text boxes
  eecar_LumpSum_db
  eecar_HireCost_db
  eecar_HireCostMadeGood_db
  eecar_BenefitRepayableTaxable
  eecar_BenefitRepayableTaxable_Reportable
  eecar_MMSectionU_Reportable
  
  eecar_TotalMiles
  eecar_NoOfCars
  
  eecar_CarsWithAlternativeMethodFalse
  eecar_TotalMilesCarsAlternativeFalse
  
  eecar_TotalExtras
  'backwards compatibility
  eecar_CompanyFPCSValue 'Formerly FPCS Now MARORS
  eecar_IRFPCSValue 'Formerly FPCS Now MARORS
  eecar_carkey_db
'MP DB  eecar_PreErBenefit
  eecar_HireCostNet
  eecar_MMRevenuReliefRate
  'Lastitem
  eecar_lastitem = eecar_MMRevenuReliefRate
  
End Enum

Public mileage As ObjectList
Public FPCSBandings As ObjectList  'Formerly FPCS Now MARORS

Private m_BenClass As BEN_CLASS
Private m_ReadFromDB As Boolean
Private m_sbookmark As String
Private m_dirty As Boolean

Private m_Parent As Employee
Private m_InvalidFields As Long
Private m_BenItems(1 To eecar_lastitem)
Private Const S_SCHEME1 As String = "Scheme 1"
Private m_TopBand As FPCSBand 'Formerly FPCS Now MARORS

Public Enum FPCS_SORT 'Formerly FPCS Now MARORS
  FPCS_CC_ASC 'Formerly FPCS Now MARORS
  FPCS_MILES_ASC 'Formerly FPCS Now MARORS
End Enum
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Employee Car - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim vMileageValue As Variant
  Dim ben As IBenefitClass
  Dim benEmployer As IBenefitClass
  Dim FPCS As FPCS 'Formerly FPCS Now MARORS
  
  On Error GoTo EmployeeCar_CalculateBody_Err
  Call xSet("EmployeeCar_CalculateBody")
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  
  Call GetIRFPCSValue(vMileageValue) 'Formerly FPCS Now MARORS
  ben.value(eecar_IRFPCSValue) = vMileageValue 'Formerly FPCS Now MARORS
  vMileageValue = 0
  Set FPCS = GetFPCSMileageValue(vMileageValue, ben.value(eecar_TotalMiles), ben.value(eecar_EngineSize_db), ben.value(eecar_FPCS_db), 0) ' EK IR Rates no longer relevant
  ben.value(eecar_CompanyFPCSValue) = vMileageValue 'Formerly FPCS Now MARORS

  ben.value(eecar_HireCostNet) = Max(ben.value(eecar_HireCost_db) - ben.value(eecar_HireCostMadeGood_db), 0)
    
  If ben.value(eecar_UseCompanyCarScheme_db) Then
    ben.value(eecar_AmountReceived_db) = ben.value(eecar_CompanyFPCSValue) 'Formerly FPCS Now MARORS
  End If
  
  ben.value(eecar_WKValue) = ben.value(eecar_AmountReceived_db) + ben.value(eecar_HireCostNet) + ben.value(eecar_LumpSum_db)
  ben.value(eecar_Value) = Max(ben.value(eecar_WKValue) - ben.value(eecar_IRFPCSValue), 0) 'Formerly FPCS Now MARORS
  ben.value(eecar_madegood) = ben.value(eecar_CarMadeGood_db)
  ben.value(eecar_MadeGood_Net) = Min(ben.value(eecar_Value), ben.value(eecar_madegood))
  
  ben.value(eecar_Benefit) = ben.value(eecar_Value) - ben.value(eecar_MadeGood_Net)
  ben.value(eecar_BenefitRepayableTaxable) = -1 * Min(0, (ben.value(eecar_AmountReceived_db) + ben.value(eecar_HireCostNet) + ben.value(eecar_LumpSum_db)) - (ben.value(eecar_IRFPCSValue) + Min(ben.value(eecar_AmountReceived_db) + ben.value(eecar_LumpSum_db), ben.value(eecar_madegood)))) 'Formerly FPCS Now MARORS
  
  'EK - 2003 can now be negative with caveats
  If ben.value(eecar_Benefit) = 0 And ben.value(eecar_BenefitRepayableTaxable) > 0 Then
    ben.value(eecar_Benefit) = -1 * ben.value(eecar_BenefitRepayableTaxable)
  End If
  
  Call CalcEmployeeCarsWithAlternativMethodFalse
  
  If ben.value(eecar_CarsWithAlternativeMethodFalse) > 1 Then ben.value(eecar_WKAlternativeSection) = True
  
  If ben.value(eecar_Benefit) <> 0 Then
    ben.value(eecar_Benefit_Reportable) = True
  End If
  
  If ben.value(eecar_BenefitRepayableTaxable) > 0 Then
    ben.value(eecar_BenefitRepayableTaxable_Reportable) = True
  Else
    ben.value(eecar_BenefitRepayableTaxable) = 0
  End If
  Set benEmployer = p11d32.CurrentEmployer
  ben.value(eecar_MMSectionU_Reportable) = (ben.value(eecar_BenefitRepayableTaxable_Reportable) Or ben.value(eecar_Benefit_Reportable)) And (benEmployer.value(employer_TreatEmployeeCarsAsUnderApprovedFPCS_db) Or FPCS.Approved)  'Formerly FPCS Now MARORS
  'get MM revenu relief rate
  ben.value(eecar_MMRevenuReliefRate) = MMRevenueReliefRate()
  Call BenCalcNIC(ben)
  IBenefitClass_CalculateBody = ben.value(eecar_Benefit)
  
EmployeeCar_CalculateBody_End:
  Call xReturn("EmployeeCar_CalculateBody")
  Exit Function
EmployeeCar_CalculateBody_Err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume EmployeeCar_CalculateBody_End
  Resume
End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "EmployeeCar"
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
  IBenefitClass_CanBringForward = True
  'always roll forward, AJ confirmed
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim eecar As EmployeeCar
  
  
  Set eecar = New EmployeeCar
  Set IBenefitClass_Copy = CopyBenStandard(Parent, eecar, Me)
End Function

Public Property Get TopBand() As FPCSBand  'Formerly FPCS Now MARORS
  Set TopBand = m_TopBand
End Property
Private Sub Class_Initialize()
  Set FPCSBandings = New ObjectList  'Formerly FPCS Now MARORS
  Set mileage = New ObjectList
  Call IBenefitClass_SetBenItemsInformation
End Sub
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_EMPLOYEE_CAR_E
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
  IBenefitClass_Name = m_BenItems(eecar_Item_db)
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass
  Dim ee As Employee
  Dim i As Integer
  Dim lTotalMiles As Variant
  Dim lTotalPayments As Variant
  Dim lTotalBenefit As Variant
  
  On Error GoTo EECAR_PRINTWK_Err

' EK Complete structural changes to the way this working paper works, now cycles through each
' company car and then puts in a payments summary, comes of the legislative changes in 2003
' These mean that cars are pooled
  
  Call xSet("EECAR_PRINTWK")
  'Set ben = Me
  Set ee = m_Parent
  lTotalMiles = 0
  lTotalPayments = 0
  lTotalBenefit = 0
  
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = BC_EMPLOYEE_CAR_E Then
        Call WKEmployeeCarDetails(rep, ben)
        Call WKEmployeeCarPayments(rep, ben)
        lTotalMiles = lTotalMiles + ben.value(eecar_TotalMiles)
        lTotalPayments = lTotalPayments + (ben.value(eecar_WKValue) - 1 * ben.value(ITEM_MADEGOOD))
        lTotalBenefit = lTotalBenefit + ben.value(eecar_Benefit)
      End If
    End If
  Next
  
  Call WKOut(rep, WK_SECTION_BREAK)
  Call WKEmployeeCarMARORS(rep, lTotalBenefit, lTotalMiles, lTotalPayments)
 '         Call rep.Out(vbCrLf & vbCrLf)
 '         Call WKOut(rep, WK_ITEM_TEXT, "Where payment of mileage allowance results in an excess over the Inland Revenue's Fixed Profit Car Scheme (FPCS) rates,") 'Formerly FPCS Now MARORS
 '         Call WKOut(rep, WK_ITEM_TEXT, "the figure should be entered in box 1.32 of the employment pages of the self assessment tax return.")
  
EECAR_PRINTWK_End:
  Call xReturn("EECAR_PRINTWK")
  Exit Function

EECAR_PRINTWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EECAR_PRINTWK", "ERR_EECAR_PRINTWK", "Error printing the employee owned car working paper")
  Resume EECAR_PRINTWK_End
  Resume
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
  
  On Error GoTo SetBenItemsInformation_err
  bc = BC_EMPLOYEE_CAR_E
  m_NeedToCalculate = True
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  Call SetStandardBenItemsInformation(bc, Me)
  
  With p11d32
    
    '.BenDataLinkDataType(bc, eecar_FPCS) = TYPE_LONG 'Formerly FPCS Now MARORS
'MP DB    .BenDataLinkDataType(bc, eecar_FPCSName_db) = TYPE_STR
    .BenDataLinkDataType(bc, eecar_MMRevenuReliefRate) = TYPE_STR
    .BenDataLinkMMFieldSize(bc, eecar_MMRevenuReliefRate) = 1
    .BenDataLinkDataType(bc, eecar_CarsWithAlternativeMethodFalse) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_WKValue) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_IRFPCSValue) = TYPE_LONG 'Formerly FPCS Now MARORS
    .BenDataLinkDataType(bc, eecar_NoOfCars) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_TotalMilesCarsAlternativeFalse) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_WKAlternativeSection) = TYPE_BOOL
    .BenDataLinkDataType(bc, eecar_CompanyFPCSValue) = TYPE_LONG 'Formerly FPCS Now MARORS
    .BenDataLinkDataType(bc, eecar_HireCostNet) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_BenefitRepayableTaxable) = TYPE_LONG
    .BenDataLinkMMFieldSize(bc, eecar_BenefitRepayableTaxable) = MM_SFS_DATA
    .BenDataLinkDataType(bc, eecar_MMSectionU_Reportable) = TYPE_BOOL
    .BenDataLinkDataType(bc, eecar_carkey_db) = TYPE_LONG
'MP DB    .BenDataLinkDataType(bc, eecar_PreErBenefit) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_CompanyFPCSValue) = TYPE_LONG 'Formerly FPCS Now MARORS
    .BenDataLinkDataType(bc, eecar_CarMadeGood_db) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_EngineSize_db) = TYPE_LONG
    .BenDataLinkMMFieldSize(bc, eecar_EngineSize_db) = 4
    .BenDataLinkDataType(bc, eecar_AmountReceived_db) = TYPE_LONG
    .BenDataLinkMMFieldSize(bc, eecar_AmountReceived_db) = MM_SFS_DATA
    .BenDataLinkDataType(bc, eecar_LumpSum_db) = TYPE_LONG
    .BenDataLinkMMFieldSize(bc, eecar_LumpSum_db) = MM_SFS_DATA
    
    .BenDataLinkDataType(bc, eecar_HireCost_db) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_HireCostMadeGood_db) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_TotalMiles) = TYPE_LONG
    .BenDataLinkDataType(bc, eecar_TotalExtras) = TYPE_LONG
    'EK No longer needed 2003 .BenDataLinkDataType(bc, eecar_AlternativeMethod) = TYPE_BOOL
'MP DB    .BenDataLinkDataType(bc, eecar_AverageIRRate_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, eecar_UseCompanyCarScheme_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, eecar_BenefitRepayableTaxable_Reportable) = TYPE_BOOL
    
    .BenDataLinkUDMDescription(bc, eecar_EngineSize_db) = "cc"
    .BenDataLinkUDMDescription(bc, eecar_TotalMiles) = "Miles"
    '.BenDataLinkUDMDescription(bc, eecar_FPCS) = "Company scheme" 'Formerly FPCS Now MARORS
'MP DB    .BenDataLinkUDMDescription(bc, eecar_FPCSName_db) = "Company scheme" 'Formerly FPCS Now MARORS
    .BenDataLinkUDMDescription(bc, eecar_UseCompanyCarScheme_db) = "Use scheme?"
    .BenDataLinkUDMDescription(bc, eecar_AmountReceived_db) = "Amount received"
    .BenDataLinkUDMDescription(bc, eecar_LumpSum_db) = "Lump sum"
    .BenDataLinkUDMDescription(bc, eecar_IRFPCSValue) = "IR allowance" 'Formerly FPCS Now MARORS
    .BenDataLinkUDMDescription(bc, eecar_BenefitRepayableTaxable) = "Excess"
    .BenDataLinkUDMDescription(bc, ITEM_MADEGOOD) = "Gross made good"

  End With
    
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub
Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub

Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim rs As Recordset
  Dim eecar As IBenefitClass
  Dim s As String
  Dim i As Long
  Dim FPCS As FPCS 'Formerly FPCS Now MARORS
  Dim ben As IBenefitClass
  
  On Error GoTo EmployeeCar_readDB_Err
  Call xSet("EmployeeCar_readDB")
  
  If m_ReadFromDB Then GoTo EmployeeCar_readDB_End
  Set ben = Me
  
  Set rs = m_Parent.Parent.rsBenTables(TBL_EMPLOYEE_CARS)
  If Len(m_sbookmark) = 0 Then
    rs.FindFirst ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
    If Not rs.NoMatch Then
      m_sbookmark = rs.Bookmark
      i = i + 1
      rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set eecar = New EmployeeCar
        eecar.RSBookMark = rs.Bookmark
        Set eecar.Parent = m_Parent
        Call m_Parent.benefits.Add(eecar)
        Set eecar = Nothing
        rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    Call StandardReadData(ben, rs)
    'general
    
    'SCREEN
    ben.value(eecar_Item_db) = IIf(IsNull(rs.Fields("item")), "", rs.Fields("item"))
    ben.value(eecar_EngineSize_db) = IIf(IsNull(rs.Fields("enginesize")), 0, rs.Fields("enginesize"))
    ben.value(eecar_AmountReceived_db) = IIf(IsNull(rs.Fields("EECarAllowance")), 0, rs.Fields("EECarAllowance"))
    
    'EK No longer needed 2003 ben.value(eecar_AlternativeMethod) = rs.Fields("EECarAltMethod")
    ben.value(eecar_carkey_db) = rs.Fields("EECarKey")
    
'MP DB ToDo confirm removal - always setting to false both for read+write
'    ben.value(eecar_AverageIRRate_db) = False ' EK IR Rates no longer relevant rs.Fields("AveIRRate")
    ben.value(eecar_UseCompanyCarScheme_db) = rs.Fields("UseFPCS") 'Formerly FPCS Now MARORS
    
    s = IIf(IsNull(rs.Fields("FPCS")), "", rs.Fields("FPCS"))
    
   If Len(s) > 0 Then
      For i = 1 To p11d32.CurrentEmployer.FPCSchemes.Count 'Formerly FPCS Now MARORS
        Set FPCS = p11d32.CurrentEmployer.FPCSchemes(i)
        If Not FPCS Is Nothing Then
          If StrComp(FPCS.Name, s, vbTextCompare) = 0 Then
            ben.value(eecar_FPCS_db) = i
'MP DB ToDo confirm not used in any rpt as stated
'            ben.value(eecar_FPCSName_db) = s 'Display name for report wizard
            Exit For
          End If
        End If
      Next
      If i > p11d32.CurrentEmployer.FPCSchemes.Count Then 'Formerly FPCS Now MARORS
        'must be old default in p11d16 ie SCHEME1
        ben.value(eecar_FPCS_db) = 1  'ir FPCS so default to IR
'MP DB        ben.value(eecar_FPCSName_db) = S_IRFPCS
      End If
    Else
      ben.value(eecar_FPCS_db) = 1  'ir FPCS so default to IR
'MP DB      ben.value(eecar_FPCSName_db) = S_IRFPCS
 
    End If
     
    'dialog for extras
    ben.value(eecar_HireCostMadeGood_db) = IIf(IsNull(rs.Fields("HCmadegood")), 0, rs.Fields("HCmadegood"))
    ben.value(eecar_HireCost_db) = IIf(IsNull(rs.Fields("HireCost")), 0, rs.Fields("HireCost"))
    ben.value(eecar_LumpSum_db) = IIf(IsNull(rs.Fields("Lumpsum")), 0, rs.Fields("Lumpsum"))
    ben.value(eecar_CarMadeGood_db) = IIf(IsNull(rs.Fields("madegood")), 0, rs.Fields("madegood"))
    'sum of extras
    ben.value(eecar_TotalExtras) = GetTotalOfExtrasPostMadeGood

    ben.value(eecar_TotalMiles) = ReadEEMileage(mileage, ben.value(eecar_carkey_db))
    
    m_ReadFromDB = True
  End If
  
EmployeeCar_readDB_End:
  Set eecar = Nothing
  Set FPCS = Nothing 'Formerly FPCS Now MARORS
  Set rs = Nothing
  IBenefitClass_ReadDB = i
  Call xReturn("EmployeeCar_readDB")
  Exit Function
EmployeeCar_readDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeCar_readDB", "ERR_EECAR_READDB", "Unable to read the employee owned car details.")
  Resume EmployeeCar_readDB_End
  Resume
End Function
Public Function GetTotalOfExtrasPostMadeGood() As Long
  GetTotalOfExtrasPostMadeGood = GetTotalHireCost + CLng(m_BenItems(eecar_LumpSum_db)) + (CLng(m_BenItems(eecar_CarMadeGood_db) * -1))
End Function
Private Function GetTotalHireCost() As Long
  GetTotalHireCost = Max(0, m_BenItems(eecar_HireCost_db) - m_BenItems(eecar_HireCostMadeGood_db))
End Function

Private Sub WKEmployeeCarDetails(rep As Reporter, ben As IBenefitClass)
  On Error GoTo WKEmployeeCarDetails_ERR
  
  Call xSet("WKEmployeeCarDetails")
  
  
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Car: ", ben.value(ITEM_DESC))
  Call WKOut(rep, WK_ITEM_TEXT, "Engine size: " & S_WK_NORMAL_BOLD_FONT & ben.value(eecar_EngineSize_db)) 'EK Trying to sort it out
  'Call WKOut(rep, WK_ITEM_TEXT, "Car Engine Size: " & IIf(Not ben.value(eecar_EngineSize) Is Null, ben.value(eecar_EngineSize), "not entered"))
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Number of business miles travelled: ", ben.value(eecar_TotalMiles))
  
  If ben.value(eecar_WKAlternativeSection) Then
    Call WKOut(rep, WK_ITEM_TEXT, "Number of cars used by employee:" & ben.value(eecar_CarsWithAlternativeMethodFalse))
    Call WKOut(rep, WK_ITEM_TEXT, "Total Miles for all cars used:" & ben.value(eecar_TotalMilesCarsAlternativeFalse))
  End If
  
WKEmployeeCarDetails_END:
  Call xReturn("WKEmployeeCarDetails")
  Exit Sub
WKEmployeeCarDetails_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKEmployeeCarDetails", "WK Employee Car Details", "Error printing the employee car WK sheet car details.")
  Resume WKEmployeeCarDetails_END
  
End Sub
Private Sub WKEmployeeCarPayments(rep As Reporter, ben As IBenefitClass)
  On Error GoTo WKEmployeeCarPayments_ERR
  
  Call xSet("WKEmployeeCarPayments")
  
  
  Call WKOut(rep, WK_SECTION_HEADER, "Payments received by employee:")
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "Payments for mileage", ben.value(eecar_AmountReceived_db), , True)
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "Lump sum payments", ben.value(eecar_LumpSum_db))
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "Hire Costs", ben.value(eecar_HireCostNet))
  
  Call WKOut(rep, WK_ITEM_Total, S_WK_LEFT_MARGIN & "Total ", ben.value(eecar_WKValue), , True)
  
'  EK change to the calc of section E 2003
'  Call WKOut(rep, WK_SECTION_HEADER, "Less: FPCS Relief") 'Formerly FPCS Now MARORS
'
'  If ben.value(eecar_WKAlternativeSection) Then
'    Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & ben.value(eecar_TotalMilesCarsAlternativeFalse) & _
'               " miles at MARORS rates apportioned @ " & ben.value(eecar_TotalMiles) & "/" & ben.value(eecar_TotalMilesCarsAlternativeFalse), FormatWN(-1 * ben.value(eecar_IRFPCSValue), "")) 'Formerly FPCS Now MARORS
'  Else
'    Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & ben.value(eecar_TotalMiles) & " miles at MARORS rates", ben.value(eecar_IRFPCSValue))
'  End If
'
'  'value
'  Call WKOut(rep, WK_ITEM_Total, S_WK_LEFT_MARGIN & "Total", ben.value(ITEM_VALUE), , True)
'  'made good
  Call WKOut(rep, WK_SECTION_HEADER, "Less:")
  Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "Amounts made good (restricted to total payments made to employee)", FormatWN(-1 * ben.value(ITEM_MADEGOOD), ""))
  'benefit
  rep.Out (vbCrLf)
  Call WKOut(rep, WK_ITEM_Total, S_WK_LEFT_MARGIN & "Payments received less amounts made good for this car ", ben.value(eecar_WKValue) - 1 * ben.value(ITEM_MADEGOOD), , True)
  rep.Out (vbCrLf)
  'repayable
'  Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "Excess for this car:", ben.value(eecar_BenefitRepayableTaxable), , True, True)
'  Call rep.Out(vbCrLf)
      
WKEmployeeCarPayments_END:
  Call xReturn("WKEmployeeCarPayments")
  Exit Sub
WKEmployeeCarPayments_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKEmployeeCarPayments", "WK Employee Car Payments", "Error printing the employee car WK sheet car payments.")
  Resume WKEmployeeCarPayments_END
  Resume
End Sub

'Private Sub WKEmployeeCarFPCS(rep As Reporter, ben As IBenefitClass) 'Formerly FPCS Now MARORS
'  Dim er As Employer
'  Dim i As Long
'  Dim FPCSBanding As FPCSBanding
'  Dim eecar As EmployeeCar
'
'  On Error GoTo WKEmployeeCarFPCS_ERR
'
'  Call xSet("WKEmployeeCarFPCS")
'
'  Set eecar = ben
'
'  If ben.value(eecar_UseCompanyCarScheme) Then
'    Call WKOut(rep, WK_SECTION_HEADER, "Details of the employers mileage scheme(s) used." & vbCrLf)
'
'' EK IR Rates no longer relevant 2003
''    If ben.value(eecar_AverageIRRate) And ben.value(eecar_FPCS) = L_IRFPCS Then
''      Call WKOut(rep, WK_ITEM_TEXT, "Average Inland Revenue rates used (see Inland Revenue booklet 480)." & vbCrLf)
''    Else
'      Call WKTblColXOffsets(10, 20)
'      Call WKTblColFormats("rn", "rn")
'      Call WKTableHeadings(rep, "Miles", "@ Rate (£)")
'      Set er = GetParentFromBenefit(Me, GPBF_EMPLOYER)
'
'      For i = 1 To eecar.FPCSBandings.Count  'Formerly FPCS Now MARORS
'        Set FPCSBanding = eecar.FPCSBandings(i)
'        If Not FPCSBanding Is Nothing Then
'          Call WKTableRow(rep, FPCSBanding.Miles, FPCSBanding.Rate)
'        End If
'      Next
' '   End If
'  End If
'
'
'WKEmployeeCarFPCS_END:
'  Call xReturn("WKEmployeeCarFPCS")
'  Exit Sub
'WKEmployeeCarFPCS_ERR:
'  Call ErrorMessage(ERR_ERROR, Err, "WKEmployeeCarFPCS", "WK Employee Car FPCS", "Error printing the employee car WK sheet car FPCS details.")
'  Resume WKEmployeeCarFPCS_END
'End Sub

Private Sub WKEmployeeCarMARORS(rep As Reporter, lTotalBenefit As Variant, lTotalMiles As Variant, lTotalPayments As Variant)
  Dim csb As FPCSBand, CSB2 As FPCSBand
  Dim er As Employer
  Dim lLowbandMiles, lHighbandMiles As Long
  Dim CS As FPCS
  Dim lMileageLimit As Long
  Dim lowmilesrate As Double, highmilesrate As Double
  
'  Dim eecar As EmployeeCar
  
  On Error GoTo WKEmployeeCarMARORS_ERR
  
  Call xSet("WKEmployeeCarMARORS")
  
  Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)

  Call WKOut(rep, WK_ITEM_TEXT, S_WK_LEFT_MARGIN & "Total payments received less made good for all cars", lTotalPayments, , True)
    
  Set CS = p11d32.CurrentEmployer.FPCSchemes(L_IRFPCS)
 
      Set csb = CS.Bands(1)
      Set CSB2 = CS.Bands(2)
      lowmilesrate = csb.Rate
      highmilesrate = CSB2.Rate
 
  lMileageLimit = 10000
  If lTotalMiles >= lMileageLimit Then
    lLowbandMiles = lMileageLimit
    lHighbandMiles = lTotalMiles - lMileageLimit
  Else
    lLowbandMiles = lTotalMiles
    lHighbandMiles = 0
  End If
  
 ' FormatWN(lowrate, "£", , True)
  Call WKOut(rep, WK_SECTION_HEADER_LESS)
  
  Call WKOut(rep, WK_ITEM_TEXT, "Approved mileage allowance payments based on " & lTotalMiles & " total miles")
  
  Call WKOut(rep, WK_ITEM_TEXT, "       " & lLowbandMiles & " miles @ " & FormatWN(lowmilesrate, "£", , True), lowmilesrate * lLowbandMiles, , , True)
  If lHighbandMiles Then
    Call WKOut(rep, WK_ITEM_TEXT, "       " & lHighbandMiles & " miles @ " & FormatWN(highmilesrate, "£", , True), highmilesrate * lHighbandMiles, , , True)
  End If
  
  Call WKOut(rep, WK_ITEM_Total, S_WK_NORMAL_BOLD_FONT & "Assessable benefit / (Shortfall) ", IIf(lTotalBenefit > 0, lTotalBenefit, -lTotalBenefit), , True, IIf(lTotalBenefit > 0, False, True))
  
  Call WKOut(rep, WK_BLANK_LINE)
  Call WKOut(rep, WK_ITEM_TEXT, "Where the mileage allowance payments result in a shortfall to the HMRC's Approved Mileage Allowance Payments (the figure above is negative), the figure ")
  Call WKOut(rep, WK_ITEM_TEXT, "should be entered in box 17 of the employment pages of the self assessment tax return to be claimed as a Mileage Allowance Relief.")
  
  
WKEmployeeCarMARORS_END:
  Call xReturn("WKEmployeeCarFPCS")
  Exit Sub
WKEmployeeCarMARORS_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WKEmployeeCarFPCS", "WK Employee Car FPCS", "Error printing the employee car WK sheet car FPCS details.")
  Resume Next
End Sub


Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me, False)
End Function

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
      
  
  ben.value(eecar_CarsWithAlternativeMethodFalse) = 0
  ben.value(eecar_WKAlternativeSection) = False
  ben.value(eecar_TotalMilesCarsAlternativeFalse) = 0
  ben.value(eecar_CompanyFPCSValue) = 0  'Formerly FPCS Now MARORS
  ben.value(eecar_IRFPCSValue) = 0
  ben.value(eecar_Value) = S_ERROR
  ben.value(eecar_Benefit) = S_ERROR
  ben.value(eecar_Benefit_Reportable) = False
  ben.value(eecar_BenefitRepayableTaxable_Reportable) = False
  ben.value(eecar_MMSectionU_Reportable) = False
  ben.value(eecar_ERROR) = ""
End Sub

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_EMPLOYEE_CARS
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_EMPLOYEE_CAR_E, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
    IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim FPCS As FPCS  'Formerly FPCS Now MARORS
  Dim ben As IBenefitClass
  
  On Error GoTo EmployeeCar_WriteDB_Err
  Call xSet("EmployeeCar_WriteDB")
  
  Set ben = Me
  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields("item") = ben.value(eecar_Item_db)
  rs.Fields("enginesize") = ben.value(eecar_EngineSize_db)
  'rs.Fields("enginesize") = IIf(Not ben.value(eecar_EngineSize) Is Null, ben.value(eecar_EngineSize), 0) 'EK Trying to sort it out

'MP DB ToDo confirm
'  rs.Fields("AveIRRate") = False 'EK no longer relevant 2003 ben.value(eecar_AverageIRRate)
  rs.Fields("UseFPCS") = ben.value(eecar_UseCompanyCarScheme_db)  'Formerly FPCS Now MARORS
  
  If ben.value(eecar_FPCS_db) = 1 Then
    rs.Fields("FPCS") = S_SCHEME1
  Else
    If p11d32.CurrentEmployer.FPCSchemes.Count >= ben.value(eecar_FPCS_db) Then
      Set FPCS = p11d32.CurrentEmployer.FPCSchemes(ben.value(eecar_FPCS_db))
      rs.Fields("FPCS") = FPCS.Name
    Else
      rs.Fields("FPCS") = S_SCHEME1
    End If
  End If
  
  Call StandardWriteData(ben, rs, False)
  If Not p11d32.BringForward.Yes Then
    
    rs.Fields("EECarAllowance") = ben.value(eecar_AmountReceived_db)
    'EK No longer needed 2003 rs.Fields("EECarAltMethod") = ben.value(eecar_AlternativeMethod)
    'dialog for extras
    rs.Fields("HCmadegood") = ben.value(eecar_HireCostMadeGood_db)
    rs.Fields("HireCost") = ben.value(eecar_HireCost_db)
    rs.Fields("Lumpsum") = ben.value(eecar_LumpSum_db)
    rs.Fields("madegood") = ben.value(eecar_CarMadeGood_db)
  End If
  
  Call BenefitCloseRecord(ben, rs)
  
  If IsNull(rs.Fields("EECarKey").value) Or p11d32.BringForward.Yes Then
    'if is null, or bringing data forward then benefit is new so add in a new GUID
    rs.Edit
    rs.Fields("EECarKey") = GenerateGUID
    rs.Update
  End If
  
  ben.value(eecar_carkey_db) = rs.Fields("EECarKey")
  GenerateGUID
  If Not p11d32.BringForward.Yes Then Call WriteEEMileage
  
  ben.Dirty = False
  IBenefitClass_WriteDB = True
  
EmployeeCar_WriteDB_End:
  Set FPCS = Nothing
  Set rs = Nothing
  Call xReturn("EmployeeCar_WriteDB")
  Exit Function
  
EmployeeCar_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeCar_WriteDB", "Employee Car Write DB", "Error writing an employee car to the database.")
  Resume EmployeeCar_WriteDB_End
  Resume
End Function
Private Function WriteEEMileage() As Long
'MP DB - changed to Private - not called from anywhere else
  Dim Miles As MileDetail
  Dim rs As Recordset
  Dim l As Long
  Dim ben As IBenefitClass, benEmployee As IBenefitClass
  
  
  On Error GoTo WriteEEMileage_Err
  Call xSet("WriteEEMileage")
  
  Set ben = Me
  Set benEmployee = m_Parent
  
  Call m_Parent.Parent.db.Execute(sql.Queries(DELETE_EECARMILES, ben.value(eecar_carkey_db)))
  
  Set rs = m_Parent.Parent.db.OpenRecordset(sql.Queries(SELECT_EECARMILES, ben.value(eecar_carkey_db)), dbOpenDynaset)
  
  For l = 1 To mileage.Count
    Set Miles = mileage.Item(l)
    If Not Miles Is Nothing Then
      rs.AddNew
      If Len(Miles.MileItem) > 0 Then
        rs.Fields("MilesItem") = Miles.MileItem
      Else
        rs.Fields("MilesItem") = Null
      End If
      If Miles.MileDate <> UNDATED Then
        rs.Fields("MilesDate") = Miles.MileDate
      Else
        rs.Fields("MilesDate") = Null
      End If
      rs.Fields("Miles") = Miles.MileAmount
      rs.Fields("EECarKey") = ben.value(eecar_carkey_db)
      rs.Fields("P_Num") = benEmployee.value(ee_PersonnelNumber_db)
'MP DB ToDo - "Reg" not read for display - so should it be Written to db? Can it be del from DB?
      rs.Fields(S_FIELD_CAR_REGISTRATION) = ben.value(eecar_Item_db) 'MP DB should this be EE surname or car description?
      rs.Update
      WriteEEMileage = WriteEEMileage + 1
    End If
  Next

WriteEEMileage_End:
  Set Miles = Nothing
  Set rs = Nothing
  Call xReturn("WriteEEMileage")
  Exit Function

WriteEEMileage_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteEEMileage", "Write Mileage", "Error writing the car mileage into the database.")
  Resume WriteEEMileage_End
  Resume
End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ibc As IBenefitClass
  
  On Error GoTo EmployeeCar_DeleteDB_Err
  
  Call xSet("EmployeeCar_DeleteDB")
  
  Set ibc = Me
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(ibc.TABLE)
  
  If ibc.HasBookMark Then
    rs.Bookmark = m_sbookmark
    rs.Delete
    'delete associated miles
    Call p11d32.CurrentEmployer.db.Execute(sql.Queries(DELETE_EECARMILES, m_BenItems(eecar_carkey_db)))
  End If
  
  IBenefitClass_DeleteDB = True
  
EmployeeCar_DeleteDB_End:
  Set ibc = Nothing
  Set rs = Nothing
  Call xReturn("EmployeeCar_DeleteDB")
  Exit Function
  
EmployeeCar_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeCar_DeleteDB", "EmployeeCar Delete DB", "Error in deleting an employee owned car from the database.")
  Resume EmployeeCar_DeleteDB_End
  Resume
End Function


Private Function FPCSExists(ByVal lFPCSIndex As Long) As Boolean  'Formerly FPCS Now MARORS
  Dim FPCS As FPCS
  On Error GoTo FPCSExits_END
  
  Set FPCS = p11d32.CurrentEmployer.FPCSchemes(lFPCSIndex)
  If Not FPCS Is Nothing Then FPCSExists = True
  
  
FPCSExits_END:
  Exit Function
End Function
Public Function GetFPCSMileageValue(value As Variant, ByVal Miles As Long, ByVal CC As Long, ByVal FPCSIndex As Long, ByVal UseAverageIRRates As Boolean) As FPCS
  Dim csb As FPCSBand, CSB2 As FPCSBand
  Dim i As Long, lTempMiles As Long
  Dim CS As FPCS, CCBands As ObjectList  'Formerly FPCS Now MARORS
  Dim dblRate0to4000 As Double, dblRate4000andAbove As Double 'only used for AverageIR
  Dim FPCSBanding As FPCSBanding
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo GetFPCSMileageValue_Err
  Call xSet("GetFPCSMileageValue")

  'if the scheme has been removed then need to calc with IR rates and set FPCS = L_IRFPCS
  
  Set CS = p11d32.CurrentEmployer.FPCSchemes(FPCSIndex)
  Call FPCSBandings.RemoveAll
  
' EK Avg IR Rates now never relevant
'  If UseAverageIRRates And (FPCSIndex = L_IRFPCS) Then
'    'we use IR scheme regardless of FPCSIndex
'    'we want the 2 middle bands for cc purposes and and average of the 2 rates at each cc mileage level
'    'there are only two levels 0 to 4000 and > 4000 miles
'    If CS Is Nothing Then Call Err.Raise(ERR_NOCARSCHEME, "GetFPCSMileageValue", "The car scheme with index " & FPCSIndex & " does not exist.")
'    '4 cc bands with 2 mileages
'    If CS.Bands.Count <> 8 Then Call Err.Raise(ERR_IRFPCS, "GetFPCSMileageValue", "The IR FPCS does not have 8 bands it has " & CS.Bands.Count & ".")
'    'see ReadIRFPCS in TaxRates.cls for order of reading in IR fixedprofit car schemes
'    Set csb = CS.Bands(3)
'    Set CSB2 = CS.Bands(5)
'    dblRate0to4000 = (csb.Rate + CSB2.Rate) / 2
'    Set csb = CS.Bands(4)
'    Set CSB2 = CS.Bands(6)
'    dblRate4000andAbove = (csb.Rate + CSB2.Rate) / 2
'    lTempMiles = Miles - p11d32.Rates.value(fpcsMILEAGELIMIT)
'    If lTempMiles > 0 Then
'      value = lTempMiles * dblRate4000andAbove
'      value = value + (p11d32.Rates.value(fpcsMILEAGELIMIT) * dblRate0to4000)
'    Else
'      value = Miles * dblRate0to4000
'    End If
'    Set m_TopBand = Nothing
'
'  Else
    Call GetAllBandsInSameCCInOrder(CCBands, CS.Bands, CC)
    value = 0#: Set m_TopBand = Nothing
    For i = CCBands.Count To 1 Step -1
      Set csb = CCBands(i)
      lTempMiles = Miles - csb.GreaterThanMiles
      If lTempMiles > 0 Then
        Set FPCSBanding = New FPCSBanding
        FPCSBanding.Miles = lTempMiles
        FPCSBanding.Rate = csb.Rate
        Call FPCSBandings.Add(FPCSBanding)
        
        If m_TopBand Is Nothing Then
          Set m_TopBand = csb
        End If
        value = value + (lTempMiles * csb.Rate)
        Miles = Miles - lTempMiles
      End If
    Next
    
'  End If
  
  
  
GetFPCSMileageValue_End:
  Set GetFPCSMileageValue = CS
  Set csb = Nothing
  Call xReturn("GetFPCSMileageValue")
  Exit Function

GetFPCSMileageValue_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "GetFPCSMileageValue", "Get FPCS Mileage Value", "Error getting the mileage value according to the selected FPCS.")
  Resume GetFPCSMileageValue_End
  Resume
End Function


Private Function GetIRFPCSValue(vMileageValue As Variant) As Boolean  'Formerly FPCS Now MARORS
  Dim ben As IBenefitClass
  
  On Error GoTo GetIRFPCSValue_Err
  Call xSet("GetIRFPCSValue")
  Set ben = Me
  vMileageValue = 0
  'EK No longer needed 2003 If ben.value(eecar_AlternativeMethod) Then
  '  Call GetFPCSMileageValue(vMileageValue, ben.value(eecar_TotalMiles), ben.value(eecar_EngineSize), L_IRFPCS, False) ' EK IR Rates no longer relevant ben.value(eecar_AverageIRRate)) 'EK Trying to sort EE Car calc out
  'Else
    'get the miles for all the cars where m_EECarItem(eecar_AlternativeMethod) = 0 and add together
    Call GetFPCSMileageValue(vMileageValue, GetNonAlternativeCarMiles, ben.value(eecar_EngineSize_db), L_IRFPCS, False) ' EK IR Rates no longer relevant ben.value(eecar_AverageIRRate))
    vMileageValue = vMileageValue * (ben.value(eecar_TotalMiles) / IIf(ben.value(eecar_TotalMilesCarsAlternativeFalse) = 0, 1, ben.value(eecar_TotalMilesCarsAlternativeFalse)))
  'End If

GetIRFPCSValue_End:
  Call xReturn("GetIRFPCSValue")
  Exit Function

GetIRFPCSValue_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "GetIRFPCSValue"), Err.Description)
  Resume GetIRFPCSValue_End
  Resume
End Function
Private Function MMRevenueReliefRate() As String
  Dim lCC As Long
  
  lCC = m_BenItems(eecar_EngineSize_db)
  
  'EK No longer needed 2003 If m_BenItems(eecar_AlternativeMethod) = True Then
  '  MMRevenueReliefRate = "A"
  'Else
    'EK Now only one revenue rate
'    If lCC < 1001 Then
      MMRevenueReliefRate = "1"
'    ElseIf lCC >= 1001 And lCC < 1501 Then
'      MMRevenueReliefRate = "2"
'    ElseIf lCC >= 1501 And lCC < 2001 Then
'      MMRevenueReliefRate = "3"
'    Else
'      MMRevenueReliefRate = "4"
'    End If
  'End If
End Function

Public Function GetNonAlternativeCarMiles() As Long
  Dim benefit As IBenefitClass
  Dim l As Long
  Dim ben As IBenefitClass
  
  On Error GoTo GetNonAlternativeCarMiles_Err
  Call xSet("GetNonAlternativeCarMiles")

  Set ben = Me

  'have to go through all current as can not reference forms list view for a shortcut
  With m_Parent
    For l = 1 To .benefits.Count
      Set benefit = .benefits(l)
      If Not benefit Is Nothing Then
        If (benefit.BenefitClass = IBenefitClass_BenefitClass) Then
          'EK No longer needed 2003 If benefit.value(eecar_AlternativeMethod) = False Then
             ben.value(eecar_TotalMilesCarsAlternativeFalse) = ben.value(eecar_TotalMilesCarsAlternativeFalse) + benefit.value(eecar_TotalMiles)
             ben.value(eecar_CarsWithAlternativeMethodFalse) = ben.value(eecar_CarsWithAlternativeMethodFalse) + 1
          ' End If
        End If
      End If
    Next
  End With
  
  GetNonAlternativeCarMiles = ben.value(eecar_TotalMilesCarsAlternativeFalse)
  
GetNonAlternativeCarMiles_End:
  Call xReturn("GetNonAlternativeCarMiles")
  Exit Function

GetNonAlternativeCarMiles_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetNonAlternativeCarMiles", "Get Non Alternative Car Miles", "Error getting the non alternative car miles total.")
  Resume GetNonAlternativeCarMiles_End
End Function
Private Function CalcEmployeeCarsWithAlternativMethodFalse() As Long
  Dim benefit As IBenefitClass, benefitX  As IBenefitClass
  Dim l As Long, m As Long
  Dim ibf As IBenefitForm2
  Dim vBenefit As Variant
  On Error GoTo CalcEmployeeCarsWithAlternativMethodFalse_Err
  Call xSet("CalcEmployeeCarsWithAlternativMethodFalse")

  Set benefitX = Me
  
  If m_Parent.CalculatingDependantEmployeeCars Then GoTo CalcEmployeeCarsWithAlternativMethodFalse_End:
  m_Parent.CalculatingDependantEmployeeCars = True
  
  
  benefitX.value(eecar_NoOfCars) = 1
  
  For l = 1 To m_Parent.benefits.Count
    Set benefit = m_Parent.benefits(l)
    If Not (benefit Is Nothing) And Not (benefit Is benefitX) Then
      If benefit.BenefitClass = benefitX.BenefitClass Then
        benefitX.value(eecar_NoOfCars) = benefitX.value(eecar_NoOfCars) + 1
        vBenefit = S_ERROR
        benefit.NeedToCalculate = True
        vBenefit = benefit.Calculate
        If CurrentForm Is F_EmployeeCar Then
          Set ibf = CurrentForm
          For m = 1 To ibf.lv.listitems.Count
            If l = ibf.lv.listitems(m).Tag Then
              ibf.lv.listitems(m).Text = benefit.Name
              ibf.lv.listitems(m).SubItems(1) = FormatWN(vBenefit, "£")
            End If
          Next
        End If
      End If
    End If
  Next l
  
  m_Parent.CalculatingDependantEmployeeCars = False
  
CalcEmployeeCarsWithAlternativMethodFalse_End:
  Call xReturn("CalcEmployeeCarsWithAlternativMethodFalse")
  Exit Function

CalcEmployeeCarsWithAlternativMethodFalse_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CalcEmployeeCarsWithAlternativMethodFalse", "Calc Employee Cars With Alternative Method False", "Error calculating the benefit for employee owned cars with the alternative method set to false.")
  Resume CalcEmployeeCarsWithAlternativMethodFalse_End
  Resume
End Function
Public Function SortBands(DstBands As ObjectList, SrcBands As ObjectList, ByVal sb As FPCS_SORT) As Long
  Dim lMinValue As Long, iMin As Long, bFirst As Boolean
  Dim i As Long, j As Long, lLastMinValue As Long
  Dim obtmp As ObjectList
  Dim Band As FPCSBand, NewSrcBands As Boolean  'Formerly FPCS Now MARORS
  
  On Error GoTo SortBands_Err
  Call xSet("SortBands")
  
  If SrcBands Is Nothing Then GoTo SortBands_End:
  
  Set DstBands = New ObjectList
  Set obtmp = New ObjectList
  
  For i = 1 To SrcBands.Count
    If Not SrcBands(i) Is Nothing Then
      Call obtmp.Add(SrcBands(i))
    End If
  Next i
  
  'get the first one
  For i = 1 To obtmp.Count
    bFirst = True
    For j = 1 To obtmp.Count
      If Not obtmp(j) Is Nothing Then
        Set Band = obtmp(j)
        Select Case sb
          Case FPCS_CC_ASC
            If Band.GreaterThanCC < lMinValue Or bFirst Then
              lMinValue = Band.GreaterThanCC
              iMin = j
              bFirst = False
            End If
          Case FPCS_MILES_ASC
            If Band.GreaterThanMiles < lMinValue Or bFirst Then
              lMinValue = Band.GreaterThanMiles
              iMin = j
              bFirst = False
            End If
        End Select
      End If
    Next j
    ' remove and add
    Call DstBands.Add(obtmp(iMin))
    Call obtmp.Remove(iMin)
  Next i

  SortBands = DstBands.Count
  
SortBands_End:
  Set Band = Nothing
  Call xReturn("SortBands")
  Exit Function
SortBands_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SortBands", "Sort Band", "Error sorting the FPCS mileage bands.")
  Resume SortBands_End
  Resume
End Function
Private Sub PrintBands(OL As ObjectList, sCaption As String)
  Dim csb As FPCSBand  'Formerly FPCS Now MARORS
  Dim i As Long
  
  
 'Debug.PrintsCaption
  For i = 1 To OL.Count
    Set csb = OL(i)
    If Not csb Is Nothing Then
     'Debug.Print" Name: " & csb.Name & " , CC:" & csb.GreaterThanCC & " , Miles:" & csb.GreaterThanMiles
    End If
  Next
End Sub
Public Function GetAllBandsInSameCCInOrder(CCBands As ObjectList, SrcBands As ObjectList, lCC As Long) As Boolean
  Dim i As Long
  Dim csb As FPCSBand 'Formerly FPCS Now MARORS
  Dim TempBands As ObjectList, TempBands2 As ObjectList
  
  On Error GoTo GetAllBandsInSameCCInOrder_Err
  Call xSet("GetAllBandsInSameCCInOrder")
  
  Set CCBands = New ObjectList
  
  Call SortBands(TempBands, SrcBands, FPCS_CC_ASC)
  
  For i = TempBands.Count To 1 Step -1
    Set csb = TempBands(i)
    If lCC >= csb.GreaterThanCC Then ' EK Trying to sort it out, had commented this line out
      Call GetAllBandsInSameCC(TempBands2, TempBands, csb.GreaterThanCC)
      Exit For
    End If
  Next
  
  GetAllBandsInSameCCInOrder = SortBands(CCBands, TempBands2, FPCS_MILES_ASC)
  
  
GetAllBandsInSameCCInOrder_End:
  Set TempBands = Nothing
  Set TempBands2 = Nothing
  Set csb = Nothing
  Call xReturn("GetAllBandsInSameCCInOrder")
  Exit Function

GetAllBandsInSameCCInOrder_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetAllBandsInSameCCInOrder", "Get All Bands In Same CC In Order", "Error getting mileage bands with the same engine size in mileage order.")
  Resume GetAllBandsInSameCCInOrder_End
End Function

Public Function GetAllBandsInSameCC(CCBands As ObjectList, SrcBands As ObjectList, lGreaterThanCC As Long) As Boolean
  Dim i As Long
  Dim Band As FPCSBand 'Formerly FPCS Now MARORS
  
  On Error GoTo GetAllBandsInSameCC_Err
  Call xSet("GetAllBandsInSameCC")
  
  Set CCBands = Nothing
  
  Set CCBands = New ObjectList
  
  For i = 1 To SrcBands.Count
    Set Band = SrcBands(i)
    If Not Band Is Nothing Then
      If Band.GreaterThanCC = lGreaterThanCC Then
        Call CCBands.Add(Band)
      End If
    End If
  Next

  GetAllBandsInSameCC = CCBands.Count

GetAllBandsInSameCC_End:
  Set Band = Nothing
  Call xReturn("GetAllBandsInSameCC")
  Exit Function

GetAllBandsInSameCC_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetAllBandsInSameCC", "Get All Bands In Same CC", "Error getting an object list of all bands with the same engine size.")
  Resume GetAllBandsInSameCC_End
  Resume
End Function

Private Function ReadEEMileage(mileage As ObjectList, ByVal lCarKey As Variant) As Long
'RC - lCarKey is defined as a variant as when bringforward.yes then this value is numeric
'     if just current year file then it is a GUID (string)
  Dim Miles As MileDetail
  Dim rs As Recordset
  Dim lMiles As Long
  
  On Error GoTo ReadEEMileage_Err
  Call xSet("ReadEEMileage")

  Set rs = m_Parent.Parent.db.OpenRecordset(sql.Queries(SELECT_EECARMILES, lCarKey), dbOpenForwardOnly)
  
  Do While Not rs.EOF
    Set Miles = New MileDetail
    Miles.MileItem = "" & rs.Fields("MilesItem").value
    If IsNull(rs.Fields("MilesDate").value) Then
      Miles.MileDate = UNDATED
    Else
      Miles.MileDate = rs.Fields("MilesDate").value
    End If
    Miles.MileAmount = IIf(IsNull(rs.Fields("Miles").value), 0, rs.Fields("Miles").value)
    lMiles = lMiles + Miles.MileAmount
    Call mileage.Add(Miles)
    rs.MoveNext
  Loop
  
ReadEEMileage_End:
  Set Miles = Nothing
  ReadEEMileage = lMiles
  Call xReturn("ReadEEMileage")
  Exit Function
ReadEEMileage_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadEEMileage", "Read Mileage", "Error reading the car mileage.")
  ReadEEMileage = 0
  Resume ReadEEMileage_End
End Function

