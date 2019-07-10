VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitClass

Public CompanyCar As IBenefitClass
Private m_CompanyDefined As Boolean
Private m_NeedToCalculate As Boolean
Public Enum FuelItems
'
'  fuel_Registration
'  fuel__SPECIFIC_START = fuel_Registration
'  fuel_Value
'  fuel_MadeGood_NET
'  fuel_Benefit
'  fuel_MadeGood
'  fuel_ActualAmountMadeGood
'  fuel_Benefit_Reportable
'  fuel_UDM_BENEFIT_TITLE
'  fuel_BoxNumber
'  fuel_MadeGoodIsTaxDeducted
'  fuel_Class1AAdjustment
'  fuel_NICClass1AAble
'  fuel_NIC_Class1A_Value
'  fuel_NIC_Class1A_Benefit
'
'  'Fuel Benefit apportionment
    fuel_availablefrom                        'MP DB (not used) - remove
    fuel_AvailableTo                          'MP DB (not used) - remove
'  fuel__SPECIFIC_END = fuel_AvailableTo
'  fuel_UnavailableDays
'  fuel_VATBenefit
'  fuel_VAT
'  fuel_GrossFuel
'  fuel_daysreduction
'  fuel_p46FuelType
'  fuel_p46FuelTypeString
   fuel_p46FuelTypeStringLong                   'MP DB (not used) - remove
'  'Lastitem
   fuel_lastitem = fuel_p46FuelTypeStringLong   'MP DB (not used) - remove
  
End Enum
Private m_BenItems(1 To fuel_lastitem)


Private Function IBenefitClass_CalculateBody() As Variant

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "CompanyCar"
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

End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function

Private Function AdjustItemEnum(ByVal Item As RelocItems) As RelocItems
  On Error GoTo AdjustItemEnum_ERR
  
  Call xSet("AdjustItemEnum")
  
  If Item < car_FUEL_SPECIFIC_START Then
    AdjustItemEnum = Item + (car_FUEL_SPECIFIC_START - 1)
  ElseIf Item > car_FUEL_SPECIFIC_END Then
    AdjustItemEnum = Item
  Else
    Call Err.Raise(ERR_ENUM__VALUE_INVALID, "AdjustItemEnum", "Enum value is not valid for Fuel benefit.")
  End If
  
AdjustItemEnum_END:
  Call xReturn("AdjustItemEnum")
  Exit Function
AdjustItemEnum_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AdjustItemEnum", "Adjust Item Enum", "The Item number passed is outside the enum range of CompanyCarItems.")
  Resume AdjustItemEnum_END
  Resume
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
  
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_FUEL_F
End Property

Private Function IBenefitClass_Calculate() As Variant
  Call CompanyCar.Calculate
'MP DB ToDo ITEM_BENEFIT enum element does not exist - remove below line?
  IBenefitClass_Calculate = IBenefitClass_value(ITEM_BENEFIT)
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  m_CompanyDefined = NewValue
  
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = m_CompanyDefined
End Property
Private Function IBenefitClass_DeleteDB() As Boolean
  'do not map to parent
End Function
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  CompanyCar.Dirty = NewValue
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = CompanyCar.Dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Sub IBenefitClass_SetCalcDefaults()
'MP DB ToDo - find if this gets called b4 CompanyCar in any circumstances - if not then no need to call again
  Call CompanyCar.SetCalcDefaults
End Sub

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = CompanyCar.value(AdjustItemEnum(Item))
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  CompanyCar.value(AdjustItemEnum(Item)) = RHS
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = CompanyCar.HasBookMark
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  CompanyCar.InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = CompanyCar.InvalidFields
End Property
Private Sub IBenefitClass_Kill()
  Set CompanyCar = Nothing
End Sub
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = CompanyCar.Name
End Property
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set CompanyCar.Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = CompanyCar.Parent
End Property
Private Property Get IBenefitClass_PrintHeader() As String

End Property
Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean

End Function
Private Function IBenefitClass_ReadDB() As Long
  IBenefitClass_ReadDB = CompanyCar.ReadDB
End Function
Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  'do not map to parent
End Property
Private Property Get IBenefitClass_ReadFromDB() As Boolean
'MP DB ToDo - needs to be called?
  IBenefitClass_ReadFromDB = CompanyCar.ReadFromDB
End Property
Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  'do not map to parent
End Property
Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = CompanyCar.RSBookMark
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
  bc = BC_FUEL_F
  m_NeedToCalculate = True
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
     
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  
  'CAD? Car and fuel
  With p11d32
    Set ben = Me
    .BenDataLinkBenfitTable(bc) = ben.TABLE
  
    .BenDataLinkUDMDescription(bc, ITEM_DESC) = S_UDM_DESCRIPTION
    .BenDataLinkUDMDescription(bc, ITEM_VALUE) = S_UDM_VALUE 'RK 19/07/03 Suppress for car fuel
    .BenDataLinkUDMDescription(bc, ITEM_MADEGOOD_NET) = S_UDM_MADE_GOOD_NET 'RK 19/07/03
    

    .BenDataLinkUDMDescription(bc, ITEM_BENEFIT) = S_UDM_BENEFIT
    .BenDataLinkUDMDescription(bc, ITEM_UDM_BENEFIT_TITLE) = S_UDM_BENEFIT_TITLE
    .BenDataLinkUDMDescription(bc, ITEM_BOX_NUMBER) = S_UDM_BOX_NUMBER
    .BenDataLinkUDMDescription(bc, ITEM_NIC_CLASS1A_BENEFIT) = S_UDM_NIC_CLASS1A_BENEFIT
    'Call SetStandardBenItemsUDMData(BC_FUEL_F)
    Call SetStandardBenItemsDataTypes(BC_FUEL_F)
    'no class1aNIC as in car
  End With
  
  Call AbacusUDMData(bc)
  
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information.")
  Resume SetBenItemsInformation_end
  Resume
End Sub
Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_COMPANY_CARS
End Property
Private Function IBenefitClass_WriteDB() As Boolean
'MP DB ToDo - keep?
  IBenefitClass_WriteDB = True
End Function


