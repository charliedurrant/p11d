VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NonSharedVan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Base 0

Implements IBenefitClass

Public Enum NonSharedVanItems
  nsvan_item_db = ITEM_DESC 'RW
  nsvan_Value = ITEM_VALUE
  nsvan_MadeGood_Net = ITEM_MADEGOOD_NET
  nsvan_Benefit = ITEM_BENEFIT 'W
  nsvan_madegood_db = ITEM_MADEGOOD
  nsvan_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  nsvan_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  nsvan_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  nsvan_BoxNumber = ITEM_BOX_NUMBER
  nsvan_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  nsvan_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  nsvan_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  nsvan_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  nsvan_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  nsvan_Error = ITEM_ERROR
  nsvan_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  nsvan_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  nsvan_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE

  nsvan_AvailableFrom_db 'RW
  nsvan_Availableto_db 'RW
  nsvan_DaysUnavailable_db 'RW
  nsvan_commuter_use_req_db
  nsvan_availablle_for_bus_use_only
  nsvan_fuel_available_db
  nsvan_fuel_value
  nsvan_fuel_benefit
  nsvan_benefit_van_only 'only the van benefit
  nsvan_is_electric_db 'CAD 2010
  nsvan_RegistrationDate_db  'km
'MP DB (not used)  nsvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear 'RW 'bf

'MP DB  nsvan_LASTITEM = nsvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear
  nsvan_LASTITEM = nsvan_RegistrationDate_db
End Enum

Private m_CompanyDefined As Boolean
Private m_InvalidFields As Long
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_Parent As IBenefitClass
Private m_BenItems(1 To nsvan_LASTITEM)
Private m_dirty As Boolean

Private m_BenClass As BEN_CLASS
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Non shared van - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim lDaysUnavailable As Long
  Dim dDateFactor As Double
  Dim ben As IBenefitClass
  
  On Error GoTo SharedVan_Calculate_err
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  
  dDateFactor = dGetDateFactor(lDaysUnavailable, ben.value(nsvan_AvailableFrom_db), ben.value(nsvan_Availableto_db), ben.value(nsvan_DaysUnavailable_db), 30, False) 'PS - have set the final argument, EK - have unset and changed variable name as did not have result required.
                                                                                                                                                            'to false as is leap year (same as Company Car)
  If ben.value(nsvan_RegistrationDate_db) >= p11d32.Rates.value(carFOURYEAROLD) Then
    ben.value(nsvan_Value) = p11d32.Rates.value(vanHighCharge) * dDateFactor
  Else
    ben.value(nsvan_Value) = p11d32.Rates.value(vanLowCharge) * dDateFactor
  End If
  If (ben.value(nsvan_is_electric_db)) Then
    ben.value(nsvan_Value) = ben.value(nsvan_Value) * p11d32.Rates.value(vanElectricMultiplier)
  End If

  ben.value(nsvan_Value_Non_OPRA) = ben.value(nsvan_Value)
  
  If (ben.value(nsvan_fuel_available_db)) And (Not ben.value(nsvan_is_electric_db)) Then
    ben.value(nsvan_fuel_value) = p11d32.Rates.value(vanRate) * dDateFactor 'cad 2010
  Else
    ben.value(nsvan_fuel_value) = 0
  End If
  
  ben.value(nsvan_benefit_van_only) = ben.value(nsvan_Value)

  ben.value(nsvan_Value) = ben.value(nsvan_benefit_van_only) + ben.value(nsvan_fuel_value)
  
  ben.value(nsvan_MadeGood_Net) = Min(ben.value(nsvan_Value), ben.value(nsvan_madegood_db))
  ben.value(nsvan_Benefit) = ben.value(nsvan_Value) - ben.value(nsvan_MadeGood_Net)

  ben.value(nsvan_fuel_benefit) = ben.value(nsvan_fuel_value)
  
  If ben.value(nsvan_MadeGood_Net) > 0 Then
    ben.value(nsvan_benefit_van_only) = ben.value(nsvan_benefit_van_only) - ben.value(nsvan_MadeGood_Net)
    If (ben.value(nsvan_benefit_van_only) < 0) Then
      ben.value(nsvan_fuel_benefit) = ben.value(nsvan_fuel_benefit) + ben.value(nsvan_benefit_van_only)
    End If
    ben.value(nsvan_fuel_benefit) = Max(0, ben.value(nsvan_fuel_benefit))
    ben.value(nsvan_benefit_van_only) = Max(0, ben.value(nsvan_benefit_van_only))
  End If
  
  If ben.value(nsvan_availablle_for_bus_use_only) = True And ben.value(nsvan_commuter_use_req_db) Then
    ben.value(nsvan_Benefit) = 0
    ben.value(nsvan_benefit_van_only) = 0
    ben.value(nsvan_fuel_benefit) = 0
    ben.value(nsvan_fuel_value) = 0
    ben.value(nsvan_Value_Non_OPRA) = 0
  End If
  
  If ben.value(nsvan_Benefit) > 0 Then ben.value(nsvan_Benefit_Reportable) = True
  
  Call BenCalcNIC(ben)
  IBenefitClass_CalculateBody = ben.value(nsvan_Benefit)
  
SharedVan_Calculate_end:
  Exit Function
SharedVan_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume SharedVan_Calculate_end
  Resume

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
  IBenefitClass_CanBringForward = StandardCanBringForward(Me, nsvan_Availableto_db)
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_nonSHAREDVAN_G
End Property


Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  m_CompanyDefined = NewValue
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = m_CompanyDefined
End Property


Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim ibc As IBenefitClass
  
  On Error GoTo NonSharedVan_DeleteDB_Err
  
  Call xSet("NonSharedVan_DeleteDB")
  
  Set ibc = Me
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(ibc.TABLE)
  
  If ibc.HasBookMark Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  
  IBenefitClass_DeleteDB = True
  
NonSharedVan_DeleteDB_End:
  Set ibc = Nothing
  Set rs = Nothing
  Call xReturn("NonSharedVan_DeleteDB")
  Exit Function
NonSharedVan_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVan_DeleteDB", "Non Shared Van DeleteDB", "Error deleting a non shared van benefit from the database.")
  Resume NonSharedVan_DeleteDB_End
  
End Function

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
  m_Parent.Dirty = m_Parent.Dirty Or m_dirty
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Call ECASE("No PrintWkBody for Non shared van, in non shared vans.")
End Function

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben  As IBenefitClass
  
  Set ben = Me
  ben.value(nsvan_Value) = S_ERROR
  ben.value(nsvan_Benefit) = S_ERROR
  ben.value(nsvan_fuel_benefit) = S_ERROR
  ben.value(nsvan_fuel_value) = S_ERROR
  
  ben.value(nsvan_MadeGood_Net) = S_ERROR
  ben.value(nsvan_Benefit_Reportable) = False
  ben.value(nsvan_Error) = ""
End Sub
Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(nsvan_item_db)
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_PrintHeader() As String

End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  IBenefitClass_PrintWk = m_Parent.PrintWk(rep)
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim rs As Recordset
  Dim ibc As IBenefitClass
  
  On Error GoTo clsLoan_ReadDB_Err
  
  xSet "clsLoan_ReadDB"
  
  
  Set ibc = Me
  If ibc.ReadFromDB Then GoTo clsLoan_ReadDB_end
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(ibc.TABLE)
  
  If Len(ibc.RSBookMark) > 0 Then
    rs.Bookmark = ibc.RSBookMark
    With rs
      Call StandardReadData(ibc, rs)
      ibc.value(nsvan_item_db) = .Fields("ITEM")
      ibc.value(nsvan_AvailableFrom_db) = .Fields("AvailFrom")
      ibc.value(nsvan_Availableto_db) = .Fields("AvailTo")
      ibc.value(nsvan_DaysUnavailable_db) = .Fields("Unavail")
      
      If Not p11d32.BringForward.Yes Then
        ibc.value(nsvan_commuter_use_req_db) = .Fields("CommuterUse")
        ibc.value(nsvan_availablle_for_bus_use_only) = .Fields("BusReqOnly")
      End If
      
      'cad 2010 is_electric
      If (p11d32.BringForward.Yes) And (p11d32.AppYear = 2010) Then
        ibc.value(nsvan_is_electric_db) = False
      Else
        ibc.value(nsvan_is_electric_db) = IsNullEx(rs.Fields("is_electric"), False)
      End If
      
      'ibc.value(nsvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear) = .Fields("RegAfter")
      ibc.value(nsvan_madegood_db) = .Fields("PrivateContrib")
      
      
      ibc.value(nsvan_RegistrationDate_db) = IIf(IsNull(rs.Fields("RegistrationDate").value), p11d32.Rates.value(VanRegDateNew), rs.Fields("RegistrationDate").value)
      If (p11d32.BringForward.Yes) And (p11d32.AppYear = 2007) Then
        ibc.value(nsvan_fuel_available_db) = False
      Else
        ibc.value(nsvan_fuel_available_db) = IIf(IsNull(rs.Fields("fuel_available").value), True, rs.Fields("fuel_available").value)
      End If
      
      
      
    End With
    IBenefitClass_ReadDB = 1
  End If
  
  IBenefitClass_ReadDB = True
  
clsLoan_ReadDB_end:
  Set rs = Nothing
  
  
  Call xReturn("clsLoan_ReadDB")
  Exit Function
  
clsLoan_ReadDB_Err:
  Resume clsLoan_ReadDB_end
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
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
  bc = BC_nonSHAREDVAN_G
  m_NeedToCalculate = True
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
    
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  
  With p11d32
    Set ben = Me
    'MMField sizes
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    .BenDataLinkDataType(bc, nsvan_AvailableFrom_db) = TYPE_DATE
    .BenDataLinkDataType(bc, nsvan_Availableto_db) = TYPE_DATE
    .BenDataLinkDataType(bc, nsvan_DaysUnavailable_db) = TYPE_LONG
    .BenDataLinkDataType(bc, nsvan_commuter_use_req_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, nsvan_availablle_for_bus_use_only) = TYPE_BOOL
    
'MP DB    .BenDataLinkDataType(bc, nsvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear) = TYPE_BOOL
'    If p11d32.AppYear > 2000 Then
    .BenDataLinkDataType(bc, nsvan_RegistrationDate_db) = TYPE_DATE 'km
    .BenDataLinkDataType(bc, nsvan_is_electric_db) = TYPE_BOOL 'cad 2010

    .BenDataLinkDataType(bc, nsvan_fuel_available_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, nsvan_fuel_benefit) = TYPE_LONG
    
    .BenDataLinkDataType(bc, nsvan_fuel_value) = TYPE_LONG
    .BenDataLinkDataType(bc, nsvan_benefit_van_only) = TYPE_LONG

    
    
 End With
 
 ' REMOVED AS WE REPORT ONLY ON 'NON SHARED VANS' Call AbacusUDMData(bc) Call AbacusUDMData(bc)
  
 
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub
Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_VANS
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_nonSHAREDVAN_G, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean

  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo NonSharedVan_WriteDB_Err
  Call xSet("NonSharedVan_WriteDB")
  
  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  
  Call BringForwardDatesWrite(ben, nsvan_AvailableFrom_db, nsvan_Availableto_db, rs, "AvailFrom", "AvailTo")
  
  Call StandardWriteData(ben, rs)
  
  rs.Fields("ITEM").value = ben.value(nsvan_item_db)
  rs.Fields("RegistrationDate").value = ben.value(nsvan_RegistrationDate_db)
  
  rs.Fields("is_electric").value = ben.value(nsvan_is_electric_db)

  If Not p11d32.BringForward.Yes Then
    rs.Fields("Unavail").value = ben.value(nsvan_DaysUnavailable_db)
    rs.Fields("PrivateContrib").value = ben.value(nsvan_madegood_db)
    rs.Fields("BusReqOnly").value = ben.value(nsvan_availablle_for_bus_use_only)
    rs.Fields("CommuterUse").value = ben.value(nsvan_commuter_use_req_db)
    'rs.Fields("fuel_available").value = ben.value(nsvan_fuel_available_db)
  End If
  rs.Fields("fuel_available").value = ben.value(nsvan_fuel_available_db)
  
  
  IBenefitClass_WriteDB = BenefitCloseRecord(ben, rs)
  
NonSharedVan_WriteDB_End:
  Set rs = Nothing
  Set ben = Nothing
  Call xReturn("NonSharedVan_WriteDB")
  Exit Function
NonSharedVan_WriteDB_Err:
  Set ben = Nothing
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVan_WriteDB", "Non Shared Van WriteDB", "Error writing a non shared van benefit to the database.")
  Resume NonSharedVan_WriteDB_End
Resume
End Function

