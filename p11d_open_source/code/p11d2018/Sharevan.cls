VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SharedVan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Implements IBenefitClass

Public Enum ShVanItems
  shvan_item_db = ITEM_DESC 'RW
  shvan_Value = ITEM_VALUE
  shvan_MadeGood_Net = ITEM_MADEGOOD_NET
  shvan_benefit = ITEM_BENEFIT 'W
  shvan_MadeGood = ITEM_MADEGOOD
  shvan_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  shvan_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  shvan_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  shvan_BoxNumber = ITEM_BOX_NUMBER
  shvan_MadeGoodIsTaxDeducted = ITEM_MADEGOOD_IS_TAXDEDUCTED
  shvan_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  shvan_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  shvan_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  shvan_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  shvan_Error = ITEM_ERROR
  shvan_Value_Non_OPRA = ITEM_VALUE_NON_OPRA 'not used as this is done at the nonsharedvans level, added as standard
  shvan_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE 'not used added as standard
  shvan_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE 'not used added as standard

  
  shvan_AvailableFrom_db 'RW
  shvan_Availableto_db 'RW
  shvan_DaysUnavailable_db 'RW
  shvan_fuel_available_db
  shvan_fuel_value
  shvan_fuel_benefit
  shvan_benefit_van_only
  shvan_is_electric_db
  shvan_RegistrationDate_db  'km
  
'MP DB not used  shvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear_db 'RW
  
'MP DB  shvan_LASTITEM = shvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear_db
  shvan_LASTITEM = shvan_RegistrationDate_db
End Enum

Private m_InvalidFields As Long
Private m_BenClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_Parent As SharedVans
Private m_BenItems(1 To shvan_LASTITEM)

Private m_dirty As Boolean
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Shared Van - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim lDaysUnavailable As Long
  Dim dDateFactor As Double
  Dim ben As IBenefitClass
  
  On Error GoTo SharedVan_Calculate_err
  
  Set ben = Me
  
  dDateFactor = dGetDateFactor(lDaysUnavailable, ben.value(shvan_AvailableFrom_db), ben.value(shvan_Availableto_db), ben.value(shvan_DaysUnavailable_db), 30)
  If ben.value(shvan_RegistrationDate_db) >= p11d32.Rates.value(carFOURYEAROLD) Then
    ben.value(shvan_Value) = p11d32.Rates.value(vanHighCharge) * dDateFactor
  Else
    ben.value(shvan_Value) = p11d32.Rates.value(vanLowCharge) * dDateFactor
  End If
  If (ben.value(shvan_is_electric_db)) Then
    ben.value(shvan_Value) = ben.value(shvan_Value) * p11d32.Rates.value(vanElectricMultiplier)
    
  End If

  If (ben.value(shvan_fuel_available_db)) And Not (ben.value(shvan_is_electric_db)) Then 'cad 2010 is_electric
    ben.value(shvan_fuel_value) = p11d32.Rates.value(vanRate) * dDateFactor 'cad 2010
  Else
    ben.value(shvan_fuel_value) = 0
  End If

  

 
  ben.value(shvan_fuel_benefit) = ben.value(shvan_fuel_value)
  
  ben.value(shvan_benefit_van_only) = ben.value(shvan_Value)
  ben.value(shvan_Value_Non_OPRA) = ben.value(shvan_Value)
  ben.value(shvan_Value) = ben.value(shvan_Value) + ben.value(shvan_fuel_value)
  ben.value(shvan_benefit) = ben.value(shvan_Value)
 
  If ben.value(shvan_benefit) > 0 Then ben.value(shvan_Benefit_Reportable) = True
  Call BenCalcNIC(ben)
    
  IBenefitClass_CalculateBody = ben.value(shvan_benefit)
  
SharedVan_Calculate_end:
  Exit Function
SharedVan_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume SharedVan_Calculate_end
  Resume

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "SharedVan"
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
  'IBenefitClass_CanBringForward = True
  'PS Added
  IBenefitClass_CanBringForward = StandardCanBringForward(Me, shvan_Availableto_db)

End Function
Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function
Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_SHAREDVAN_G
End Property
Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "SharedVan_CompanyDefined"
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  
End Property

Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim ibc As IBenefitClass
  
  On Error GoTo SharedVan_DeleteDB_Err
  
  Call xSet("SharedVan_DeleteDB")
  
  Set ibc = Me
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(ibc.TABLE)
  
  If ibc.HasBookMark Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  
  IBenefitClass_DeleteDB = True
  
SharedVan_DeleteDB_End:
  Set ibc = Nothing
  Set rs = Nothing
  Call xReturn("SharedVan_DeleteDB")
  Exit Function
SharedVan_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SharedVan_DeleteDB", "Shared Van Delete DB", "Error deleting a shared van from the database.")
  Resume SharedVan_DeleteDB_End

End Function

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  Dim ben As IBenefitClass
  
  m_dirty = DirtyHelper(Me, NewValue)
  Set ben = m_Parent
  ben.Dirty = ben.Dirty Or NewValue
  
  Set ben = Nothing
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  m_BenItems(shvan_Value) = S_ERROR
  m_BenItems(shvan_benefit) = S_ERROR
  m_BenItems(shvan_MadeGood_Net) = 0
  m_BenItems(shvan_MadeGood) = 0
  m_BenItems(shvan_Benefit_Reportable) = False
  m_BenItems(shvan_Error) = ""
End Sub

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_SHAREDVAN_G, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property
Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(shvan_item_db)
End Property

Private Property Get IBenefitClass_PrintHeader() As String
  ECASE ("IBenefitClass_PrintHeader")
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call m_Parent.IBenefitClass_PrintWk(rep)
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo SharedVan_ReadDB_ERR
  
  Call xSet("SharedVan_ReadDB")
  
  If m_ReadFromDB Then GoTo SharedVan_ReadDB_END
  Set ben = Me
  If Not ben.HasBookMark Then
    'err.raise
  Else
    
    Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_SHARED_VANS)
    With rs
      rs.Bookmark = ben.RSBookMark
      ben.value(shvan_item_db) = .Fields("ITEM")
      ben.value(shvan_AvailableFrom_db) = .Fields("AvailFrom")
      ben.value(shvan_Availableto_db) = .Fields("AvailTo")
      ben.value(shvan_DaysUnavailable_db) = .Fields("Unavail")
      ben.value(shvan_RegistrationDate_db) = IIf(.Fields("RegistrationDate") = Null, p11d32.Rates.value(VanRegDateNew), .Fields("RegistrationDate"))
      
      If (p11d32.BringForward.Yes) And (p11d32.AppYear = 2007) Then
        ben.value(shvan_fuel_available_db) = False
      Else
        ben.value(shvan_fuel_available_db) = IIf(IsNull(.Fields("fuel_available")), False, .Fields("fuel_available"))
      End If
    
      If (p11d32.BringForward.Yes) And (p11d32.AppYear = 2010) Then 'cAD 2010
        ben.value(shvan_is_electric_db) = False
      Else
        ben.value(shvan_is_electric_db) = IIf(IsNull(.Fields("is_electric")), False, .Fields("is_electric"))
      End If
    
    

      IBenefitClass_ReadFromDB = True
      IBenefitClass_ReadDB = 1
    End With
  End If
    
SharedVan_ReadDB_END:
  Set ben = Nothing
  Set rs = Nothing
  Call xReturn("SharedVan_ReadDB")
  Exit Function
SharedVan_ReadDB_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SharedVan_ReadDB", "SharedVan ReadDB", "Unable to read a the shared vans")
  Resume SharedVan_ReadDB_END
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
  
  bc = BC_SHAREDVAN_G
  
  m_NeedToCalculate = True
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  ' REMOVED AS WE REPORT ONLY ON 'NON SHARED VANS' Call AbacusUDMData(bc)
    
  With p11d32
    Set ben = Me
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    .BenDataLinkDataType(bc, shvan_AvailableFrom_db) = TYPE_DATE
    .BenDataLinkDataType(bc, shvan_Availableto_db) = TYPE_DATE
    .BenDataLinkDataType(bc, shvan_DaysUnavailable_db) = TYPE_LONG
'MP DB    .BenDataLinkDataType(bc, shvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear_db) = TYPE_BOOL
'    If p11d32.AppYear > 2000 Then
    .BenDataLinkDataType(bc, shvan_RegistrationDate_db) = TYPE_DATE 'km
    .BenDataLinkDataType(bc, shvan_is_electric_db) = TYPE_BOOL 'cad 2010
    
    .BenDataLinkDataType(bc, shvan_fuel_available_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, shvan_fuel_benefit) = TYPE_LONG
    .BenDataLinkDataType(bc, shvan_fuel_value) = TYPE_LONG
 

  End With

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub
Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim ey As Employer
  
  On Error GoTo SharedVan_WriteDB_Err
  Call xSet("SharedVan_WriteDB")
  
  Set ben = Me
  
  If ben.TABLE = 0 Then Call Err.Raise(ERR_IS_NOTHING, "BenefitAddNewRecord", "Table for benefit is 0 length string")
  
  Set ey = GetParentFromBenefit(ben, GPBF_EMPLOYER)
  Set rs = ey.rsBenTables(ben.TABLE)
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitAddNewRecord", "Recordset is nothing")
  
  If (p11d32.BringForward.Yes) Then
    If (ben.CanBringForward) Then
      rs.AddNew
      rs.Fields(S_FIELD_PERSONEL_NUMBER).value = S_SHAREDVAN_KEY
    Else
      GoTo SharedVan_WriteDB_End
    End If
  Else
    If (Len(ben.RSBookMark) = 0) Then
      rs.AddNew
      rs.Fields(S_FIELD_PERSONEL_NUMBER).value = S_SHAREDVAN_KEY
    Else
    
      rs.Bookmark = ben.RSBookMark
      rs.Edit
    End If
  End If
  
  
  rs.Fields("ITEM").value = ben.value(shvan_item_db)
  
  Call BringForwardDatesWrite(ben, shvan_AvailableFrom_db, shvan_Availableto_db, rs, "AvailFrom", "AvailTo")
  
  If Not p11d32.BringForward.Yes Then
    rs.Fields("Unavail").value = ben.value(shvan_DaysUnavailable_db)
  End If
  
  rs.Fields("RegistrationDate").value = ben.value(shvan_RegistrationDate_db)
  rs.Fields("is_electric").value = ben.value(shvan_is_electric_db) 'CAD 2010
   
  
  rs.Fields("fuel_available").value = ben.value(shvan_fuel_available_db)

  IBenefitClass_WriteDB = BenefitCloseRecord(ben, rs)
  
SharedVan_WriteDB_End:
  Call xReturn("SharedVan_WriteDB")
  Exit Function
SharedVan_WriteDB_Err:

  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "SharedVan_WriteDB", "Shared Van Write DB", "Error writing a shared van to the database.")
  Resume SharedVan_WriteDB_End
  Resume
End Function

Public Function PrintWk(rep As Reporter) As Boolean
  
'  On Error GoTo printwk_Err
'  Call xSet("printwk")
'  Call rep.InitReport("Van Report", gtarget, PORTRAIT, False, False)
'
'  If m_BenItems(shvan_Multiple)  Then
'    rep.Out "{Arial=8,b}{x=0}More than one van was available for private use at the same time in the tax year" & vbCrLf & vbCrLf
'  Else
'    rep.Out "{Arial=8,b}{x=0}Only one van was available for private use at any time in the tax year" & vbCrLf & vbCrLf
'  End If
'  rep.Out "{LINE}" & vbCrLf
'  rep.Out "{Arial=8}{x=0}Shared van benefit" & vbCrLf & vbCrLf
'  If m_BenItems(shvan_ShowAlternative)  Then
'    rep.Out "{Arial=8}{x=0}(i)"
'  Else
'    rep.Out "{Arial=8}{x=0}"
'  End If
'  rep.Out "Benefit Calculated by the normal method" & vbCrLf & vbCrLf
'  rep.Out "{Arial=8,n}{x=0}Aggregate benefit of shared van pool"
'  rep.Out "{Arial=8,rb}{x=85}'A'{x=97}" & FormatWN(m_BenItems(shvan_TotalVan) , "£") & vbCrLf & vbCrLf
'  rep.Out "{Arial=8,n}{x=0}Number of employees using the pool{Arial=8,rb}{x=85}'B'{x=97}/ " & FormatWN(m_BenItems(shvan_Employees) , "") & vbCrLf
'  rep.Out "{x=97}{LINE=-7}" & vbCrLf
'  rep.Out "{Arial=8,n}{x=0}Individual benefit for van pool - 'A' / 'B'"
'  rep.Out "{Arial=8,rb}{x=97}" & FormatWN(m_BenItems(shvan_GrossBenefit) , "") & vbCrLf & vbCrLf
'  rep.Out "{Arial=8,n}{x=0}Less: Payments made as a condition of private use"
'  rep.Out "{Arial=8,rb}{x=97}" & FormatWN(m_BenItems(shvan_Contribution) , "", True) & vbCrLf
'  rep.Out "{x=97}{LINE=-7}" & vbCrLf
'  rep.Out "{Arial=8,n}{x=0}Cash Benefit"
'  rep.Out "{Arial=8,rb}{x=97}" & FormatWN(m_BenItems(shvan_benefit) , "£") & vbCrLf
'  rep.Out "{x=97}{LINE=-7,d}" & vbCrLf & vbCrLf
'  If m_BenItems(shvan_ShowAlternative)  Then
'    rep.Out "{Arial=8,nb}{x=0}(ii)Benefit calculated by the alternative daily method" & vbCrLf & vbCrLf
'    rep.Out "{Arial=8,n}{x=0}Daily calculation based on " & m_BenItems(shvan_Days)  & " days at " & FormatWN(P11d32.Rates.Value(vanDailyCharge), "£") & " per day"
'    rep.Out "{Arial=8,rb}{x=97}" & FormatWN(m_BenItems(shvan_DailyGross) , "£") & vbCrLf & vbCrLf
'    rep.Out "{Arial=8,n}{x=0}Less: Payments made as a condition of private use"
'    rep.Out "{Arial=8,rb}{x=97}" & FormatWN(m_BenItems(shvan_Contribution) , "", True) & vbCrLf
'    rep.Out "{x=97}{LINE=-7}" & vbCrLf
'    rep.Out "{Arial=8,n}{x=0}Benefit calculated by alternative method"
'    rep.Out "{Arial=8,rb}{x=97}" & FormatWN(m_BenItems(shvan_DailyBenefit) , "£") & vbCrLf
'    rep.Out "{x=97}{LINE=-7,d}" & vbCrLf & vbCrLf
'    rep.Out "{Arial=8,nb}{x=0}Notes" & vbCrLf
'    rep.Out "{Arial=8,ni}{x=0}" & "Where the alternative method is beneficial, a claim should be made on your tax return under Schedule 6a paragraph 8 of ICTA 1988." & vbCrLf & vbCrLf
'  Else
'    rep.Out "{Arial=8,nb}{x=0}Notes" & vbCrLf
'    rep.Out "{Arial=8,ni}It is possible to elect for a daily method of calculation of the van benefit where this gives a lower figure. This method is calculated by applying a daily charge" & vbCrLf & "of " & FormatWN(P11d32.Rates.Value(vanDailyCharge), "£") & " to the number of days the van is used privately.  This figure can be used where it is lower than the above figure." & vbCrLf
'    rep.Out "{Arial=8}A claim should be made on your tax return under Schedule 6A paragraph 8 of ICTA 1988." & vbCrLf & vbCrLf
'  End If
'  Call rep.EndReport
'printwk_End:
'  Call xReturn("printwk")
'  Exit Function
'printwk_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "printwk", "ERR_UNDEFINED", "Undefined error.")
'  Resume printwk_End
End Function
Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_SHARED_VANS
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property



