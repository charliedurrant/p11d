VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Option Explicit
Implements IBenefitClass

Public Enum PhoneItems
  pho_item_db = ITEM_DESC
  pho_Value = ITEM_VALUE
  pho_MadeGood_Net = ITEM_MADEGOOD_NET
  pho_Benefit = ITEM_BENEFIT
  pho_MadeGood = ITEM_MADEGOOD
  pho_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  pho_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  pho_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  pho_BoxNumber = ITEM_BOX_NUMBER
  pho_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  pho_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  pho_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  pho_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  pho_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  pho_Error = ITEM_ERROR
  pho_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  pho_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  pho_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE
  
  
  Pho_RentValue_db
  Pho_RentMadeGood_db
  Pho_CallsMadeGood_db
  pho_MadeGoodNet_Calls
  pho_MadeGoodNet_Rent
  pho_Benefit_Rent
  pho_Benefit_Calls
  Pho_CallsValue_db
'MP DB (never displayed)  pho_unavailable_db
'MP DB  pho_Reduction
'MP DB (not used)             pho_TotalUnavailable
    
'MP DB (never displayed)  Pho_PrivateUse_db
'MP DB (never displayed)      Pho_reqmakegood_db
'MP DB (never displayed)      Pho_ActMadeGood_db
'MP DB (never displayed)      Pho_CapMadeGood_db
'MP DB (never displayed)        pho_availablefrom_db 'BF
'MP DB (never displayed)        pho_availableto_db 'BF
'MP DB  Pho_LASTITEM = pho_availableto_db
  Pho_LASTITEM = Pho_CallsValue_db
End Enum

Private m_Parent As Object
Private m_BenClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_CompanyDefined  As Boolean
Private m_BenItems(1 To Pho_LASTITEM)
Private m_dirty As Boolean

Private m_InvalidFields As Long
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Phones - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim l As Long
  Dim i As Long
  Dim d As Double
  Dim ben As IBenefitClass
  
  On Error GoTo Phone_Calculate_err
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  
  Select Case ben.BenefitClass
    
    Case BC_PHONE_HOME_N
      
      ben.value(pho_Value_Non_OPRA) = ben.value(Pho_CallsValue_db) + ben.value(Pho_RentValue_db)
      Call CalculateOpRAValue(ben)
      
      
      ben.value(pho_MadeGoodNet_Calls) = Min(ben.value(Pho_CallsMadeGood_db), ben.value(Pho_CallsValue_db))
      ben.value(pho_MadeGoodNet_Rent) = Min(ben.value(Pho_RentMadeGood_db), ben.value(Pho_RentValue_db))
      ben.value(pho_MadeGood_Net) = ben.value(pho_MadeGoodNet_Calls) + ben.value(pho_MadeGoodNet_Rent)
            
      ben.value(pho_Benefit_Calls) = ben.value(Pho_CallsValue_db) - ben.value(pho_MadeGoodNet_Calls)
      ben.value(pho_Benefit_Rent) = ben.value(Pho_RentValue_db) - ben.value(pho_MadeGoodNet_Rent)
            
      ben.value(pho_MadeGood) = ben.value(Pho_CallsMadeGood_db) + ben.value(Pho_RentMadeGood_db)
      
      ben.value(pho_Benefit) = ben.value(pho_Value) - ben.value(pho_MadeGood_Net)
      If ben.value(pho_Value) > 0 Then ben.value(pho_Benefit_Reportable) = True
    Case Else
      ECASE "Unknown Phone class"
  End Select
  
  Call BenCalcNIC(ben)
  IBenefitClass_CalculateBody = ben.value(pho_Benefit)
  
Phone_Calculate_end:
  Exit Function

Phone_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume Phone_Calculate_end
  Resume

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Phone"
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
'MP DB ToDO MUST - pho_availableto_db by default get allocated end of tax year date
' and never get modified, so always below call will always return True
'MP DB  IBenefitClass_CanBringForward = StandardCanBringForward(Me, pho_availableto_db)
  IBenefitClass_CanBringForward = True
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim phone As IBenefitClass
  
  Set phone = New phone
  
  Set IBenefitClass_Copy = CopyBenStandard(Parent, phone, Me)
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = m_BenClass
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


Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo ServicesProvided_PrintWKBody_Err
  Call xSet("ServicesProvided_PrintWKBody")
  
  Set ben = Me
  
  Call WKOut(rep, WK_ITEM_TEXT, "Description: " & S_WK_NORMAL_BOLD_FONT & ben.value(ITEM_DESC))
    
  Call WKTblOtherTypeTable(rep, "", "Value", "Made good", "Benefit")
  Call WKTableRow(rep, "Rent", FormatWN(ben.value(Pho_RentValue_db)), FormatWN(ben.value(pho_MadeGoodNet_Rent), , True), FormatWN(ben.value(pho_Benefit_Rent)))
  Call WKTableRow(rep, "Calls", FormatWN(ben.value(Pho_CallsValue_db)), FormatWN(ben.value(pho_MadeGoodNet_Calls), , True), FormatWN(ben.value(pho_Benefit_Calls)))
  
  'Call WKOut(rep, WK_ITEM_Total, "Total (Rent + Calls)", ben.value(ITEM_BENEFIT), , True)
    
  Call WKTableTotalBen(rep, ben)
      
    
ServicesProvided_PrintWKBody_End:
  Call xReturn("ServicesProvided_PrintWKBody")
  Exit Function

ServicesProvided_PrintWKBody_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ServicesProvided_PrintWKBody", "Services Provided Print WK Body", "Error printing the services provided working paper")
  Resume ServicesProvided_PrintWKBody_End
  Resume
End Function

Private Sub IBenefitClass_SetCalcDefaults()
  m_BenItems(pho_Value) = S_ERROR
  m_BenItems(pho_MadeGood) = S_ERROR
  m_BenItems(pho_MadeGood_Net) = S_ERROR
'MP DB  m_BenItems(pho_Reduction) = S_ERROR
  m_BenItems(pho_Benefit) = S_ERROR
  m_BenItems(pho_Benefit_Reportable) = False
  m_BenItems(pho_Error) = ""
  m_BenItems(pho_OPRA_Ammount_Foregone_Used_For_Value) = False
End Sub

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(m_BenClass, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property


Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  
  On Error GoTo SetBenItemsInformation_err
  
  m_NeedToCalculate = True
  
  bc = BC_PHONE_HOME_N
    
    With p11d32
      If .DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
    
      Call SetStandardBenItemsInformation(bc, Me)
      
      .BenDataLinkDataType(bc, pho_MadeGoodNet_Calls) = TYPE_LONG
      .BenDataLinkDataType(bc, pho_MadeGoodNet_Rent) = TYPE_LONG
      .BenDataLinkDataType(bc, pho_Benefit_Rent) = TYPE_LONG
      .BenDataLinkDataType(bc, pho_Benefit_Calls) = TYPE_LONG
      .BenDataLinkDataType(bc, Pho_RentValue_db) = TYPE_LONG
      .BenDataLinkDataType(bc, Pho_RentMadeGood_db) = TYPE_LONG
      .BenDataLinkDataType(bc, Pho_CallsMadeGood_db) = TYPE_LONG
      .BenDataLinkDataType(bc, Pho_CallsValue_db) = TYPE_LONG
'MP DB (never displayed)          .BenDataLinkDataType(bc, pho_unavailable_db) = TYPE_LONG
'MP DB      .BenDataLinkDataType(bc, pho_Reduction) = TYPE_LONG
'MP DB      .BenDataLinkDataType(bc, pho_TotalUnavailable) = TYPE_LONG
'MP DB (never displayed)          .BenDataLinkDataType(bc, Pho_PrivateUse_db) = TYPE_BOOL
'MP DB (never displayed)          .BenDataLinkDataType(bc, Pho_reqmakegood_db) = TYPE_BOOL
'MP DB (never displayed)          .BenDataLinkDataType(bc, Pho_ActMadeGood_db) = TYPE_BOOL
'MP DB (never displayed)          .BenDataLinkDataType(bc, Pho_CapMadeGood_db) = TYPE_BOOL
'MP DB (never displayed)          .BenDataLinkDataType(bc, pho_availablefrom_db) = TYPE_DATE
'MP DB (never displayed)          .BenDataLinkDataType(bc, pho_availableto_db) = TYPE_DATE
  
      .BenDataLinkUDMDescription(bc, Pho_RentValue_db) = "Rent"
      .BenDataLinkUDMDescription(bc, Pho_RentMadeGood_db) = "Rent made good"
      .BenDataLinkUDMDescription(bc, Pho_CallsValue_db) = "Calls"
      .BenDataLinkUDMDescription(bc, Pho_CallsMadeGood_db) = "Calls made good"
      
      .DataLinkInitialised(bc) = True
    End With

  
SetBenItemsInformation_end:
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub


Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub


Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(pho_item_db)
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

Private Function IBenefitClass_ReadDB() As Long
  Dim phone As IBenefitClass
  Dim ben As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim sIndex As String
  Dim i As Long
  
  On Error GoTo Phone_ReadDB_err
  xSet "Phone_ReadDB"
  If m_ReadFromDB Then GoTo Phone_ReadDB_end
  
  Set rs = m_Parent.Parent.rsBenTables(TBL_PHONES)
  
  Set ben = Me
  
  If Len(m_sbookmark) = 0 Then
    s = m_Parent.PersonnelNumber
    rs.FindFirst ("P_NUM = '" & s & "'")
    If Not rs.NoMatch Then
      i = i + 1
      m_sbookmark = rs.Bookmark
      rs.FindNext ("P_NUM = '" & s & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set phone = New phone
        phone.RSBookMark = rs.Bookmark
        phone.CompanyDefined = m_CompanyDefined
        Set phone.Parent = m_Parent
        Call m_Parent.benefits.Add(phone)
        Set phone = Nothing
        rs.FindNext ("P_NUM = '" & s & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    rs.Bookmark = m_sbookmark

    m_BenClass = p11d32.Rates.DBClassToBCClass(CStr("" & rs.Fields("Class").value))
    Call StandardReadData(ben, rs)
    ben.value(Pho_RentValue_db) = IIf(IsNull(rs.Fields("ValueRent").value), 0, rs.Fields("ValueRent").value)
    ben.value(Pho_RentMadeGood_db) = IIf(IsNull(rs.Fields("ContribRent").value), 0, rs.Fields("ContribRent").value)
    ben.value(Pho_CallsValue_db) = IIf(IsNull(rs.Fields("ValueCalls").value), 0, rs.Fields("ValueCalls").value)
    ben.value(Pho_CallsMadeGood_db) = IIf(IsNull(rs.Fields("ContribCalls").value), 0, rs.Fields("ContribCalls").value)
'MP DB (never displayed)
'MP DB    ben.value(pho_unavailable_db) = IIf(IsNull(rs.Fields("Unavail").value), 0, rs.Fields("Unavail").value)
            
    ben.value(pho_item_db) = "" & rs.Fields("Item").value
'MP DB (never displayed)    ben.value(Pho_PrivateUse_db) = rs.Fields("PvtUse").value
'MP DB (never displayed)        ben.value(Pho_reqmakegood_db) = rs.Fields("MakeGood").value
'MP DB (never displayed)        ben.value(Pho_ActMadeGood_db) = rs.Fields("MadeGood").value
'MP DB (never displayed)        ben.value(Pho_CapMadeGood_db) = rs.Fields("CapMadeGood").value
'MP DB (never displayed)        ben.value(pho_availablefrom_db) = IIf(IsNull(rs.Fields("AvailFrom").value), p11d32.Rates.value(TaxYearStart), rs.Fields("AvailFrom").value)
'MP DB (never displayed)        ben.value(pho_availableto_db) = IIf(IsNull(rs.Fields("AvailTo").value), p11d32.Rates.value(TaxYearEnd), rs.Fields("AvailTo").value)
    m_ReadFromDB = True
  End If
Phone_ReadDB_end:
  IBenefitClass_ReadDB = i
  xReturn "Phone_ReadDB"
  Exit Function
Phone_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "Phone_ReadDB", "ERR_READ_Phone_BENEFITS", "Unable to read in the Phone benefits"
   Resume Phone_ReadDB_end
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

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo Phone_DeleteDB_Err
  Call xSet("Phone_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_PHONES)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
Phone_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("Phone_DeleteDB")
  Exit Function

Phone_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Phone_DeleteDB", "Phone Delete DB", "Error deleting the phone benefit.")
  Resume Phone_DeleteDB_End
End Function


Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo Phone_WriteDB_Err
  Call xSet("Phone_WriteDB")

  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields("Class").value = p11d32.Rates.BenClassTo(m_BenClass, BCT_DBCLASS)
  
  rs.Fields("Item").value = ben.value(pho_item_db)
  
  Call StandardWriteData(ben, rs)
  
  If Not p11d32.BringForward.Yes Then
    rs.Fields("ValueRent").value = ben.value(Pho_RentValue_db)
    rs.Fields("ContribRent").value = ben.value(Pho_RentMadeGood_db)
    rs.Fields("ValueCalls").value = ben.value(Pho_CallsValue_db)
    rs.Fields("ContribCalls").value = ben.value(Pho_CallsMadeGood_db)
  End If
  
  IBenefitClass_WriteDB = BenefitCloseRecord(ben, rs)
Phone_WriteDB_End:
  Set rs = Nothing
  Call xReturn("Phone_WriteDB")
  Exit Function
Phone_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "Phone_WriteDB", "Phone Write DB", "Error writing the phone benefit to the database.")
  Resume Phone_WriteDB_End
  Resume
End Function

Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_PHONES
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property





