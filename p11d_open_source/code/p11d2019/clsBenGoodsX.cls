VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServicesProvided"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IBenefitClass
Public Enum ServicesProvidedItems
  ServicesProvided_item_db = ITEM_DESC 'OK
  ServicesProvided_value = ITEM_VALUE 'OK
  ServicesProvided_MadeGood_Net = ITEM_MADEGOOD_NET
  ServicesProvided_Benefit = ITEM_BENEFIT 'OK
  ServicesProvided_MadeGood_db = ITEM_MADEGOOD 'OK
  ServicesProvided_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  ServicesProvided_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  ServicesProvided_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  ServicesProvided_BoxNumber = ITEM_BOX_NUMBER
  ServicesProvided_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  ServicesProvided_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  ServicesProvided_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  ServicesProvided_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  ServicesProvided_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  ServicesProvided_ERROR = ITEM_ERROR
  ServicesProvided_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  ServicesProvided_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  ServicesProvided_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE

 ' ServicesProvided_Class_db
  ServicesProvided_MarketValue_db 'OK
  ServicesProvided_Marginal_db 'OK
  
  ServicesProvided_LASTITEM = ServicesProvided_Marginal_db
End Enum

Private m_Parent As Object
Private m_BenClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_CompanyDefined  As Boolean
Private m_BenItems(1 To ServicesProvided_LASTITEM)

Private m_dirty As Boolean

Private m_InvalidFields As Long
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Service Provided - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim ben As IBenefitClass
  
  On Error GoTo Goods_Calculate_err
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  
  ben.value(ServicesProvided_Value_Non_OPRA) = ben.value(ServicesProvided_MarketValue_db) + ben.value(ServicesProvided_Marginal_db)
  ben.value(ServicesProvided_MadeGood_Net) = Min(ben.value(ServicesProvided_value), ben.value(ServicesProvided_MadeGood_db))
  Call CalculateOpRAValue(ben)
  ben.value(ServicesProvided_Benefit) = ben.value(ServicesProvided_value) - ben.value(ServicesProvided_MadeGood_Net)
  
  
  IBenefitClass_CalculateBody = ben.value(ServicesProvided_Benefit)
  If ben.value(ServicesProvided_value) > 0 Then ben.value(ServicesProvided_Benefit_Reportable) = True
  Call BenCalcNIC(ben)
  
Goods_Calculate_end:
  Exit Function
Goods_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume Goods_Calculate_end

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
  IBenefitClass_CanBringForward = True
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim serv As ServicesProvided
  
  Set serv = New ServicesProvided
  Set IBenefitClass_Copy = CopyBenStandard(Parent, serv, Me)
  
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_SERVICES_PROVIDED_K
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
  
  'PS TTP#84
  Call WKOut(rep, WK_ITEM_TEXT, "Description")
  Call WKTblOtherTypeTable(rep, "", "Value", "Made good", "Benefit")
  
  Call WKTableRow(rep, ben.value(ITEM_DESC), FormatWN(ben.value(ITEM_VALUE)), FormatWN(ben.value(ITEM_MADEGOOD_NET), , True), FormatWN(ben.value(ITEM_BENEFIT)))
  
  Call WKTableTotalBen(rep, ben)
  
  'Call WKOut(rep, WK_ITEM_Total, "Total", ben.value(ITEM_BENEFIT), , True)
      
ServicesProvided_PrintWKBody_End:
  Call xReturn("ServicesProvided_PrintWKBody")
  Exit Function

ServicesProvided_PrintWKBody_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ServicesProvided_PrintWKBody", "Services Provided Print WK Body", "Error printing the services provided working paper")
  Resume ServicesProvided_PrintWKBody_End
  Resume
  

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  m_BenItems(ServicesProvided_value) = S_ERROR
  m_BenItems(ServicesProvided_Benefit) = S_ERROR
  m_BenItems(ServicesProvided_MadeGood_Net) = S_ERROR
  m_BenItems(ServicesProvided_Benefit_Reportable) = False
  m_BenItems(ServicesProvided_ERROR) = ""
  m_BenItems(ServicesProvided_OPRA_Ammount_Foregone_Used_For_Value) = False
End Sub

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_SERVICES_PROVIDED_K, Item, RHS)
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
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(ServicesProvided_item_db)
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
  Dim ServicesProvided As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim sIndex As String
  Dim i As Long
  
  On Error GoTo ServicesProvided_ReadDB_err
  xSet "ServicesProvided_ReadDB"
  If m_ReadFromDB Then GoTo ServicesProvided_ReadDB_end
  Set rs = m_Parent.Parent.rsBenTables(TBL_SERVICESPROVIDED)
  If Len(m_sbookmark) = 0 Then
    s = m_Parent.PersonnelNumber
    rs.FindFirst ("P_NUM = '" & s & "'")
    If Not rs.NoMatch Then
      i = i + 1
      m_sbookmark = rs.Bookmark
      rs.FindNext ("P_NUM = '" & s & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set ServicesProvided = New ServicesProvided
        ServicesProvided.RSBookMark = rs.Bookmark
        ServicesProvided.CompanyDefined = m_CompanyDefined
        Set ServicesProvided.Parent = m_Parent
        Call m_Parent.benefits.Add(ServicesProvided)
        Set ServicesProvided = Nothing
        rs.FindNext ("P_NUM = '" & s & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    Set ServicesProvided = Me
    rs.Bookmark = m_sbookmark
    Call StandardReadData(ServicesProvided, rs)
    
    ServicesProvided.value(ServicesProvided_item_db) = "" & rs.Fields("Item").value
    ServicesProvided.value(ServicesProvided_MarketValue_db) = IIf(IsNull(rs.Fields("MKTVAL").value), 0, rs.Fields("MKTVAL").value)
    ServicesProvided.value(ServicesProvided_Marginal_db) = IIf(IsNull(rs.Fields("MARGINAL").value), 0, rs.Fields("MARGINAL").value)
    ServicesProvided.value(ServicesProvided_MadeGood_db) = IIf(IsNull(rs.Fields("MADEGOOD").value), 0, rs.Fields("MADEGOOD").value)
    
    'If p11d32.BringForward.Yes And p11d32.AppYear = 2000 Then
    '  ServicesProvided.value(ServicesProvided_Class_db) = "Services Provided"
   ' Else
    '  ServicesProvided.value(ServicesProvided_Class_db) = "" & rs.Fields("Class").value
   ' End If

    m_ReadFromDB = True
  End If
ServicesProvided_ReadDB_end:
  IBenefitClass_ReadDB = i
  xReturn "ServicesProvided_ReadDB"
  Exit Function
ServicesProvided_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "ServicesProvided_ReadDB", "ERR_READ_ServicesProvided_BENEFITS", "Unable to read in the Services / Assets benefits"
   Resume ServicesProvided_ReadDB_end
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
  Dim i As Long
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
  
  bc = BC_SERVICES_PROVIDED_K
  m_NeedToCalculate = True
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  Set ben = Me
  
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  Call SetStandardBenItemsUDMData(bc)
  
  With p11d32
    'MMField sizes
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    .BenDataLinkDataType(bc, ServicesProvided_MarketValue_db) = TYPE_LONG
    .BenDataLinkDataType(bc, ServicesProvided_Marginal_db) = TYPE_LONG
  End With
  
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub
Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo SharedVan_DeleteDB_Err
  Call xSet("SharedVan_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_SERVICESPROVIDED)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
SharedVan_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("SharedVan_DeleteDB")
  Exit Function

SharedVan_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SharedVan_DeleteDB", "Shared Van Delete DB", "Error deleting a shared van.")
  Resume SharedVan_DeleteDB_End
End Function


Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo ServicesProvided_WriteDB_Err
  Call xSet("ServicesProvided_WriteDB")
  
  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields("ITEM").value = ben.value(ServicesProvided_item_db)
  rs.Fields("UDBCODE").value = S_SERVICESPROVIDED_UDBCODE
  'rs.Fields("Class").value = ben.value(ServicesProvided_Class_db)
  
  Call StandardWriteData(ben, rs)
  If Not p11d32.BringForward.Yes Then
    rs.Fields("MKTVAL").value = ben.value(ServicesProvided_MarketValue_db)
    rs.Fields("MARGINAL").value = ben.value(ServicesProvided_Marginal_db)
    rs.Fields("MADEGOOD").value = ben.value(ServicesProvided_MadeGood_db)
  End If
  
  
  IBenefitClass_WriteDB = BenefitCloseRecord(ben, rs)
  
ServicesProvided_WriteDB_End:
  
  Call xReturn("ServicesProvided_WriteDB")
  Exit Function
ServicesProvided_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "ServicesProvided_WriteDB", "Services Provided WriteDB", "Error writing the services provided benefit.")
  Resume ServicesProvided_WriteDB_End
End Function


Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_SERVICESPROVIDED
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property



