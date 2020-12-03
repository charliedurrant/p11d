VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssetsTransferred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IBenefitClass

Public Enum TransItems 'not bring forward
  trans_Item_db = ITEM_DESC
  trans_value_db = ITEM_VALUE           'MP DB ToDo - used in Calc but not displayed/modified - chk del?
  trans_MadeGood_Net = ITEM_MADEGOOD_NET
  trans_benefit = ITEM_BENEFIT
  Trans_MadeGood_db = ITEM_MADEGOOD
  trans_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  trans_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  trans_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  trans_BoxNumber = ITEM_BOX_NUMBER
  trans_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  trans_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  trans_NICClass1Aable = ITEM_NIC_CLASS1A_ABLE
  trans_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  trans_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  trans_ERROR = ITEM_ERROR
  trans_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  trans_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  trans_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE
  
  
  trans_IRDesc_db
    
  trans_MarketValueatTrans_db
  trans_MarketValueorig_db
  trans_BenefitAlready_db
  trans_EmployeeReference_db

  trans_IsCAr_db
  
  trans_LASTITEM = trans_IsCAr_db
End Enum

Private m_Parent As Object
Private m_BenClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_CompanyDefined  As Boolean
Private m_BenItems(1 To trans_LASTITEM)

Private m_dirty As Boolean

Private m_InvalidFields As Long
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Assets transferred - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim ben As IBenefitClass
  
  On Error GoTo AssetsTransferred_CalculateBody_err
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  
  If ben.value(trans_IsCAr_db) Then
    ben.value(ITEM_VALUE_NON_OPRA) = ben.value(trans_MarketValueatTrans_db)
  Else
    ben.value(ITEM_VALUE_NON_OPRA) = Max(ben.value(trans_MarketValueatTrans_db), ben.value(trans_MarketValueorig_db) - ben.value(trans_BenefitAlready_db))
  End If
  
  
  Call CalculateOpRAValue(ben)
  ben.value(trans_MadeGood_Net) = Min(ben.value(ITEM_VALUE), ben.value(Trans_MadeGood_db))
  
  ben.value(trans_benefit) = ben.value(ITEM_VALUE) - ben.value(trans_MadeGood_Net)
  
  If ben.value(ITEM_VALUE) > 0 Then ben.value(trans_Benefit_Reportable) = True
  Call BenCalcNIC(ben)
  IBenefitClass_CalculateBody = ben.value(trans_benefit)
  ' ben.value(
  
AssetsTransferred_CalculateBody_end:
  Exit Function
AssetsTransferred_CalculateBody_err:
  ben.value(trans_Benefit_Reportable) = True
  IBenefitClass_CalculateBody = S_ERROR
  Resume AssetsTransferred_CalculateBody_end

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
  'default to false
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim ass As AssetsTransferred
  
  Set ass = New AssetsTransferred
  Set IBenefitClass_Copy = CopyBenStandard(Parent, ass, Me)
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_ASSETSTRANSFERRED_A
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
  
  On Error GoTo AssetsTransferred_PrintWKBody_Err
  
  Call xSet("AssetsTransferred_PrintWKBody")
  
  Set ben = Me
  
  
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Description: ", ben.value(trans_Item_db) & vbCrLf)
  
  If ben.value(trans_IsCAr_db) Then
    Call WKOut(rep, WK_ITEM_TEXT, "The asset given or transferred is a car or has not been a benefit previously.")
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "The asset given or transferred is not a car and has been provided as a benefit previously.")
    Call rep.Out(vbCrLf)
    Call WKOut(rep, WK_ITEM_Total, "Market value when transferred", FormatWN(IIf(IsNull(ben.value(trans_MarketValueatTrans_db)), "0", ben.value(trans_MarketValueatTrans_db)), "£"), "'A'", True)
    Call rep.Out(vbCrLf)
    Call rep.Out(vbCrLf)
    Call WKOut(rep, WK_ITEM_TEXT, "Market value when first applied as a benefit", FormatWN(IIf(IsNull(ben.value(trans_MarketValueorig_db)), "0", ben.value(trans_MarketValueorig_db)), "£"))
    
    Call WKOut(rep, WK_ITEM_TEXT, "Less: Sum of taxable benefits already taken into account deriving from use of the asset", ben.value(trans_BenefitAlready_db), , , True)
    Call WKOut(rep, WK_ITEM_Total, "", ben.value(trans_MarketValueorig_db) - ben.value(trans_BenefitAlready_db), "'B'", True)
  End If
  
  Call WKOut(rep, WK_SECTION_HEADER_VALUE)
  
  Call WKOut(rep, WK_ITEM_TEXT, IIf(ben.value(trans_IsCAr_db), "Market value when transferred", "Greater of 'A' and 'B'"), ben.value(trans_value_db))
  
  
  Call OPRAWorkingPaperValue(ben, rep)
  
  Call WKOut(rep, WK_SECTION_HEADER_LESS)
  Call WKOut(rep, WK_ITEM_TEXT, "Amount paid by the employee or taxed under PAYE", FormatWN(ben.value(Trans_MadeGood_db), , True))
  Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)
  Call WKOut(rep, WK_ITEM_Total, "Taxable benefit", ben.value(trans_benefit), , True)
  
  
  
  
AssetsTransferred_PrintWKBody_End:
  
  
  Call xReturn("AssetsTransferred_PrintWKBody")
  Exit Function
  
AssetsTransferred_PrintWKBody_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AssetsTransferred_PrintWKBody", "Assets Transferred Print WK Body", "Error printing an assets transferred work sheet.")
  Resume AssetsTransferred_PrintWKBody_End

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
  ben.value(trans_ERROR) = ""
  ben.value(trans_value_db) = S_ERROR
  ben.value(trans_MadeGood_Net) = S_ERROR
  ben.value(trans_Benefit_Reportable) = False
  ben.value(trans_OPRA_Ammount_Foregone_Used_For_Value) = False
  
  ben.value(trans_benefit) = S_ERROR
End Sub

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_ASSETSTRANSFERRED_A, Item, RHS)
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
  IBenefitClass_Name = m_BenItems(trans_Item_db)
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
  Dim Trans As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim sIndex As String
  Dim i As Long
  Dim ben As IBenefitClass
  Dim sIRDesc As String
  
  On Error GoTo AssetsTransferred_ReadDB_err
  xSet "AssetsTransferred_ReadDB"
  If m_ReadFromDB Then GoTo AssetsTransferred_ReadDB_end
  Set ben = Me
  Set rs = m_Parent.Parent.rsBenTables(TBL_TRANS)
  If Len(m_sbookmark) = 0 Then
    s = m_Parent.PersonnelNumber
    rs.FindFirst ("P_Num = '" & s & "'")
    If Not rs.NoMatch Then
      i = i + 1
      m_sbookmark = rs.Bookmark
      rs.FindNext ("P_NUM = '" & s & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set Trans = New AssetsTransferred
        Trans.RSBookMark = rs.Bookmark
        Trans.CompanyDefined = m_CompanyDefined
        Set Trans.Parent = m_Parent
        Call m_Parent.benefits.Add(Trans)
        Set Trans = Nothing
        rs.FindNext ("P_NUM = '" & s & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    rs.Bookmark = m_sbookmark
    Call StandardReadData(ben, rs)
    ben.value(trans_EmployeeReference_db) = "" & rs.Fields(S_FIELD_PERSONEL_NUMBER).value
    ben.value(trans_MarketValueatTrans_db) = IIf(IsNull(rs.Fields("MktVal").value), 0, rs.Fields("MktVal").value)

    ben.value(trans_Item_db) = "" & rs.Fields("Item").value
    ben.value(trans_IsCAr_db) = IIf(IsNull(rs.Fields("IsCar").value), 0, rs.Fields("IsCar").value)
    ben.value(trans_MarketValueorig_db) = IIf(IsNull(rs.Fields("FirstMV").value), 0, rs.Fields("FirstMV").value)
    ben.value(trans_BenefitAlready_db) = IIf(IsNull(rs.Fields("SumBen").value), 0, rs.Fields("SumBen").value)
    ben.value(Trans_MadeGood_db) = IIf(IsNull(rs.Fields("MADEGOOD").value), 0, rs.Fields("MADEGOOD").value)
    ben.value(trans_value_db) = IIf(IsNull(rs.Fields("Value").value), 0, rs.Fields("Value").value)
    
    
    Call IRDescDB(ben, rs, True)
    m_ReadFromDB = True
    
    
  End If


AssetsTransferred_ReadDB_end:
  IBenefitClass_ReadDB = i
  xReturn "AssetsTransferred_ReadDB"
  Exit Function
AssetsTransferred_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "AssetsTransferred_ReadDB", "ERR_READ_Trans_BENEFITS", "Unable to read in the Services / Assets benefits"
   Resume AssetsTransferred_ReadDB_end
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
  bc = BC_ASSETSTRANSFERRED_A
  m_NeedToCalculate = True
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  With p11d32
    Call SetStandardBenItemsInformation(bc, Me)
    
    .BenDataLinkDataType(bc, trans_MarketValueorig_db) = TYPE_LONG
    .BenDataLinkDataType(bc, trans_MarketValueatTrans_db) = TYPE_LONG
    .BenDataLinkDataType(bc, trans_BenefitAlready_db) = TYPE_LONG
    .BenDataLinkDataType(bc, trans_EmployeeReference_db) = TYPE_STR
    .BenDataLinkDataType(bc, trans_IsCAr_db) = TYPE_BOOL
    
    
    
    .BenDataLinkUDMDescription(bc, trans_MarketValueorig_db) = "Original MV"
    .BenDataLinkUDMDescription(bc, trans_MarketValueatTrans_db) = "Transfer MV"
    .BenDataLinkUDMDescription(bc, trans_BenefitAlready_db) = "Benefit used"
    
    Call SetIRDesriptionInformation(bc)
  End With

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub

Private Function IBenefitClass_SetValue(ByVal Item As Long, value As Variant) As Boolean
  m_BenItems(Item) = value
  IBenefitClass_SetValue = True
End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo AssetsTransferred_DeleteDB_Err
  Call xSet("AssetsTransferred_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_TRANS)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
AssetsTransferred_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("AssetsTransferred_DeleteDB")
  Exit Function

AssetsTransferred_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AssetsTransferred_DeleteDB", "Assets Transferred Delete DB", "Error deleting an asset transferred.")
  Resume AssetsTransferred_DeleteDB_End
End Function


Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo AssetsTransferred_WriteDB_Err
  Call xSet("AssetsTransferred_WriteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_TRANS)
  Set ben = Me
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Edit
  Else
    rs.AddNew
    rs.Fields(S_FIELD_PERSONEL_NUMBER).value = ben.value(trans_EmployeeReference_db)
  End If
  Call StandardWriteData(ben, rs)
  rs.Fields("MKTVAL").value = ben.value(trans_MarketValueatTrans_db)
  rs.Fields("IsCar").value = ben.value(trans_IsCAr_db)
  rs.Fields("FirstMV").value = ben.value(trans_MarketValueorig_db)
  rs.Fields("SumBen").value = ben.value(trans_BenefitAlready_db)
  rs.Fields("MADEGOOD").value = ben.value(Trans_MadeGood_db)
  
  rs.Fields("ITEM").value = ben.value(trans_Item_db)
  rs.Fields("Value").value = ben.value(trans_value_db)
  
  Call IRDescDB(ben, rs, False)
    
  rs.Update
  If Len(m_sbookmark) = 0 Then
    m_sbookmark = rs.LastModified
  End If
  
  ben.Dirty = False
  IBenefitClass_WriteDB = True
AssetsTransferred_WriteDB_End:
  Set rs = Nothing
  Call xReturn("AssetsTransferred_WriteDB")
  Exit Function
AssetsTransferred_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "AssetsTransferred_WriteDB", "Assets Transferred WriteDB", "Error writing an assets transferred to the database.")
  Resume AssetsTransferred_WriteDB_End
  Resume
End Function


Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function
Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_TRANS
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property


