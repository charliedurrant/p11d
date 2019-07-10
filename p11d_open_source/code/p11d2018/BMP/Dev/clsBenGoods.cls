VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssetsAtDisposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements IBenefitClass

Public Enum AssetsAtDisposalItems
  AssetsAtDisposal_item_db = ITEM_DESC 'bf
  AssetsAtDisposal_Value = ITEM_VALUE
  AssetsAtDisposal_MadeGood_Net = ITEM_MADEGOOD_NET
  AssetsAtDisposal_Benefit = ITEM_BENEFIT
  AssetsAtDisposal_MadeGood_db = ITEM_MADEGOOD
  AssetsAtDisposal_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  AssetsAtDisposal_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  AssetsAtDisposal_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  AssetsAtDisposal_BoxNumber = ITEM_BOX_NUMBER
  AssetsAtDisposal_MadeGoodIsTaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  AssetsAtDisposal_Class1AAdjustment = ITEM_CLASS1A_ADJUSTMENT
  AssetsAtDisposal_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  AssetsAtDisposal_NIC_Class1A_Value = ITEM_NIC_CLASS1A_VALUE
  AssetsAtDisposal_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  
  AssetsAtDisposal_IRDesc_db
  
  AssetsAtDisposal_MarketValue_db 'bf
  AssetsAtDisposal_availablefrom_db 'bf - Reset
  AssetsAtDisposal_availableto_db 'bf test - reset
  AssetsAtDisposal_DateAvail_db 'bf
  
  
  AssetsAtDisposal_Marginal_db
  AssetsAtDisposal_Rent_db
  AssetsAtDisposal_BasicValue
  AssetsAtDisposal_BasicValueAndMarginal
  AssetsAtDisposal_ComputerRelated_db
  
  AssetsAtDisposal_TotalUnavailable
  AssetsAtDisposal_RelevantPercentage
  AssetsAtDisposal_LASTITEM = AssetsAtDisposal_RelevantPercentage
    
End Enum

Private m_Parent As Object
Private m_BenClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_CompanyDefined  As Boolean
Private m_BenItems(1 To AssetsAtDisposal_LASTITEM)
Private m_dirty As Boolean

Private m_InvalidFields As Long
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Assets at disposal - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim i As Long
  Dim d As Double
  Dim ben As IBenefitClass
  
  On Error GoTo AssetsAtDisposal_Calculate_err
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  
  If ben.value(AssetsAtDisposal_DateAvail_db) < #6/4/1980# Then
    ben.value(AssetsAtDisposal_RelevantPercentage) = p11d32.Rates.value(goodLOWPERCENTAGE)
  Else
    ben.value(AssetsAtDisposal_RelevantPercentage) = p11d32.Rates.value(goodHIGHPERCENTAGE)
  End If
  
  ben.value(AssetsAtDisposal_BasicValue) = Max(ben.value(AssetsAtDisposal_Rent_db), (ben.value(AssetsAtDisposal_MarketValue_db) * ben.value(AssetsAtDisposal_RelevantPercentage)))
  ben.value(AssetsAtDisposal_BasicValueAndMarginal) = ben.value(AssetsAtDisposal_BasicValue) + ben.value(AssetsAtDisposal_Marginal_db)
  
  d = dGetDateFactor(i, ben.value(AssetsAtDisposal_availablefrom_db), ben.value(AssetsAtDisposal_availableto_db))
  ben.value(AssetsAtDisposal_TotalUnavailable) = i
  ben.value(AssetsAtDisposal_Value) = d * ben.value(AssetsAtDisposal_BasicValueAndMarginal)
  
  
  ben.value(AssetsAtDisposal_MadeGood_Net) = Min(ben.value(AssetsAtDisposal_Value), ben.value(AssetsAtDisposal_MadeGood_db))
  
  ben.value(AssetsAtDisposal_Benefit) = ben.value(AssetsAtDisposal_Value) - ben.value(AssetsAtDisposal_MadeGood_Net)
  
  If ben.value(AssetsAtDisposal_ComputerRelated_db) Then
    ben.value(AssetsAtDisposal_Benefit) = ben.value(AssetsAtDisposal_Benefit) - 500
    ben.value(AssetsAtDisposal_Benefit) = Max(ben.value(AssetsAtDisposal_Benefit), 0)
    ben.value(AssetsAtDisposal_Value) = ben.value(AssetsAtDisposal_Benefit) + ben.value(AssetsAtDisposal_MadeGood_Net)
  End If
  ben.value(AssetsAtDisposal_Benefit_Reportable) = ben.value(AssetsAtDisposal_Value) > 0
  
  Call BenCalcNIC(ben)
  IBenefitClass_CalculateBody = ben.value(AssetsAtDisposal_Benefit)
  
AssetsAtDisposal_Calculate_end:
  Exit Function
AssetsAtDisposal_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume AssetsAtDisposal_Calculate_end
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
  IBenefitClass_CanBringForward = StandardCanBringForward(Me, AssetsAtDisposal_availableto_db)
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim ass As AssetsAtDisposal
  
  Set ass = New AssetsAtDisposal
  Set IBenefitClass_Copy = CopyBenStandard(Parent, ass, Me)
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_ASSETSATDISPOSAL_L
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

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property
Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(AssetsAtDisposal_item_db)
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

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo AssetsAtDisposal_PrintWKBody_Err
  Call xSet("AssetsAtDisposal_PrintWKBody")
  
  Set ben = Me
  
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  Call WKOut(rep, WK_ITEM_TEXT, "Higher of rent paid and relevant percentage of market value", ben.value(AssetsAtDisposal_BasicValue), "A", True)
  Call WKOut(rep, WK_ITEM_TEXT, "Costs: ", ben.value(AssetsAtDisposal_Marginal_db), "B", True)
  Call WKOut(rep, WK_ITEM_TEXT, "Available from: " & S_WK_NORMAL_BOLD_FONT & ben.value(AssetsAtDisposal_availablefrom_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Available to: " & S_WK_NORMAL_BOLD_FONT & ben.value(AssetsAtDisposal_availableto_db))
  
  If ben.value(AssetsAtDisposal_ComputerRelated_db) Then
    Call WKOut(rep, WK_BLANK_LINE)
    Call WKOut(rep, WK_ITEM_TEXT_BOLD, "The asset is a computer related asset under s.320 ITEPA 2003 - only the cash equivalent value over £500 is taxable.")
  End If
  
  
  Call WKOut(rep, WK_SECTION_HEADER_VALUE)
  Call WKOut(rep, WK_ITEM_TEXT, "(A + B) x (" & p11d32.Rates.value(DaysInYear) - ben.value(AssetsAtDisposal_TotalUnavailable) & " / " & p11d32.Rates.value(DaysInYear) & ")", ben.value(ITEM_VALUE))
  Call WKOut(rep, WK_SECTION_HEADER_LESS)
  Call WKOut(rep, WK_ITEM_TEXT, "Made good", ben.value(ITEM_MADEGOOD_NET), , , True)
  Call WKOut(rep, WK_SECTION_HEADER_BENEFIT)
  Call WKOut(rep, WK_ITEM_Total, "Total benefit", ben.value(ITEM_BENEFIT), , True)
  
    
AssetsAtDisposal_PrintWKBody_End:
  Call xReturn("AssetsAtDisposal_PrintWKBody")
  Exit Function

AssetsAtDisposal_PrintWKBody_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AssetsAtDisposal_PrintWKBody", "Assets At Disposal Print WK Body", "Error printing the AssetsAtDisposal working paper")
  Resume AssetsAtDisposal_PrintWKBody_End
  Resume

End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim AssetsAtDisposal As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim sIndex As String
  Dim i As Long
  Dim ben As IBenefitClass
  
  On Error GoTo AssetsAtDisposal_ReadDB_err
  
  Call xSet("AssetsAtDisposal_ReadDB")
  
  If m_ReadFromDB Then GoTo AssetsAtDisposal_ReadDB_end
  Set ben = Me
  
  Set rs = m_Parent.Parent.rsBenTables(ben.TABLE)
  If Len(m_sbookmark) = 0 Then
    s = m_Parent.PersonnelNumber
    rs.FindFirst ("P_NUM = '" & s & "'")
    If Not rs.NoMatch Then
      i = i + 1
      m_sbookmark = rs.Bookmark
      rs.FindNext ("P_NUM = '" & s & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set AssetsAtDisposal = New AssetsAtDisposal
        AssetsAtDisposal.RSBookMark = rs.Bookmark
        AssetsAtDisposal.CompanyDefined = m_CompanyDefined
        Set AssetsAtDisposal.Parent = m_Parent
        Call m_Parent.benefits.Add(AssetsAtDisposal)
        Set AssetsAtDisposal = Nothing
        rs.FindNext ("P_NUM = '" & s & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    rs.Bookmark = m_sbookmark
    Call StandardReadData(ben, rs)
    
    ben.value(AssetsAtDisposal_item_db) = "" & rs.Fields("Item").value
    ben.value(AssetsAtDisposal_availablefrom_db) = IIf(IsNull(rs.Fields("AvailFrom").value), p11d32.Rates.value(TaxYearStart), rs.Fields("AvailFrom").value)
    ben.value(AssetsAtDisposal_DateAvail_db) = IIf(IsNull(rs.Fields("DATEAVAIL").value), p11d32.Rates.value(goodDEFAULTDATE), rs.Fields("DATEAVAIL").value)
    ben.value(AssetsAtDisposal_availableto_db) = IIf(IsNull(rs.Fields("AvailTo").value), p11d32.Rates.value(TaxYearEnd), rs.Fields("AvailTo").value)
    ben.value(AssetsAtDisposal_MarketValue_db) = IIf(IsNull(rs.Fields("MKTVAL").value), 0, rs.Fields("MKTVAL").value)
    ben.value(AssetsAtDisposal_Marginal_db) = IIf(IsNull(rs.Fields("MARGINAL").value), 0, rs.Fields("MARGINAL").value)
    ben.value(AssetsAtDisposal_Rent_db) = rs.Fields("RENT").value
'MP DB - not in calc or display
'    ben.value(AssetsAtDisposal_BusinessUse_db) = IIf(IsNull(rs.Fields("BUSINESS").value), 0, rs.Fields("BUSINESS").value)
    ben.value(AssetsAtDisposal_MadeGood_db) = IIf(IsNull(rs.Fields("MADEGOOD").value), 0, rs.Fields("MADEGOOD").value)
    ben.value(AssetsAtDisposal_ComputerRelated_db) = IIf(IsNull(rs.Fields("ComputerRelated").value), False, rs.Fields("ComputerRelated").value)

    
    Call IRDescDB(ben, rs, True)
    m_ReadFromDB = True
  End If
  
AssetsAtDisposal_ReadDB_end:
  IBenefitClass_ReadDB = i
  xReturn "AssetsAtDisposal_ReadDB"
  Exit Function
  
AssetsAtDisposal_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "AssetsAtDisposal_ReadDB", "ERR_READ_GOODS_BENEFITS", "Unable to read in the Services / Assets benefits"
   Resume AssetsAtDisposal_ReadDB_end
   
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
  bc = BC_ASSETSATDISPOSAL_L
  m_NeedToCalculate = True
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  With p11d32
    Call SetStandardBenItemsInformation(bc, Me)
    
    .BenDataLinkDataType(bc, AssetsAtDisposal_BasicValue) = TYPE_LONG
    .BenDataLinkDataType(bc, AssetsAtDisposal_MarketValue_db) = TYPE_LONG
    .BenDataLinkDataType(bc, AssetsAtDisposal_Marginal_db) = TYPE_LONG
    .BenDataLinkDataType(bc, AssetsAtDisposal_Rent_db) = TYPE_LONG
'MP DB    .BenDataLinkDataType(bc, AssetsAtDisposal_BusinessUse_db) = TYPE_LONG
    .BenDataLinkDataType(bc, AssetsAtDisposal_BasicValueAndMarginal) = TYPE_LONG
    .BenDataLinkDataType(bc, AssetsAtDisposal_TotalUnavailable) = TYPE_LONG
    .BenDataLinkDataType(bc, AssetsAtDisposal_RelevantPercentage) = TYPE_DOUBLE
    

    
    .BenDataLinkDataType(bc, AssetsAtDisposal_availablefrom_db) = TYPE_DATE
    .BenDataLinkDataType(bc, AssetsAtDisposal_availableto_db) = TYPE_DATE
    .BenDataLinkDataType(bc, AssetsAtDisposal_DateAvail_db) = TYPE_DATE
       
    .BenDataLinkUDMDescription(bc, AssetsAtDisposal_MarketValue_db) = "Market value"
    .BenDataLinkUDMDescription(bc, AssetsAtDisposal_Rent_db) = "Rent"
    
    SetIRDesriptionInformation (bc)
    
  End With
  
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information.")
  Resume SetBenItemsInformation_end
End Sub

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  Call SetCalcDefaultsStandard(ben)
  
  ben.value(AssetsAtDisposal_BasicValue) = S_ERROR
  ben.value(AssetsAtDisposal_BasicValueAndMarginal) = S_ERROR
  
  ben.value(AssetsAtDisposal_RelevantPercentage) = S_ERROR
  ben.value(AssetsAtDisposal_Value) = S_ERROR
  ben.value(AssetsAtDisposal_TotalUnavailable) = S_ERROR
  
End Sub

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo AssetsAtDisposal_DeleteDB_Err
  Call xSet("AssetsAtDisposal_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_ASSETSATDISPOSAL)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
AssetsAtDisposal_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("AssetsAtDisposal_DeleteDB")
  Exit Function

AssetsAtDisposal_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AssetsAtDisposal_DeleteDB", "Assets At Disposal DeleteDB", "Error deleting 'assets at disposal' from the database.")
  Resume AssetsAtDisposal_DeleteDB_End
End Function


Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_ASSETSATDISPOSAL_L, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo AssetsAtDisposal_WriteDB_Err
  Call xSet("AssetsAtDisposal_WriteDB")
  Set ben = Me
  
  Call BenefitAddNewRecord(ben, rs)
  
  rs.Fields("MKTVAL").value = ben.value(AssetsAtDisposal_MarketValue_db)

  rs.Fields("DATEAVAIL").value = ben.value(AssetsAtDisposal_DateAvail_db)
  rs.Fields("ITEM").value = ben.value(AssetsAtDisposal_item_db)
  rs.Fields("UDBCODE").value = S_ASSETSATDISPOSAL_UDBCODE
  rs.Fields("ComputerRelated").value = ben.value(AssetsAtDisposal_ComputerRelated_db)
  
  Call BringForwardDatesWrite(ben, AssetsAtDisposal_availablefrom_db, AssetsAtDisposal_availableto_db, rs, "AVAILFROM", "AVAILTO")
  
  Call StandardWriteData(ben, rs)
  If Not p11d32.BringForward.Yes Then
    rs.Fields("MARGINAL").value = ben.value(AssetsAtDisposal_Marginal_db)
    rs.Fields("RENT").value = ben.value(AssetsAtDisposal_Rent_db)
    rs.Fields("MADEGOOD").value = ben.value(AssetsAtDisposal_MadeGood_db)
  End If
  
  Call IRDescDB(ben, rs, False)
  rs.Update
  
  If Len(m_sbookmark) = 0 Then m_sbookmark = rs.LastModified
  
  ben.Dirty = False
  
  IBenefitClass_WriteDB = True
  
AssetsAtDisposal_WriteDB_End:
  Set rs = Nothing
  Call xReturn("AssetsAtDisposal_WriteDB")
  Exit Function
AssetsAtDisposal_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "AssetsAtDisposal_WriteDB", "Assets At Disposal WriteDB", "Error writing 'assets at disposal' to the database.")
  
  Resume AssetsAtDisposal_WriteDB_End
  Resume
End Function
Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_ASSETSATDISPOSAL
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property




