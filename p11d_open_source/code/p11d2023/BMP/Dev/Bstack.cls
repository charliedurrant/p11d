VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Relocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum RelocItems
  reloc_Item_db = ITEM_DESC
  reloc_Value = ITEM_VALUE
  reloc_MadeGood_Net = ITEM_MADEGOOD_NET
  reloc_Benefit = ITEM_BENEFIT
  reloc_MadeGood = ITEM_MADEGOOD
  reloc_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  reloc_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  reloc_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  reloc_BoxNumber = ITEM_BOX_NUMBER
  reloc_MadeGood_Is_TaxDeducted = ITEM_MADEGOOD_IS_TAXDEDUCTED
  reloc_Class1A_Adjustment = ITEM_CLASS1A_ADJUSTMENT
  reloc_NICClass1AAble = ITEM_NIC_CLASS1A_ABLE
  reloc_NIC_Class1A_Value = ITEM_NIC_CLASS1A_VALUE
  reloc_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  
  reloc_NQitem
  reloc_NQ_SPECIFIC_START = reloc_NQitem
  reloc_NQValue
  reloc_NQMadeGood_NET
  reloc_NQBenefit
  reloc_NQMadeGood
  reloc_NQActualAmountMadeGood
  reloc_NQBenefit_Reportable
  reloc_NQUDM_BENEFIT_TITLE
  reloc_NQBoxNumber
  reloc_NQMadeGood_Is_TaxDeducted
  reloc_NQClass1A_Adjustment
  reloc_NQNICClass1AAble
  reloc_NQNIC_Class1A_Value
  reloc_NQNIC_Class1A_Benefit
  reloc_NQ_SPECIFIC_END = reloc_NQNIC_Class1A_Benefit
  
  reloc_TaxYear_db 'bf
  reloc_UsedThisYear_db 'bf reloc_UsedThisYear = reloc_UsedLastyear
  reloc_UsedLastyear_db

  reloc_QualifyTotal
  reloc_Key_db
  reloc_RemainingReliefFromLastYear
  reloc_RemainingRelief
  reloc_LASTITEM = reloc_RemainingRelief
End Enum

Public RelocDetails As ObjectList
Private m_MaxKey As Long
Private m_BenClass As BEN_CLASS
Private m_ReadFromDB As Boolean
Private m_sbookmark As String
Private m_dirty As Boolean


Private m_Parent As Employee
Private m_InvalidFields As Long
Private m_BenItems(1 To reloc_LASTITEM)

Public NonQualifyingRelocation As NonQualifyingRelocation
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Relocation - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
  Dim i As Long
  Dim detail As RelocationDetail
  Dim bValidRelocationDetail As Boolean
  Dim ben As IBenefitClass
  Dim benNQReloc As IBenefitClass
  On Error GoTo clsBenReloc_Calculate_Err
  Call xSet("clsBenReloc_Calculate")
  
  Set ben = Me
  
  Call ben.SetCalcDefaults
  ben.value(reloc_UsedLastyear_db) = ben.value(reloc_UsedLastyear_db)
  ben.value(reloc_NQitem) = ben.value(reloc_Item_db)
  
  If RelocDetails.Count Then
    For i = 1 To RelocDetails.Count
      Set detail = RelocDetails(i)
      If Not detail Is Nothing Then
        bValidRelocationDetail = True
        Call detail.Calculate
        If detail.Qualify Then
          ben.value(reloc_Value) = ben.value(reloc_Value) + detail.value
          ben.value(reloc_MadeGood) = ben.value(reloc_MadeGood) + detail.MadeGood
          ben.value(reloc_MadeGood_Net) = ben.value(reloc_MadeGood_Net) + Min(detail.MadeGood, detail.value)
          If detail.IsTaxDeducted Then
            ben.value(reloc_MadeGood_Is_TaxDeducted) = True
            ben.value(reloc_Class1A_Adjustment) = ben.value(reloc_Class1A_Adjustment) + ben.value(reloc_MadeGood)
          End If
        Else
          ben.value(reloc_NQValue) = ben.value(reloc_NQValue) + detail.value
          ben.value(reloc_NQMadeGood) = ben.value(reloc_NQMadeGood) + detail.MadeGood
          ben.value(reloc_NQMadeGood_NET) = ben.value(reloc_NQMadeGood_NET) + Min(detail.MadeGood, detail.value)
        End If
      End If
    Next i
    If Not bValidRelocationDetail Then GoTo NO_BENEFIT
    ben.value(reloc_RemainingReliefFromLastYear) = Max(0, L_RELOCEXEMPT - ben.value(reloc_UsedLastyear_db))
    ben.value(reloc_QualifyTotal) = ben.value(reloc_Value) - ben.value(reloc_MadeGood_Net)
    
    If ben.value(reloc_QualifyTotal) > 0 Then
      ben.value(reloc_Benefit) = ben.value(reloc_QualifyTotal) - ben.value(reloc_RemainingReliefFromLastYear)
      ben.value(reloc_RemainingRelief) = Max(-1 * (ben.value(reloc_Benefit)), 0)
      ben.value(reloc_Benefit) = Max(ben.value(reloc_Benefit), 0)
      ben.value(reloc_UsedThisYear_db) = ben.value(reloc_RemainingReliefFromLastYear) - ben.value(reloc_RemainingRelief)
    Else
      ben.value(reloc_Benefit) = 0
      ben.value(reloc_RemainingRelief) = ben.value(reloc_RemainingReliefFromLastYear)
      ben.value(reloc_UsedThisYear_db) = 0
    End If
    
    ben.value(reloc_NQBenefit) = ben.value(reloc_NQValue) - ben.value(reloc_NQMadeGood_NET)
    If ben.value(reloc_NQValue) > 0 Then ben.value(reloc_NQBenefit_Reportable) = True
    If ben.value(reloc_Benefit) > 0 Then ben.value(reloc_Benefit_Reportable) = True
  Else
NO_BENEFIT:
    ben.value(reloc_NQBenefit) = 0
    ben.value(reloc_Benefit) = 0
  End If
  
  If m_BenClass = BC_NON_QUALIFYING_RELOCATION_N Then
    'switch the benefits
    ben.value(reloc_Value) = ben.value(reloc_NQValue)
    ben.value(reloc_Benefit) = ben.value(reloc_NQBenefit)
    ben.value(reloc_MadeGood) = ben.value(reloc_NQMadeGood)
    ben.value(reloc_MadeGood_Net) = ben.value(reloc_NQMadeGood_NET)
    ben.value(reloc_Benefit_Reportable) = ben.value(reloc_NQBenefit_Reportable)
  End If
  Call BenCalcNIC(ben)
  Set benNQReloc = NonQualifyingRelocation
  If (Not benNQReloc Is Nothing) Then
    Call BenCalcNIC(benNQReloc)
  End If
  
  
  
  IBenefitClass_CalculateBody = ben.value(reloc_Benefit)
  
clsBenReloc_Calculate_End:
  Call xReturn("clsBenReloc_Calculate")
  Exit Function
clsBenReloc_Calculate_Err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume clsBenReloc_Calculate_End
  Resume

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Accommodation"
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
  Dim ben As IBenefitClass
  Set ben = Me
  Call ben.Calculate
  IBenefitClass_CanBringForward = (StrComp(ben.value(reloc_TaxYear_db), p11d32.Rates.value(RelocationLastYear), vbTextCompare) = 0)
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim reloc As Relocation
  
  On Error GoTo Copy_END
  Call xSet("Copy")
  
  
  Set reloc = New Relocation
  Set IBenefitClass_Copy = CopyBenStandard(Parent, reloc, Me)
  
  reloc.AddNonQualifyingRelocation
  
Copy_END:
  Call xReturn("Copy")
  Exit Function
Copy_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Copy", "Copy", "Error copying a relocation benefit.")
  Resume Copy_ERR
End Function

Private Sub Class_Initialize()
  Set RelocDetails = New ObjectList
  Call IBenefitClass_SetBenItemsInformation
End Sub
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  'do not set IBenefitClass_BenefitClass to m_benclass see CAD for explination
  IBenefitClass_BenefitClass = BC_QUALIFYING_RELOCATION_J
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

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me)
End Function
Private Sub PrintTable(rep As Reporter, bQualifying As Boolean)
  Dim i As Long
  Dim RD As RelocationDetail
  Dim bFirst As Boolean
  Dim ben As IBenefitClass
   
  On Error GoTo PrintTable_ERR
  
  Call xSet("PrintTable")
  
  Set ben = Me
   
  Call WKTblOtherTypeTable(rep, "~~~Description", "~Cost to~you", "Amount made~good or taxed~under PAYE", "~Cash~equivalent")
   
  For i = 1 To RelocDetails.Count
    Set RD = RelocDetails(i)
    If Not RD Is Nothing Then
      If RD.Qualify = bQualifying Then
        If Not bFirst Then
          Call WKTableRow(rep, RD.Item, FormatWN(RD.value), FormatWN(RD.MadeGood, , True), FormatWN(RD.benefit))
          bFirst = True
        Else
          Call WKTableRow(rep, RD.Item, FormatWN(RD.value), FormatWN(RD.MadeGood, "", True), FormatWN(RD.benefit, ""))
        End If
      End If
    End If
  Next
   
  If bFirst Then
    If bQualifying Then
      Call WKTableTotals(rep, "", FormatWN(ben.value(reloc_Value)), FormatWN(ben.value(reloc_MadeGood), , True), FormatWN(ben.value(reloc_QualifyTotal)))
    Else
      Call WKTableTotals(rep, "", FormatWN(ben.value(reloc_NQValue)), FormatWN(ben.value(reloc_NQMadeGood), , True), FormatWN(ben.value(reloc_NQBenefit)))
    End If
  Else
    Call WKOut(rep, WK_ITEM_TEXT, "     'NO DETAILS'")
  End If
  
  Call WKOut(rep, WK_BLANK_LINE)
  
PrintTable_END:
  Call xReturn("PrintTable")
  Exit Sub
PrintTable_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PrintTable", "PrintTable", "Error printing table for relocation benefit.")
  Resume PrintTable_END
End Sub
Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass

  On Error GoTo Reloc_PrintWK_Err
  Call xSet("Reloc_PRINTWK")
  
  Set ben = Me
  Call WKOut(rep, WK_SECTION_HEADER_DETAILS)
  Call WKOut(rep, WK_ITEM_DESCRIPTION, "Address: ", ben.value(ITEM_DESC))
  Call WKOut(rep, WK_ITEM_TEXT, "Tax year of move: " & S_WK_NORMAL_BOLD_FONT & ben.value(reloc_TaxYear_db))
  'Qualifying
  Call WKOut(rep, WK_SECTION_HEADER, "Qualifying relocation")
  Call WKOut(rep, WK_ITEM_TEXT, "Qualifying relocation items:")
  'change to print tabel data
  Call PrintTable(rep, True)
  Call WKOut(rep, WK_ITEM_TEXT, "Add: Amount of " & FormatWN(L_RELOCEXEMPT) & " exemption used in prior year.", ben.value(reloc_UsedLastyear_db))
  Call WKOut(rep, WK_ITEM_TEXT, "Less: " & FormatWN(L_RELOCEXEMPT) & " exemption.", L_RELOCEXEMPT, , , True)
  Call WKOut(rep, WK_ITEM_Total, "Qualifying benefit: ", ben.value(ITEM_BENEFIT), , True)
  
  'non qualifying
  Call WKOut(rep, WK_SECTION_HEADER, "Non Qualifying relocation")
  Call WKOut(rep, WK_ITEM_TEXT, "Non Qualifying relocation items:")
  Call PrintTable(rep, False)
  Call WKOut(rep, WK_ITEM_Total, "Non Qualifying benefit: ", ben.value(reloc_NQBenefit), , True)
  

Reloc_PrintWK_End:
  Call xReturn("Reloc_PrintWK")
  Exit Function

Reloc_PrintWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Reloc_PrintWK", "ERR_Reloc_PrintWK", "Error printing the Relocation working paper")
  Resume Reloc_PrintWK_End
  Resume

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  
  m_BenItems(reloc_MadeGood) = 0
  m_BenItems(reloc_NQMadeGood) = 0
  
  m_BenItems(reloc_Value) = 0
  m_BenItems(reloc_NQValue) = 0
  
  m_BenItems(reloc_Benefit) = S_ERROR
  m_BenItems(reloc_NQBenefit) = S_ERROR
  m_BenItems(reloc_QualifyTotal) = 0
  
  m_BenItems(reloc_NQMadeGood_NET) = 0
  m_BenItems(reloc_MadeGood_Net) = 0
  
  m_BenItems(reloc_RemainingRelief) = 0
  m_BenItems(reloc_UsedThisYear_db) = 0
  
  m_BenItems(reloc_Class1A_Adjustment) = 0
  
  m_BenItems(reloc_Benefit_Reportable) = False
  m_BenItems(reloc_NQBenefit_Reportable) = False
End Sub
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
  IBenefitClass_Name = m_BenItems(reloc_Item_db)
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
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
  Dim NQR As NonQualifyingRelocation
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
  
  m_NeedToCalculate = True
  
  bc = BC_QUALIFYING_RELOCATION_J
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  With p11d32
    Set ben = Me
    .BenDataLinkBenfitTable(bc) = ben.TABLE
    Call SetStandardBenItemsDataTypes(bc)
    Call SetStandardBenItemsUDMData(bc)
    Call SetStandardBenItemsMMFieldSize(bc)
  
    .BenDataLinkDataType(bc, reloc_NQitem) = TYPE_STR
    .BenDataLinkDataType(bc, reloc_NQValue) = TYPE_LONG
    .BenDataLinkDataType(bc, reloc_NQMadeGood) = TYPE_LONG
    .BenDataLinkDataType(bc, reloc_NQBenefit) = TYPE_LONG
    .BenDataLinkDataType(bc, reloc_NQMadeGood_NET) = TYPE_LONG
    .BenDataLinkDataType(bc, reloc_NQBenefit_Reportable) = TYPE_BOOL
    
    .BenDataLinkDataType(bc, reloc_TaxYear_db) = TYPE_STR
    .BenDataLinkDataType(bc, reloc_UsedLastyear_db) = TYPE_LONG
    .BenDataLinkDataType(bc, reloc_UsedThisYear_db) = TYPE_LONG
    
    .BenDataLinkDataType(bc, reloc_QualifyTotal) = TYPE_LONG
    
    .BenDataLinkDataType(bc, reloc_Key_db) = TYPE_LONG
    
    .BenDataLinkDataType(bc, reloc_RemainingReliefFromLastYear) = TYPE_LONG
    .BenDataLinkDataType(bc, reloc_RemainingRelief) = TYPE_LONG
  
    .BenDataLinkUDMDescription(bc, reloc_TaxYear_db) = "Tax year"
    .BenDataLinkUDMDescription(bc, reloc_UsedLastyear_db) = "Used last year"
    .BenDataLinkUDMDescription(bc, reloc_NQValue) = "NQ Value"
    
    'MMField sizes
    .DataLinkInitialised(bc) = True
    'init data links

    Set NQR = New NonQualifyingRelocation
    Set NQR = Nothing
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
End Function
Private Sub IBenefitClass_Kill()
  Set NonQualifyingRelocation = Nothing
  Set m_Parent = Nothing
End Sub
Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Public Sub AddNonQualifyingRelocation()
  Static bCreatedNonQualifyingReloc As Boolean
  Dim ben As IBenefitClass
  On Error GoTo AddNonQualifyingRelocation_ERR
  
  Call xSet("AddNonQualifyingRelocation")
  
  If Not bCreatedNonQualifyingReloc Then
    bCreatedNonQualifyingReloc = True
    Set NonQualifyingRelocation = New NonQualifyingRelocation
    Set NonQualifyingRelocation.RelocationBen = Me
    Set ben = NonQualifyingRelocation
    ben.CompanyDefined = True
    Call m_Parent.benefits.Add(NonQualifyingRelocation)
  End If

AddNonQualifyingRelocation_END:
  Call xReturn("AddNonQualifyingRelocation")
  Exit Sub
AddNonQualifyingRelocation_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddNonQualifyingRelocation", "Add Non Qualifying Relocation", "Error adding a non qualifying relocation benefit with a qualifying relocation benefit.")
  Resume AddNonQualifyingRelocation_END
End Sub
Private Function IBenefitClass_ReadDB() As Long
  Dim rs As Recordset
  Dim reloc As Relocation
  Dim relocben As IBenefitClass
  Dim s As String
  Dim i As Long
  Dim ben As IBenefitClass

  On Error GoTo clsBenReloc_readDB_Err
  Call xSet("clsBenReloc_readDB")

  If m_ReadFromDB Then GoTo clsBenReloc_readDB_End
  Set rs = m_Parent.Parent.rsBenTables(TBL_RELOC)
  Set ben = Me
  
  If Len(m_sbookmark) = 0 Then
    rs.FindFirst ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
    If Not rs.NoMatch Then
      Call AddNonQualifyingRelocation
      m_sbookmark = rs.Bookmark
      i = i + 1
      rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set reloc = New Relocation
        Set relocben = reloc
        relocben.RSBookMark = rs.Bookmark
        Set relocben.Parent = m_Parent
        Call reloc.AddNonQualifyingRelocation
        Call m_Parent.benefits.Add(reloc)
        Set reloc = Nothing
        Set relocben = Nothing
        rs.FindNext ("P_NUM = '" & m_Parent.PersonnelNumber & "'")
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
    
  If Len(m_sbookmark) > 0 Then
    Call StandardReadData(ben)
    rs.Bookmark = m_sbookmark
    Call StandardReadData(NonQualifyingRelocation)
    ben.value(reloc_Item_db) = "" & rs.Fields("Item").value
    ben.value(reloc_TaxYear_db) = IIf(IsNull(rs.Fields("TaxYear").value), p11d32.Rates.value(RelocationThisYear), rs.Fields("TaxYear").value)
    ben.value(reloc_UsedLastyear_db) = IIf(IsNull(rs.Fields("LastYear").value), 0, rs.Fields("LastYear").value)
    ben.value(reloc_Key_db) = IIf(IsNull(rs.Fields("RelocKey").value), 0, rs.Fields("RelocKey").value)
    Call ReadDetails
    m_ReadFromDB = True
  End If
  
clsBenReloc_readDB_End:
  IBenefitClass_ReadDB = i
  Set rs = Nothing
  Call xReturn("clsBenReloc_readDB")
  Exit Function
clsBenReloc_readDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenReloc_readDB", "ERR_RELOCATION_READDB", "Unable to read the relocation details.")
  Resume clsBenReloc_readDB_End
  Resume
End Function
Public Property Get IBenefitClass_PrintHeader() As String
  IBenefitClass_PrintHeader = "{x=0}{Arial=10,n}Relocation benefit" & vbCrLf
End Property

Private Function ReadDetails() As Long
  Dim detail As RelocationDetail
  Dim rs As Recordset
   
  On Error GoTo ReadDetails_Err
  Call xSet("ReadDetails")
  Set rs = m_Parent.Parent.db.OpenRecordset(sql.Queries(SELECT_RELOCDETAILS, m_BenItems(reloc_Key_db)), dbOpenForwardOnly)
  Do While Not rs.EOF
    Set detail = New RelocationDetail
    detail.Item = "" & rs.Fields("Item").value
    detail.value = rs.Fields("Value").value
    detail.MadeGood = rs.Fields("MadeGood").value
    detail.Comments = "" & rs.Fields("Comments").value
    detail.Qualify = rs.Fields("Qualify").value
    detail.Key = RelocDetails.Add(detail)
    
    If p11d32.BringForward.Yes And p11d32.AppYear = 2000 Then
      detail.IsTaxDeducted = False
    Else
      detail.IsTaxDeducted = rs.Fields("MadeGoodIsTaxDeducted").value
    End If
    Set detail = Nothing
    rs.MoveNext
  Loop

ReadDetails_End:
  Call xReturn("ReadDetails")
  Exit Function
  
ReadDetails_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadDetails", "ERR_READ_DETAILS", "Unable to read the relocation details from the database.")
  Resume ReadDetails_End
End Function

Private Function WriteDetails() As Long
  Dim detail As RelocationDetail
  Dim rs As Recordset
  Dim benEmployee As IBenefitClass
  
  Dim i As Long
  
  On Error GoTo WriteDetails_Err
  Call xSet("WriteDetails")
  
  If p11d32.BringForward.Yes Then GoTo WriteDetails_End
  
  If Not (m_BenItems(reloc_Key_db) = Empty) Then Call m_Parent.Parent.db.Execute(sql.Queries(DELETE_RELOCDETAILS, m_BenItems(reloc_Key_db)))
  
  Set benEmployee = m_Parent
  Set rs = m_Parent.Parent.db.OpenRecordset(sql.Queries(SELECT_RELOCDETAILS, m_BenItems(reloc_Key_db)), dbOpenDynaset)
  
  For i = 1 To RelocDetails.Count
    Set detail = RelocDetails(i)
    If Not detail Is Nothing Then
      rs.AddNew
      rs.Fields("Item").value = detail.Item
      rs.Fields("Value").value = detail.value
      rs.Fields("MadeGood").value = detail.MadeGood
      rs.Fields("Comments").value = detail.Comments
      rs.Fields("Qualify").value = IIf(IsNull(detail.Qualify), False, detail.Qualify)
      rs.Fields("MadeGoodIsTaxDeducted").value = IIf(IsNull(detail.IsTaxDeducted), False, detail.IsTaxDeducted)
      rs.Fields("RelocKey").value = m_BenItems(reloc_Key_db)
      rs.Fields("Address").value = m_BenItems(reloc_Item_db)
      rs.Fields("P_Num").value = benEmployee.value(ee_PersonnelNumber_db)
      rs.Update
    End If
  Next i

WriteDetails_End:
  Set rs = Nothing
  Call xReturn("WriteDetails")
  Exit Function
  
WriteDetails_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteDetails", "ERR_READ_DETAILS", "Unable to read the relocation details from the database.")
  Resume WriteDetails_End
  Resume
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_RELOC
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = CorrectBenValue(BC_QUALIFYING_RELOCATION_J, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo clsBenReloc_WriteDB_Err
  Call xSet("clsBenReloc_WriteDB")

  Set ben = Me
   
  Call BenefitAddNewRecord(ben, rs)
   
  rs.Fields("Item").value = "" & ben.value(reloc_Item_db)
  rs.Fields("TaxYear").value = IIf(IsNull(ben.value(reloc_TaxYear_db)), p11d32.Rates.value(RelocationThisYear), ben.value(reloc_TaxYear_db))

  If p11d32.BringForward.Yes Then
    Call ben.Calculate
    rs.Fields("LastYear").value = ben.value(reloc_UsedThisYear_db)
  Else
    rs.Fields("LastYear").value = ben.value(reloc_UsedLastyear_db)
  End If
  
  Call BenefitCloseRecord(ben, rs)
  
  If IsNull(rs.Fields("RelocKey").value) Or p11d32.BringForward.Yes Then
    'if is null, or bringing data forward then benefit is new so add in a new GUID
    rs.Edit
    rs.Fields("RelocKey") = GenerateGUID
    rs.Update
  End If
  
  ben.value(reloc_Key_db) = rs.Fields("RelocKey").value
  
  Call WriteDetails
  
  IBenefitClass_WriteDB = True
clsBenReloc_WriteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenReloc_WriteDB")
  Exit Function
clsBenReloc_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "clsBenReloc_WriteDB", "ERR_WRITE_DB", "Unable to write relocation details to the database.")
  Resume clsBenReloc_WriteDB_End
End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  On Error GoTo clsBenReloc_DeleteDB_Err
  
  Call xSet("clsBenReloc_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_RELOC)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
    Call m_Parent.Parent.db.Execute(sql.Queries(DELETE_RELOCDETAILS, m_BenItems(reloc_Key_db)))
  End If
  
  IBenefitClass_DeleteDB = True
  
clsBenReloc_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenReloc_DeleteDB")
  Exit Function
clsBenReloc_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenReloc_DeleteDB", "ERR_DELETE_DB", "Unable to delete the relocation benefit from the database.")
  Resume clsBenReloc_DeleteDB_End
  Resume
End Function



