VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0
Implements IBenefitClass
Implements ISortFunction

Public Enum OtherItems
  oth_item_db = ITEM_DESC
  oth_value = ITEM_VALUE
  oth_MadeGood_Net = ITEM_MADEGOOD_NET
  oth_benefit = ITEM_BENEFIT
  oth_madegood_db = ITEM_MADEGOOD
  oth_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  oth_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  oth_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  oth_BoxNumber = ITEM_BOX_NUMBER
  oth_MadeGood_Is_TaxDeducted_db = ITEM_MADEGOOD_IS_TAXDEDUCTED
  oth_Class1A_Adjustment = ITEM_CLASS1A_ADJUSTMENT
  oth_NIC_Class1A_Able = ITEM_NIC_CLASS1A_ABLE
  oth_NIC_Class1A_Value = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  oth_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  oth_Error = ITEM_ERROR
  oth_Value_Non_OPRA = ITEM_VALUE_NON_OPRA
  oth_OPRA_Ammount_Foregone_Used_For_Value = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  oth_OPRA_Ammount_Foregone_db = ITEM_OPRA_AMOUNT_FOREGONE
  
  
  oth_IRDesc_db
  
  oth_GrossAmountPaidByEmployer_db
  oth_employeereference ' EK added as missing - not actualy used.
  oth_category_db
  oth_HMITCode_UDBCode_db
  oth_CompanyDefinedCategory 'bf
  oth_CompanyDefinedCategoryKey_db 'bf
  oth_class_db
  oth_availablefrom_db
  oth_availableto_db
  oth_amountwaived
  oth_LASTITEM = oth_amountwaived
  
  'IF CHANGE HERE THEN CHANGE IN lOANS AND IN aCCOMMODATION
  
End Enum

Public Accommodation As IBenefitClass
Public Loan As IBenefitClass

Private m_Parent As Employee
Private m_BenClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_BenItems(1 To oth_LASTITEM)

Private m_dirty As Boolean
Private m_CompanyDefined_db As Boolean
Private m_InvalidFields As Long
Private m_LinkBen As Boolean
Private m_NeedToCalculate As Boolean
'Private m_CalculatingNIC As Boolean
Private m_CDBEmployee As IBenefitClass
Private m_CDBMasterBenefitLink As IBenefitClass

Private Sub Class_Terminate()
 'Debug.Print"Other - terminate"
End Sub
'cad cdb
Public Property Set CDBMasterBenefitLink(ByVal NewValue As IBenefitClass)
  Set m_CDBMasterBenefitLink = NewValue
End Property
Public Property Get CDBMasterBenefitLink() As IBenefitClass
  Set CDBMasterBenefitLink = m_CDBMasterBenefitLink
End Property


Private Function IBenefitClass_CalculateBody() As Variant
  Dim l As Long
  Dim ben As IBenefitClass
  Dim benCDC As IBenefitClass
  Dim b As Boolean
  
  On Error GoTo Other_Calculate_err
          
  IBenefitClass_CalculateBody = S_ERROR
  
  If Not Accommodation Is Nothing Then 'links for expenses
    Set ben = Accommodation
    Call ben.Calculate
  ElseIf Not Loan Is Nothing Then 'links for amount waived
    Set ben = Loan
    Call ben.Calculate
  Else
    
    Set ben = Me
    If Not CDBMasterBenefitLink Is Nothing Then
      Call CDBMasterBenefitLink.Calculate
    End If
    Call ben.SetCalcDefaults
    'get CDC description
    If ben.value(oth_CompanyDefinedCategoryKey_db) > 0 Then
      For l = 1 To m_Parent.Parent.CDCs.Count
        Set benCDC = m_Parent.Parent.CDCs(l)
        If Not benCDC Is Nothing Then

          If ben.value(oth_CompanyDefinedCategoryKey_db) = benCDC.value(cdc_Key_db) Then
            ben.value(oth_CompanyDefinedCategory) = benCDC.value(cdc_name_db)
            b = True
            Exit For
          End If
        End If
      Next
      If Not b Then ECASE ("CDC key is valid but no description found??.")
    Else
      ben.value(oth_CompanyDefinedCategory) = ""
    End If
    ben.value(oth_madegood_db) = ben.value(oth_madegood_db)
    ben.value(oth_Value_Non_OPRA) = RoundDownEx(ben.value(oth_GrossAmountPaidByEmployer_db) * dGetDateFactor(l, ben.value(oth_availablefrom_db), ben.value(oth_availableto_db)), 0)
    
    Call CalculateOpRAValue(ben)
    ben.value(oth_MadeGood_Net) = Min(ben.value(ITEM_VALUE), ben.value(oth_madegood_db))
    
    ben.value(oth_benefit) = ben.value(ITEM_VALUE) - ben.value(oth_MadeGood_Net)
    If (m_BenClass = BC_PAYMENTS_ON_BEFALF_B Or m_BenClass = BC_TAX_NOTIONAL_PAYMENTS_B) Then
      If ben.value(oth_benefit) > 0 Then ben.value(oth_Benefit_Reportable) = True
    Else
      If ben.value(ITEM_VALUE) > 0 Then ben.value(oth_Benefit_Reportable) = True
    End If
  End If
  'cad loan fix
  Call BenCalcNIC(ben, AdjustItemEnum(ITEM_BENEFIT_SUBJECT_TO_CLASS1A), AdjustItemEnum(ITEM_NIC_CLASS1A_BENEFIT), AdjustItemEnum(ITEM_VALUE), AdjustItemEnum(ITEM_BENEFIT), AdjustItemEnum(ITEM_CLASS1A_ADJUSTMENT), AdjustItemEnum(ITEM_MADEGOOD_IS_TAXDEDUCTED), AdjustItemEnum(ITEM_MADEGOOD_NET), AdjustItemEnum(ITEM_NIC_CLASS1A_ABLE))
  IBenefitClass_CalculateBody = ben.value(AdjustItemEnum(ITEM_BENEFIT))
   
  
Other_Calculate_end:
  Exit Function
Other_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume Other_Calculate_end

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  Dim s As String

  Select Case m_BenClass
    Case BC_PRIVATE_MEDICAL_I
      s = "Medical"
    Case BC_CLASS_1A_M, BC_VOUCHERS_AND_CREDITCARDS_C
      s = "Vouchers"
    Case Else
      s = "Other"
  End Select

  IBenefitClass_ImageListKey = s
End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = m_NeedToCalculate
End Property

Private Property Let IBenefitClass_LinkBen(RHS As Boolean)
  m_LinkBen = RHS
End Property
Private Property Get IBenefitClass_LinkBen() As Boolean
  IBenefitClass_LinkBen = m_LinkBen
End Property
Private Function IBenefitClass_CanBringForward() As Boolean
  Dim ben As IBenefitClass
  Dim ee As Employee
  
  Set ben = Me
  
  If ben.CompanyDefined Then
    Set ee = ben.Parent
    If IsCBDEmployee(ee.PersonnelNumber) Then
      IBenefitClass_CanBringForward = True
    End If
  Else
    IBenefitClass_CanBringForward = StandardCanBringForward(ben, oth_availableto_db)
  End If
  
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim other As IBenefitClass
  
  Set other = New other
  
  Set IBenefitClass_Copy = CopyBenStandard(Parent, other, Me)
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
  
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Dim ben As IBenefitClass
  Dim benefit, value, MadeGood
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim vbenOthers() As Variant
  Dim ee As Employee
  Dim lCategoryTotal As Long
  Dim vTotalBen As Variant
  Dim sCurrency As String
  
  Dim i As Long, j As Long, lCategoryKey As Long, lRowCount As Long
  
  Dim lGrossAmountPaidByEmployer As Long 'cad 20/02 review lGrossAmountPaidByEmployer, changed by CAD cos rdc no do
  
  On Error GoTo Other_PrintWK_Err
  
  Call xSet("Other_PrintWK")
  
  'other type benefits differ as they print a total of all of the same type
  If p11d32.CurrentEmployer Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "PrintWkBody", "The current employer is nothing.")
  Set ee = p11d32.CurrentEmployer.CurrentEmployee
  If ee Is Nothing Then Call Err.Raise(ERR_EMPLOYEE_IS_NOTHING, "PrintWkBody", "Parent is nothing in other benefit, other name:" & IBenefitClass_Name & ".")
  
  If m_BenClass = BC_OOTHER_N Then
    'JN to see
    Call WKTblColXOffsets(L_WK_OTHER_TABLE_COL1, L_WK_OTHER_TABLE_COL1 + 30, L_WK_OTHER_TABLE_COL2, L_WK_OTHER_TABLE_COL3, L_WK_OTHER_TABLE_COL4)
    Call WKTblColFormats("n", "n", "rn", "rn", "rn")
    Call WKTableHeadings(rep, "Description", "Company defined category", "Value", "Made good", "Benefit")
  Else
    If p11d32.ReportPrint.DatesOnWorkingPaper = True Then  ' rdc added 12/02/02
     Call WKTblOtherTypeTableWithDates(rep, "Description", S_FIELD_AVAILABLE_FROM, "Available to", "No. of days", "Gross amount", "Made Good", "Benefit") 'rdc added 12/02/02
    Else
     Call WKTblOtherTypeTable(rep, "Description", "Value", "Made good", "Benefit")
    End If
  End If
  
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = m_BenClass Then
        j = j + 1
        ReDim Preserve vbenOthers(1 To j)
        Set vbenOthers(j) = ben
      End If
    End If
  Next
  
  Call QSortEx(vbenOthers, 1, j, Me)
  'this is crap CAD
  vTotalBen = 0
  
  For i = 1 To j
    Set ben = vbenOthers(i)
    If i = 1 Then lCategoryKey = ben.value(oth_CompanyDefinedCategoryKey_db)
    If m_BenClass = BC_OOTHER_N Then
      If (ben.value(oth_CompanyDefinedCategoryKey_db) <> lCategoryKey) Then
        Call WKTableTotals(rep, "", "", "", "", vTotalBen)
        Call rep.Out(vbCrLf & vbCrLf)
        Call WKTableRow(rep, ben.value(ITEM_DESC), ben.value(oth_CompanyDefinedCategory), FormatWN(ben.value(ITEM_VALUE), ""), FormatWN(ben.value(ITEM_MADEGOOD), "", True), FormatWN(ben.value(ITEM_BENEFIT), ""))
        vTotalBen = 0
      ElseIf i = j Then
        Call WKTableRow(rep, ben.value(ITEM_DESC), ben.value(oth_CompanyDefinedCategory), FormatWN(ben.value(ITEM_VALUE), ""), FormatWN(ben.value(ITEM_MADEGOOD), "", True), FormatWN(ben.value(ITEM_BENEFIT), ""))
        'RUnwin only total for >1 benefit
        If Not i = 1 Then Call WKTableTotals(rep, "", "", "", "", FormatWN(vTotalBen + ben.value(ITEM_BENEFIT)))
        vTotalBen = 0
      Else
        If i > 1 Then
          Call WKTableRow(rep, ben.value(ITEM_DESC), ben.value(oth_CompanyDefinedCategory), FormatWN(ben.value(ITEM_VALUE)), FormatWN(ben.value(ITEM_MADEGOOD), , True), FormatWN(ben.value(ITEM_BENEFIT)))
        Else
          Call WKTableRow(rep, ben.value(ITEM_DESC), ben.value(oth_CompanyDefinedCategory), FormatWN(ben.value(ITEM_VALUE), ""), FormatWN(ben.value(ITEM_MADEGOOD), "", True), FormatWN(ben.value(ITEM_BENEFIT), ""))
        End If
      End If
      vTotalBen = vTotalBen + ben.value(ITEM_BENEFIT)
    Else
      'this is a copy of above less company defined category
      If i = 1 Then
        sCurrency = S_CURRENCY
      Else
        sCurrency = ""
      End If
      Call NonOtherTableRows(lGrossAmountPaidByEmployer, rep, ben, sCurrency)
    End If
    lCategoryKey = ben.value(oth_CompanyDefinedCategoryKey_db)
  Next
  
  'print sum total
  BenArr(1) = m_BenClass
  Call ee.SumBenefit("", value, MadeGood, benefit, BenArr(), , ITEM_MADEGOOD)
  If m_BenClass = BC_OOTHER_N Then
    Call rep.Out(vbCrLf & vbCrLf)
    Call WKTableTotals(rep, "", "", FormatWN(value), FormatWN(MadeGood, , True), FormatWN(benefit))
  Else
    If p11d32.ReportPrint.DatesOnWorkingPaper = True Then  ' rdc added 12/02/02
'MP RPT CHANGES HERE for TTP WK items
      Call WKTableTotals(rep, "", "", "", "", FormatWN(lGrossAmountPaidByEmployer, , True), FormatWN(MadeGood, , True), FormatWN(benefit))
    Else
      Call WKTableTotals(rep, "", FormatWN(value), FormatWN(MadeGood, , True), FormatWN(benefit))
    End If
 End If
  
Other_PrintWK_End:
  Call xReturn("Other_PrintWK")
  Exit Function
Other_PrintWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Other_PrintWK", "Other Print WK", "Error printing the 'other' benefit work sheet.")
  Resume Other_PrintWK_End
  Resume
End Function

Private Sub NonOtherTableRows(ByRef lGrossAmountPaidByEmployer As Long, ByVal rep As Reporter, ByVal ben As IBenefitClass, Optional ByVal sCurrency As String = S_CURRENCY)
'  If p11d32.AppYear > 2000 Then
    If p11d32.ReportPrint.DatesOnWorkingPaper = True Then  ' rdc added 12/02/02
      Call WKTableRow(rep, ben.value(ITEM_DESC), ben.value(oth_availablefrom_db), ben.value(oth_availableto_db), DateDiff("d", DateAdd("D", -1, ben.value(oth_availablefrom_db)), ben.value(oth_availableto_db)), FormatWN(ben.value(oth_GrossAmountPaidByEmployer_db)), FormatWN(ben.value(ITEM_MADEGOOD), , True), FormatWN(ben.value(ITEM_BENEFIT)))  ' RDC added 12/02/02
      'Call WKTableRow(rep, ben.value(ITEM_DESC), ben.value(oth_availablefrom), ben.value(oth_availableto), DateDiff("d", ben.value(oth_availablefrom) - 1, ben.value(oth_availableto)), FormatWN(ben.value(oth_GrossAmountPaidByEmployer)), FormatWN(ben.value(ITEM_MADEGOOD), , True), FormatWN(ben.value(ITEM_BENEFIT))) ' RDC added 12/02/02
    Else
      Call WKTableRow(rep, ben.value(ITEM_DESC), FormatWN(ben.value(ITEM_VALUE), sCurrency), FormatWN(ben.value(ITEM_MADEGOOD), sCurrency, True), FormatWN(ben.value(ITEM_BENEFIT), sCurrency))
    End If
'  Else
'    Call WKTableRow(rep, ben.value(ITEM_DESC), FormatWN(ben.value(ITEM_VALUE), sCurrency), FormatWN(ben.value(ITEM_MADEGOOD), sCurrency, True), FormatWN(ben.value(ITEM_BENEFIT), sCurrency))
'  End If

End Sub
Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property
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
  m_CompanyDefined_db = NewValue
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = m_CompanyDefined_db
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Sub IBenefitClass_SetCalcDefaults()
  m_BenItems(oth_Benefit_Reportable) = False
  m_BenItems(oth_OPRA_Ammount_Foregone_Used_For_Value) = False
  m_BenItems(oth_Value_Non_OPRA) = S_ERROR
  m_BenItems(oth_value) = S_ERROR
  m_BenItems(oth_benefit) = S_ERROR
  m_BenItems(oth_Error) = ""
  
End Sub
Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
  Set Accommodation = Nothing
  Set Loan = Nothing
End Sub
Private Property Get IBenefitClass_Name() As String
  Dim ben As IBenefitClass
  Set ben = Me
  IBenefitClass_Name = ben.value(ITEM_DESC)
  
End Property
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property
Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
Private Function IBenefitClass_ReadDB() As Long
  Dim other As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  Dim i As Long
  Dim Rates As Rates
  Dim bc As BEN_CLASS
  Dim IRDesc As Variant
  
  On Error GoTo clsOther_ReadDB_err
  Call xSet("clsOther_ReadDB")
  
  If m_ReadFromDB Then GoTo clsOther_ReadDB_end
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(IBenefitClass_TABLE)
  
  If Not IBenefitClass_HasBookMark Then
    If m_CompanyDefined_db Then
      rs.FindFirst ("UDB = " & True)
    Else
      Set ben = m_Parent
      s = ben.value(ee_PersonnelNumber_db)
      rs.FindFirst ("P_NUM = '" & s & "'")
    End If
    If Not rs.NoMatch Then
      i = i + 1
      IBenefitClass_RSBookMark = rs.Bookmark
      If m_CompanyDefined_db Then
        Call rs.FindNext("UDB = " & True)
      Else
        rs.FindNext ("P_NUM = '" & s & "'")
      End If
      Do While Not rs.NoMatch
        i = i + 1
        Set other = New other
        other.RSBookMark = rs.Bookmark
        other.CompanyDefined = m_CompanyDefined_db
        Set other.Parent = m_Parent
        other.BenefitClass = p11d32.Rates.DBClassToBCClass("" & rs.Fields("Class").value)
        Call m_Parent.benefits.Add(other)
        Set other = Nothing
        If m_CompanyDefined_db Then
          Call rs.FindNext("UDB = " & True)
        Else
          rs.FindNext ("P_NUM = '" & s & "'")
        End If
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  
  If IBenefitClass_HasBookMark And m_ReadFromDB = False Then
    Set ben = Me
    rs.Bookmark = IBenefitClass_RSBookMark

    m_BenClass = p11d32.Rates.DBClassToBCClass("" & rs.Fields("Class").value)
    Call StandardReadData(ben, rs)
    ben.value(oth_class_db) = "" & rs.Fields("Class").value
    ben.value(oth_category_db) = "" & rs.Fields("Category").value
    ben.value(oth_HMITCode_UDBCode_db) = "" & rs.Fields("UDBCode").value
    ben.value(oth_item_db) = "" & rs.Fields("Item").value
    ben.value(oth_CompanyDefinedCategoryKey_db) = IIf(IsNull(rs.Fields("CDCKey").value), 0, rs.Fields("CDCKey").value)
    
    If m_CompanyDefined_db Then
      ben.value(oth_availablefrom_db) = p11d32.Rates.value(TaxYearStart)
      ben.value(oth_availableto_db) = p11d32.Rates.value(TaxYearEnd)
    Else
      ben.value(oth_availablefrom_db) = IIf(IsNull(rs.Fields("From").value), p11d32.Rates.value(TaxYearStart), rs.Fields("From").value)
      ben.value(oth_availableto_db) = IIf(IsNull(rs.Fields("To").value), p11d32.Rates.value(TaxYearEnd), rs.Fields("To").value)
    End If
    
    ben.value(oth_GrossAmountPaidByEmployer_db) = IIf(IsNull(rs.Fields("Value").value), 0, rs.Fields("Value").value)
    ben.value(oth_madegood_db) = IIf(IsNull(rs.Fields("Madegood").value), 0, rs.Fields("Madegood").value)
    If (HasIRDescription(m_BenClass)) Then
      Call IRDescDB(ben, rs, True)
    End If
    
    m_ReadFromDB = True
    ben.Dirty = False
  End If
  
clsOther_ReadDB_end:
  Set ben = Nothing
  IBenefitClass_ReadDB = i
  xReturn "clsOther_ReadDB"
  Exit Function
clsOther_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "Other_ReadDB", "ERR_READ_OTHER_BENEFITS", "Unable to read in the other benefits."
   Resume clsOther_ReadDB_end
   Resume
End Function
Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property
Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property
Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  On Error GoTo SetBenItemsInformation_err
    
  m_NeedToCalculate = True
    
  With p11d32
    For bc = [BC_FIRST_ITEM] To [BC_REAL_BENEFITS_LAST_ITEM]
      If IsBenOtherClass(bc) Then
        
        If .DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
        
        If bc = BC_ENTERTAINMENT_N Then
          bc = BC_ENTERTAINMENT_N
        End If
        Set ben = Me
        p11d32.BenDataLinkBenfitTable(bc) = ben.TABLE
        .BenDataLinkDataType(bc, oth_GrossAmountPaidByEmployer_db) = TYPE_LONG
        .BenDataLinkDataType(bc, oth_category_db) = TYPE_STR
        .BenDataLinkDataType(bc, oth_HMITCode_UDBCode_db) = TYPE_STR
        .BenDataLinkDataType(bc, oth_CompanyDefinedCategory) = TYPE_STR
        .BenDataLinkDataType(bc, oth_CompanyDefinedCategoryKey_db) = TYPE_LONG
        .BenDataLinkDataType(bc, oth_class_db) = TYPE_STR
        .BenDataLinkDataType(bc, oth_MadeGood_Is_TaxDeducted_db) = TYPE_BOOL
        .BenDataLinkDataType(bc, oth_IRDesc_db) = TYPE_LONG
        .BenDataLinkDataType(bc, oth_availablefrom_db) = TYPE_DATE
        .BenDataLinkDataType(bc, oth_availableto_db) = TYPE_DATE
        
        Call SetStandardBenItemsDataTypes(bc)
        Call SetStandardBenItemsMMFieldSize(bc)
        Call SetStandardBenItemsUDMData(bc)
        .BenDataLinkUDMDescription(bc, oth_availablefrom_db) = S_UDM_FROM
        .BenDataLinkUDMDescription(bc, oth_availableto_db) = S_UDM_To
  
          
        If bc = BC_OOTHER_N Then .BenDataLinkUDMDescription(bc, oth_CompanyDefinedCategory) = "Company defined category"
              
        If (HasIRDescription(bc)) Then
          Call SetIRDesriptionInformation(bc)
        End If
        
        .DataLinkInitialised(bc) = True
      End If
    Next
  End With

SetBenItemsInformation_end:
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information.")
  Resume SetBenItemsInformation_end
End Sub

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
On Error GoTo Other_DeleteDB_Err
  
  Call xSet("Other_DeleteDB")
  
  Set ben = Me
  Set rs = m_Parent.Parent.rsBenTables(ben.TABLE)
  If IBenefitClass_HasBookMark Then
    rs.Bookmark = IBenefitClass_RSBookMark
    rs.Delete
    IBenefitClass_RSBookMark = ""
  End If
  IBenefitClass_DeleteDB = True
  
Other_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("Other_DeleteDB")
  Exit Function

Other_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Other_DeleteDB", "Other Delete DB", "Error deleting the 'other' benefit from the database.")
  Resume Other_DeleteDB_End
End Function

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  Dim ben As IBenefitClass
  
  Set ben = Me
  If Not Accommodation Is Nothing Then
    Accommodation.value(AdjustItemEnum(Item)) = RHS
  ElseIf Not Loan Is Nothing Then
    Loan.value(AdjustItemEnum(Item)) = RHS
  Else
    m_BenItems(Item) = CorrectBenValue(m_BenClass, Item, RHS)
  End If
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  On Error GoTo Err_Err
  Dim ben As IBenefitClass
  Dim benOther As other
  
  Set ben = Me
  If Not Accommodation Is Nothing Then
    IBenefitClass_value = Accommodation.value(AdjustItemEnum(Item))
  ElseIf Not Loan Is Nothing Then
    IBenefitClass_value = Loan.value(AdjustItemEnum(Item))
  Else
    IBenefitClass_value = m_BenItems(Item)
    If ben.CompanyDefined Then
      If CDBMasterBenefitLink Is Nothing Then
        IBenefitClass_value = m_BenItems(Item)
      Else
          'cad cdb thing
        If Item >= ITEM_CLASS1A_ADJUSTMENT And Item <= ITEM_NIC_CLASS1A_BENEFIT Then
          IBenefitClass_value = m_BenItems(Item)
        Else
          IBenefitClass_value = m_CDBMasterBenefitLink.value(Item)
        End If
      End If
    Else
      IBenefitClass_value = m_BenItems(Item)
    End If
  End If
  'cad cdb thing
  
Err_End:
  
  Exit Property
Err_Err:
  Call Err.Raise(Err.Description, ErrorSource(Err, "Value"), Err.Description)
  Resume
End Property
Private Property Get CDBMasterCopy() As Boolean
  Dim ben As IBenefitClass
  Dim benEE As IBenefitClass
  
  Set ben = Me
  If ben.CompanyDefined Then
    Set benEE = GetParentFromBenefit(ben, GPBF_EMPLOYEE)
    CDBMasterCopy = IsCBDEmployee(benEE.value(ee_PersonnelNumber_db))
  End If
End Property
Private Property Get CanWriteDB() As Boolean
  Dim ben As IBenefitClass
  
  Set ben = Me
  If ben.CompanyDefined Then
    CanWriteDB = CDBMasterCopy
  Else
    CanWriteDB = True
  End If
  
End Property
Private Function AdjustItemEnum(ByVal Item As OtherItems) As Long
  On Error GoTo AdjustItemEnum_ERR
  Call xSet("AdjustItemEnum")
    
  If Not Me.Accommodation Is Nothing Then
    AdjustItemEnum = Item + (accom_EXP_SPECIFIC_START - 1)
    If (AdjustItemEnum < accom_EXP_SPECIFIC_START) Or (AdjustItemEnum > accom_EXP_SPECIFIC_END) Then
      Call Err.Raise(ERR_ENUM__VALUE_INVALID, "AdjustItemEnum", "Enum value is not valid for OOther benefit linked to accommodation.")
    End If
  ElseIf Not Me.Loan Is Nothing Then
    AdjustItemEnum = Item + (ln_WAIVED_SPECIFIC_START - 1)
    If (AdjustItemEnum < ln_WAIVED_SPECIFIC_START) Or (AdjustItemEnum > ln_WAIVED_SPECIFIC_END) Then
      Call Err.Raise(ERR_ENUM__VALUE_INVALID, "AdjustItemEnum", "Enum value is not valid for OOther benefit linked to loan.")
    End If
  Else
    AdjustItemEnum = Item
  End If
  
AdjustItemEnum_END:
  Call xReturn("AdjustItemEnum")
  Exit Function
AdjustItemEnum_ERR:
'MP DB  Call ErrorMessage(ERR_ERROR, Err, "AdjustItemEnum", "Adjust Item Enum", "The Item number passed is outside the enum range of RelocItems.")
  Call ErrorMessage(ERR_ERROR, Err, "AdjustItemEnum", "Adjust Item Enum", "The Item number passed is outside the enum range of OtherItems.")
  Resume AdjustItemEnum_END
  Resume
End Function
Private Function IsLinkBen() As Boolean
  IsLinkBen = (Not Me.Accommodation Is Nothing) Or (Not Me.Loan Is Nothing)
End Function
Public Function PersonnelNumber() As String
  Dim ben As IBenefitClass
  Set ben = Me
  If m_CompanyDefined_db Then
    PersonnelNumber = GetEmployeeNumber(ben) & ben.value(oth_HMITCode_UDBCode_db)
  Else
    PersonnelNumber = GetEmployeeNumber(ben)
  End If
End Function
'cad cdb
Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim l As Long
  Dim ben As IBenefitClass
  Dim benEY As IBenefitClass
  Dim HMITS As HMIT_SECTIONS
  Dim benEE As IBenefitClass
  Dim employees As ObjectList
  Dim ee As Employee
  Dim benCDB As IBenefitClass
  
  Dim iBenIndex As Long
  
  On Error GoTo Other_WriteDB_Err
  Call xSet("Other_WriteDB")
  
  Set ben = Me
  
  'cad cdb stuff prevent writing of cdb link bens
  If Not CanWriteDB Then
    ben.Dirty = False
    IBenefitClass_WriteDB = True
    GoTo Other_WriteDB_End
  End If
    
  If ben.CompanyDefined Then
    'make all the emmployees with benefits linking to myself dirty myself
    Set employees = p11d32.CurrentEmployer.employees
    
    For l = 1 To employees.Count
      Set ee = employees(l)
      If Not ee Is Nothing Then
        iBenIndex = ee.HasCDBBenefit(ben)
        If iBenIndex > 0 Then
          Set benCDB = ee.benefits(iBenIndex)
          benCDB.NeedToCalculate = True
        End If
      End If
    Next
  End If
  
  If ben.LinkBen Then
    IBenefitClass_WriteDB = True
    GoTo Other_WriteDB_End
  End If
  
  Set rs = m_Parent.Parent.rsBenTables(TBL_OTHER)
  
  Call BenefitAddNewRecord(ben, rs)
  Set benEY = ben.Parent
  Call BringForwardDatesWrite(ben, oth_availablefrom_db, oth_availableto_db, rs, "From", "To")
  'need to write out pNum as can be un company defined
  rs.Fields(S_FIELD_PERSONEL_NUMBER).value = benEY.value(ee_PersonnelNumber_db)
  
 
  rs.Fields("Class").value = p11d32.Rates.BenClassTo(m_BenClass, BCT_DBCLASS)
  rs.Fields("Category").value = p11d32.Rates.BenClassTo(m_BenClass, BCT_DBCATEGORY)
  rs.Fields("UDB").value = m_CompanyDefined_db
  If m_CompanyDefined_db Then
    rs.Fields("UDBCode").value = ben.value(oth_HMITCode_UDBCode_db)
  Else
    rs.Fields("UDBCode").value = p11d32.Rates.BenClassTo(m_BenClass, BCT_HMIT_SECTION_STRING)
  End If
  rs.Fields("Item").value = ben.value(oth_item_db)
  
  rs.Fields("CDCKey").value = ben.value(oth_CompanyDefinedCategoryKey_db)
  
  Call StandardWriteData(ben, rs)
  
  If (HasIRDescription(m_BenClass)) Then
    Call IRDescDB(ben, rs, False)
  End If
  
  If Not p11d32.BringForward.Yes Then
    rs.Fields("Value").value = ben.value(oth_GrossAmountPaidByEmployer_db)
    rs.Fields("Madegood").value = ben.value(oth_madegood_db)
  End If
  
  
  IBenefitClass_WriteDB = BenefitCloseRecord(ben, rs)
  
Other_WriteDB_End:
  Set rs = Nothing
  Call xReturn("Other_WriteDB")
  Exit Function
  
Other_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "Other_WriteDB", "Other Write DB", "Error writing the 'other' benefit to the database.")
  Resume Other_WriteDB_End
  Resume
End Function


Private Property Get IBenefitClass_PrintHeader() As String
  IBenefitClass_PrintHeader = "{Arial=8,n}{x=75}Amount made" & vbCrLf & "{x=60}Cost to" & _
    "{x=90}Cash{x=75}good or taxed" & vbCrLf & _
    "Description{x=60}you{x=75}under PAYE" & _
    "{x=90}equivalent" & vbCrLf & "{x=67}" & S_CURRENCY & "   {x=82}" & S_CURRENCY & "   {Arial=8,ri}{x=97}" & S_CURRENCY & "   " & vbCrLf & vbCrLf
End Property
Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call PrintWKHelper(rep, Me, False)
End Function
Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_OTHER
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property


Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim ben1 As IBenefitClass, ben2 As IBenefitClass
  
  Set ben1 = v0
  Set ben2 = v1

  If p11d32.ReportPrint.SortOtherTypeBenefitsAlphabetically Then
    ISortFunction_CompareItems = StrComp(ben1.value(ITEM_DESC), ben2.value(ITEM_DESC))
  Else
    ISortFunction_CompareItems = StrComp(ben1.value(oth_CompanyDefinedCategory), ben2.value(oth_CompanyDefinedCategory))
  End If
  
  
End Function



