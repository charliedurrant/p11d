VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IBenefitClass
Implements ISortFunction

Private m_BenClass As BEN_CLASS
Private m_Parent As IBenefitClass
Private m_ReadFromDB
Private m_loans As ObjectList
Private m_dirty As Boolean
Private m_InvalidFields As Long
Private m_bAllLoansDeminimus As Long
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Loans - terminate"
End Sub

Private Property Get IBenefitClass_ImageListKey() As String

End Property

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim ben0 As IBenefitClass, ben1 As IBenefitClass
  
  On Error GoTo Loan_ISortFunction_ERR
  
  Set ben0 = v0
  Set ben1 = v1
  
  
  'also sort cars
  If (Not BenefitIsLoan(ben0.BenefitClass)) Or (Not BenefitIsLoan(ben1.BenefitClass)) Then Call Err.Raise(ERR_BENCLASS_INVALID, "Loan_ISortFunction", "One of the benefits sorted is not a loan.")
  If ben0.value(ln_DateDischargedOrMade) < ben1.value(ln_DateDischargedOrMade) Then ISortFunction_CompareItems = -1
  If ben0.value(ln_DateDischargedOrMade) > ben1.value(ln_DateDischargedOrMade) Then ISortFunction_CompareItems = 1
  
  
Loan_ISortFunction_END:
  Exit Function
Loan_ISortFunction_ERR:
  Call ErrorMessage(ERR_ALLOWIGNORE, Err, "Loan_ISortFunction", "Loan ISort Function", "Error sorting loans for printing.")
  Resume Loan_ISortFunction_END
End Function
Public Function SortLoans() As ObjectList
  
  On Error GoTo SortLoans_ERR
  
  'must not do on the current page as will break the listviews
  If (CurrentForm Is F_AllBenefits) Or (CurrentForm Is F_Loan) Then GoTo SortLoans_END
  Call m_loans.Sort(Me)
  
SortLoans_END:
  Exit Function
SortLoans_ERR:
  Call ErrorMessage(ERR_ALLOWIGNORE, Err, "SortLoans", "Sort Loans", "Error sorting loans.")
  Resume SortLoans_END
End Function

Private Function CalculateLoans() As Variant
  Dim i As Long
  Dim Loan As Loan
  Dim ben As IBenefitClass
  Dim vBen As Variant
  For i = 1 To m_loans.Count
    Set Loan = m_loans(i)
    If Not Loan Is Nothing Then
      Set ben = m_loans(i)
      Call ben.SetCalcDefaults
      vBen = vBen + Loan.SubCalculate
    End If
  Next
  
  CalculateLoans = vBen
End Function
Private Function IBenefitClass_CalculateBody() As Variant
  Dim Loan As Loan
  Dim ben As IBenefitClass
  Dim i As Long
  Dim vBen As Variant
  Dim ibf As IBenefitForm2
  Dim li As ListItem
  
  On Error GoTo Loans_Calculate_err
  Call xSet("Loans_Calculate")
  
  'If the whole or part is deminimus then do not report this part of the benefit
  m_bAllLoansDeminimus = False
  vBen = CalculateLoans() 'calc the loans with deminimus false so all are included for calc
  m_bAllLoansDeminimus = IsDeminimus
  vBen = CalculateLoans()
  Call RedrawFormLoanItems
  
  IBenefitClass_CalculateBody = vBen
  
Loans_Calculate_end:
  xReturn "Loans_Calculate"
  Exit Function
Loans_Calculate_err:
  IBenefitClass_CalculateBody = S_ERROR
  Resume Loans_Calculate_end
End Function
Public Sub RedrawFormLoanItems()
  Dim li As ListItem
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  
  If CurrentForm Is F_Loan Then
    Set ibf = CurrentForm
    For Each li In ibf.lv.listitems
      Set ben = m_loans(li.Tag)
      If Not ben Is ibf.benefit Then
        Call F_Loan.UpdateBenefitListViewItemEx(li, ben, li.Tag, False, True)
      End If
    Next
  End If

End Sub
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
  'let loan deal with it
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function

Public Property Get LoansCount() As Boolean
  LoansCount = m_loans.CountValid
End Property
Public Property Get AllLoansDeminimus() As Boolean
  AllLoansDeminimus = m_bAllLoansDeminimus
End Property
Public Property Let AllLoansDeminimus(NewValue As Boolean)
  m_bAllLoansDeminimus = NewValue
End Property
Public Property Get loans() As ObjectList
  Set loans = m_loans
End Property

Private Sub Class_Initialize()
  
  
  Set m_loans = New ObjectList
  
  Call IBenefitClass_SetBenItemsInformation
  
  m_bAllLoansDeminimus = -1
  
End Sub
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_LOANS_H
End Property
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property

Public Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "Loans_CompanyDefined"
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  
End Property
Private Function IBenefitClass_DeleteDB() As Boolean
  Dim i As Long
  
  Dim ben As IBenefitClass
  
  For i = 1 To m_loans.Count
    Set ben = m_loans(i)
    If Not ben Is Nothing Then
      Call ben.DeleteDB
    End If
  Next
  
  Set ben = Nothing
  
End Function

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  Dim ben As IBenefitClass
  Dim i As Long
  
  For i = 1 To m_loans.Count
    Set ben = m_loans(i)
    If Not ben Is Nothing Then
      ben.Dirty = NewValue
    End If
  Next i
  Call DirtyHelper(Me, NewValue)

  

End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  Dim Loan As Loan
  Dim i As Long
  Dim ben As IBenefitClass
  
  For i = 1 To m_loans.Count
    Set Loan = m_loans(i)
    If Not Loan Is Nothing Then
      Set ben = Loan
      If ben.Dirty Then
        IBenefitClass_Dirty = True
        Exit For
      End If
    End If
  Next i
  
  Set Loan = Nothing
End Property


Private Property Get IBenefitClass_HasBookMark() As Boolean
  'cd IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Sub IBenefitClass_Kill()
  Dim i As Long
  Dim Loan As IBenefitClass
  
  For i = 1 To m_loans.Count
    Set Loan = m_loans(i)
    If Not Loan Is Nothing Then
      Loan.Kill
    End If
  Next i
  m_loans.RemoveAll
  Set m_loans = Nothing
  Set m_Parent = Nothing
End Sub



Private Sub IBenefitClass_MakeDirty()

End Sub

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = "Loans"
End Property

Public Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Public Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_PrintHeader() As String
  ECASE "Loans_PrintHeader"
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  ECASE "Loans_PrintWK"
End Function

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Function IBenefitClass_ReadDB() As Long
  'read all loans
  Dim rs As Recordset
  Dim Loan As Loan, ben As IBenefitClass
  Dim s As String
  Dim i As Long, j As Long
  Dim ee As Employee
  
  On Error GoTo Loans_readdb_Err
  Call xSet("Loans_readdb")
  If m_ReadFromDB Then GoTo Loans_readdb_End
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_BENLOAN) ' normal
  Set ee = m_Parent
  rs.FindFirst ("P_NUM = '" & ee.PersonnelNumber & "'")
    Do While Not rs.NoMatch
      i = i + 1
      Set Loan = New Loan
      Set ben = Loan
      ben.RSBookMark = rs.Bookmark
      Set ben.Parent = Me
      ben.BenefitClass = BC_LOAN_OTHER_H
      Call m_loans.Add(ben)
      Set Loan = Nothing
      Set ben = Nothing
      Set ben = m_Parent
      rs.FindNext ("P_NUM = '" & ben.value(ee_PersonnelNumber_db) & "'")
    Loop
  
    
  For j = 1 To m_loans.Count
    Set ben = m_loans(j)
    If Not ben Is Nothing Then ben.ReadDB
  Next j
  m_ReadFromDB = True
Loans_readdb_End:
  IBenefitClass_ReadDB = i
  Set rs = Nothing
  Call xReturn("Loans_readdb")
  Exit Function
Loans_readdb_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Loans_readdb", "ERR_LOANCOL_READDB", "Unable to read the loan details.")
  Resume Loans_readdb_End
  Resume
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
    Dim ben As IBenefitClass
  Dim i As Long
  
  For i = 1 To m_loans.Count
    Set ben = m_loans(i)
    If Not ben Is Nothing Then
      ben.ReadFromDB = NewValue
    End If
  Next i
  
  
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property


Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  ECASE "Loans_sBookmark"
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  ECASE "Loans_RSBookMark"
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  Dim Loan As Loan
  
  On Error GoTo SetBenItemsInformation_err
  bc = BC_LOANS_H
  m_NeedToCalculate = True
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  Set Loan = New Loan

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub

Private Sub IBenefitClass_SetCalcDefaults()
  
End Sub

Private Function IBenefitClass_SetValue(ByVal Item As Long, value As Variant) As Boolean
  ECASE "Loans_SetValue"
End Function

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  Dim ben As IBenefitClass
  Dim i As Long
  
  IBenefitClass_value = 0
  Select Case Item
    Case ITEM_BENEFIT, ITEM_MADEGOOD, ITEM_VALUE, ITEM_MADEGOOD_NET, ITEM_VALUE, ITEM_DESC, ITEM_MADEGOOD_IS_TAXDEDUCTED, ITEM_CLASS1A_ADJUSTMENT, ITEM_NIC_CLASS1A_ABLE, ITEM_BENEFIT_SUBJECT_TO_CLASS1A, ITEM_NIC_CLASS1A_BENEFIT
      For i = 1 To loans.Count
        Set ben = loans(i)
        If Not ben Is Nothing Then
          IBenefitClass_value = IBenefitClass_value + ben.value(Item)
        End If
      Next
    Case ITEM_BENEFIT_REPORTABLE
      For i = 1 To loans.Count
        Set ben = loans(i)
        If Not ben Is Nothing Then
          IBenefitClass_value = IBenefitClass_value Or ben.value(Item)
        End If
      Next
    Case ITEM_ERROR
      For i = 1 To loans.Count
        Set ben = loans(i)
        If Not ben Is Nothing Then
          If Len(ben.value(ITEM_ERROR)) > 0 Then
            IBenefitClass_value = ben.value(ITEM_ERROR)
            Exit For
          End If
        End If
      Next
      If IBenefitClass_value = 0 Then
        IBenefitClass_value = ""
      End If
    Case Else
      ECASE "Invalid loans value request"
  End Select
   
  If IsEmpty(IBenefitClass_value) Then IBenefitClass_value = False
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim ben As IBenefitClass
  Dim i As Long
  Dim lWritten As Long
  Dim lBenefitsDirty As Long
  
  For i = 1 To m_loans.Count
    Set ben = m_loans(i)
    If Not ben Is Nothing Then
      If ben.Dirty Or p11d32.BringForward.Yes Then
        lBenefitsDirty = lBenefitsDirty + 1
        If ben.InvalidFields = 0 Then
          lWritten = lWritten + 1
          If p11d32.BringForward.Yes Then
            If ben.CanBringForward Then Call ben.WriteDB
          Else
            Call ben.WriteDB
          End If
        End If
      End If
    End If
  Next i
  
  If lBenefitsDirty - lWritten = 0 Then
    IBenefitClass_WriteDB = True
  Else
    IBenefitClass_WriteDB = False
  End If
  
  Set ben = Nothing
  
End Function

Public Function Add(Loan As Loan) As Long
   Add = m_loans.Add(Loan)
End Function
Public Function Remove(Index As Long)
  Call m_loans.Remove(Index)
End Function
Private Function IsDeminimus() As Boolean
  Dim dDate As Date
  Dim lLastIndex As Long
  Dim lLoanBalanceIndexes() As Long
  Dim i As Long, Count As Long
  Dim Loan As Loan
  Dim ben As IBenefitClass
  Dim bProcessAllDatesInOrderMethod As Boolean
  Dim dates() As Date
  
  
  On Error GoTo IsDeminimus_Err
  
  Call xSet("IsDeminimus")
  
  If loans.Count = 0 Then GoTo IsDeminimus_End:
  
  ReDim lLoanBalanceIndexes(1 To loans.Count)
  
  IsDeminimus = True
  
  bProcessAllDatesInOrderMethod = True
  dDate = p11d32.Rates.value(TaxYearStart)
  
  For i = 1 To m_loans.Count
    Set Loan = m_loans(i)
    If Not (Loan Is Nothing) Then
       'cadloan
       Set ben = Loan
       If (ben.value(ln_OPRA_Ammount_Foregone_db) > 0) Then
        IsDeminimus = False
        Exit For
       End If
    End If
  Next i
   
  If (IsDeminimus) Then
    If (bProcessAllDatesInOrderMethod) Then
     If SumOfBalancesOnDate(dDate) > L_LOANDEMINIMUS Then
       IsDeminimus = False
     Else
       dates = BalanceSheetDatesOrdered(Count)
       If (Count > 0) Then
         For i = 1 To Count
           If SumOfBalancesOnDate(dates(i)) > L_LOANDEMINIMUS Then
             IsDeminimus = False
             Exit For
           End If
         Next
       End If
     End If
    Else
      Do
        If SumOfBalancesOnDate(dDate) > L_LOANDEMINIMUS Then
          IsDeminimus = False
          Exit Do
        End If
      Loop Until NextActionDate(dDate, dDate, lLoanBalanceIndexes()) = False
    End If
  End If
  
  
IsDeminimus_End:
  Call xReturn("IsDeminimus")
  Exit Function
IsDeminimus_Err:
  ErrorMessage ERR_ERROR, Err, "IsDeminimus", "ERR_IS_DEMINIMUS", "Error finding the de minimus nature of the loans collection " & "LOAN"
  Resume IsDeminimus_End
  Resume
End Function

Private Function BalanceSheetDatesOrdered(Count As Long) As Date()
  Dim l As Long
  Dim OL As ObjectList
  Set OL = New ObjectList
  Dim StaticData() As BEN_DATA_STATIC
  Dim dates() As Date
  Dim Loan As Loan
  Dim d As Date
  Dim v As Variant
  Dim bi As BalanceItem
  
On Error GoTo err_err
  Count = 0
  For l = 1 To m_loans.Count
    Set Loan = m_loans(l)
    If Not (Loan Is Nothing) Then
      Call Loan.BalanceSheetItemsAdd(OL)
    End If
  Next l
  If OL.Count > 0 Then
    ReDim dates(1 To OL.Count)
    For l = 1 To OL.Count
      Set bi = OL.Item(l)
      dates(l) = bi.DateFrom
    Next
    v = dates()
    Call QSortAsc(v, 1, OL.Count)
    Count = OL.Count
  End If
  BalanceSheetDatesOrdered = dates
  
err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "BalanceSheetDatesOrdered"), Err.Description)
  Resume err_end
  Resume
End Function


Private Function SumOfBalancesOnDate(ByVal dDate As Date) As Double
  Dim Loan As Loan
  Dim l As Long
  Dim lLoanBalanceIndexes() As Long
  Dim ben As IBenefitClass
  Dim benEY As IBenefitClass
  
  On Error GoTo lSumBalance_Err
  
  Call xSet("lSumBalance")
  Set benEY = p11d32.CurrentEmployer
  
  ReDim lLoanBalanceIndexes(1 To loans.Count)
  If dDate <> UNDATED Then
    For l = 1 To m_loans.Count
      Set Loan = m_loans(l)
      If Not (Loan Is Nothing) Then
         'cadloan
         Set ben = Loan
         If (Not (ben.value(ln_CheapTaxableLoan_db))) Or benEY.value(employer_CheapTaxableLoansAllNon_db) Then
           'If (ben.value(ITEM_BENEFIT) <= 0) Then
           If Not ben.value(ITEM_BENEFIT_REPORTABLE) Then
             GoTo NEXT_ITEM
           End If
         Else
           If (ben.value(ITEM_BENEFIT) <= 0) And (ben.value(ln_InterestPaid_db) > 0) And (ben.value(ln_MaxOutstandingAtAnyPoint) > 0) Then  'this feels wrong, we are assuming that inerest paid made the loan benefit 0, also use max outstanding to determine if there was some positive balance
             GoTo NEXT_ITEM
           End If
         End If
        SumOfBalancesOnDate = SumOfBalancesOnDate + Loan.BalanceOnDate(dDate, lLoanBalanceIndexes(l))
      End If
NEXT_ITEM:
    Next l
  End If
lSumBalance_End:
  Set Loan = Nothing
  Call xReturn("SumOfBalancesOnDate")
  Exit Function
lSumBalance_Err:
  ErrorMessage ERR_ERROR, Err, "SumOfBalancesOnDate", "ERR_SUMBALANCE", "Error summing the balances of loans on " & Format(dDate, "dd/mmmm/yyyy")
  Resume lSumBalance_End
  Resume
End Function
Private Function NextActionDate(dNextDate As Date, ByVal dStartDate As Date, lLoanBalanceIndexes() As Long) As Boolean
  Dim Loan As Loan
  Dim dMaxDate As Date
  Dim l As Long
  Dim bFound As Boolean
  Dim dNextDateLocal As Date
  On Error GoTo NextActionDate_Err
  
  Call xSet("NextActionDate")
  
  dMaxDate = p11d32.Rates.value(NextTaxYearStart)
  dNextDateLocal = UNDATED
  
  For l = 1 To m_loans.Count
    Set Loan = m_loans(l)
    If Not (Loan Is Nothing) Then
      lLoanBalanceIndexes(l) = lLoanBalanceIndexes(l) + 1
      If Loan.NextActionDate(dNextDateLocal, dStartDate, lLoanBalanceIndexes(l)) Then
        If dNextDateLocal < dMaxDate And dNextDateLocal > dStartDate Then
          If (dNextDateLocal <> UNDATED) Then
            If (dNextDate < dNextDateLocal) Then
              dNextDate = dNextDateLocal
            End If
          Else
            dNextDate = dNextDateLocal
          End If
          'dMaxDate = dNextDate
          bFound = True
        End If
      End If
    End If
  Next
  
  NextActionDate = bFound
  
NextActionDate_End:
  Call xReturn("NextActionDate")
  Exit Function

NextActionDate_Err:
  Resume NextActionDate_End
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_BENLOAN
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  ECASE ("IBenefitClass_InvalidFields for loans collection")
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  Dim Loan As Loan
  Dim i As Long
  
  For i = 1 To m_loans.Count
    Set Loan = m_loans(i)
    If Not Loan Is Nothing Then
       IBenefitClass_InvalidFields = IBenefitClass_InvalidFields + Loan.InvalidFields
    End If
  Next i
  
  Set Loan = Nothing
End Property


Public Property Get Count() As Long
  Count = m_loans.Count
End Property
Public Property Get Item(Index As Long) As Loan
  Set Item = m_loans.Item(Index)
End Property



