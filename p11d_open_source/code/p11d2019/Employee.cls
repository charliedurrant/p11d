VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Employee"
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

'Employee class module
Public Enum EmployeeItems
  ee_Surname_db = ITEM_DESC 'bf
  ee_value = ITEM_VALUE
  ee_MadeGood_Net = ITEM_MADEGOOD_NET
  ee_Benefit = ITEM_BENEFIT
  ee_MadeGood = ITEM_MADEGOOD
  ee_ActualAmountMadeGood = ITEM_ACTUALAMOUNTMADEGOOD
  ee_Benefit_Reportable = ITEM_BENEFIT_REPORTABLE
  ee_UDM_BENEFIT_TITLE = ITEM_UDM_BENEFIT_TITLE
  ee_BoxNumber = ITEM_BOX_NUMBER
  ee_MadeGood_Is_TaxDeducted = ITEM_MADEGOOD_IS_TAXDEDUCTED
  ee_Class1A_Adjustment = ITEM_CLASS1A_ADJUSTMENT
  ee_NIC_Class1A_Able = ITEM_NIC_CLASS1A_ABLE
  ee_NIC_BENEFIT_SUBJECT_TO_CLASS1A = ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  ee_NIC_Class1A_Benefit = ITEM_NIC_CLASS1A_BENEFIT
  ee_Error = ITEM_ERROR
  
  ee_Selected
  ee_Firstname_db
  ee_Email_db
  ee_Salutation_db
  ee_Title_db
  ee_Initials_db
  ee_FullName
  ee_PersonnelNumber_db
  'ee_filename  RK Redundant? 19/06/03
  ee_Username_db
  ee_Password_db
  'ee_IntranetLogin RK Redundant? 19/06/03
  ee_Group1_db
  ee_Group2_db
  ee_Group3_db
  ee_Comments_db
  ee_NINumber_db
  ee_joined_db
  
  ee_left_db
  ee_Gender_db
  ee_DOB_db
  ee_Director_db
'MP DB (not used)  ee_CheckAddress
  ee_AddressLine1_db
  [_EE_ADDRESS_DETAILS_FIRST_ITEM] = ee_AddressLine1_db
  ee_AddressLine2_db
  ee_AddressLine3_db
  ee_City_db
  ee_County_db
  ee_PostCode_db
  ee_Country_db
  [_EE_ADDRESS_DETAILS_LAST_ITEM] = ee_Country_db
  
'MP DB (not used)  ee_TotalEmployeeCarMiles 'backwards
'MP DB (not used)  ee_TotalEmployeeCarAllowance 'backwards
  ee_PaymentsForPrivateUseOfSharedVans_db
  ee_VansOPRAAmountForegone_db
  ee_VansOPRAAmmountForegoneUsedForValue
  ee_VansFuelOPRAAmountForegone_db
  ee_VansFuelOPRAAmmountForegoneUsedForValue
  ee_OneOrMoreSharedVanAvailable_db
  ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db
  ee_ReportyDailyCalculationOfSharedVans_db
  ee_Class1AEmployeeIsNotSubjectTo_db 'FC - Class1A
  ee_RelevantDaysForDailySharedVanCalc_db
  ee_NIInvalid
  ee_NIC_AdjustmentDeduct
  ee_TotalBenefitsPotentiallySubjectToClass1A
  ee_LASTITEM = ee_TotalBenefitsPotentiallySubjectToClass1A

End Enum

Public benefits As ObjectList

Private m_Locked As Boolean
Private m_dateedit_db As Date
Private m_dirty As Boolean

' inititalized at -1, when loaded contains no of bens.
Private m_initOK As Boolean
Private m_Parent As Employer
Private m_InvalidFields As Long
Private m_InvalidVanFields As Long
Private m_bCalculatingDependantEmployeeCars As Boolean

Private m_BenItems(1 To ee_LASTITEM)

Private m_BenClass As BEN_CLASS
Private m_ReadFromDB As Boolean
Private m_bBenefitsLoaded As Boolean
Private m_bHasAddress As Boolean
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Employee " & m_BenItems(ee_Surname_db) & " - terminate"
End Sub

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Employees"
End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = m_NeedToCalculate
End Property

Private Function IBenefitClass_CanBringForward() As Boolean
  Dim ben As IBenefitClass
  
  Set ben = Me
  IBenefitClass_CanBringForward = (ben.value(ee_left_db) = UNDATED)
  
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim ee As IBenefitClass
  
  On Error GoTo Copy_END
  Call xSet("Copy")
  
  Set ee = New Employee
  If CopyBenData(ee, Me) Then
    Set ee.Parent = Parent
    Call CopyBenStandardEnd(ee)
    Set IBenefitClass_Copy = ee
  End If
  
Copy_END:
  Call xReturn("Copy")
  Exit Function
Copy_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Copy", "Copy", "Error copying an employee.")
  Resume Copy_ERR
End Function
'cad cdb stuff
Public Function HasCDBBenefit(benCDB As IBenefitClass) As Long
  Dim j As Long
  Dim benEE As IBenefitClass
  Dim benOther As other
  
  On Error GoTo err_err
  
  For j = 1 To benefits.Count
    Set benEE = benefits(j)
    If Not benEE Is Nothing Then
      If benEE.CompanyDefined And (TypeOf benEE Is other) Then
        Set benOther = benEE
        If Not benOther.CDBMasterBenefitLink Is Nothing Then
          If benOther.CDBMasterBenefitLink Is benCDB Then
            HasCDBBenefit = j
            Exit For
          End If
        End If
      End If
    End If
  Next
  
err_end:
  Exit Function
err_err:
   Call Err.Raise(Err.Number, ErrorSource(Err, "HasCDBBenefit"), Err.Description)
   Resume
End Function

Public Property Get HasBenefit(ByVal bc As BEN_CLASS, Optional bIgnoreBenefitsLoaded As Boolean = False) As Boolean
  Dim i As Long
  Dim ben As IBenefitClass
  
  If (Not Me.BenefitsLoaded) And Not bIgnoreBenefitsLoaded Then Call Err.Raise(ERR_BENEFITS, "HasBenefit", "Benefits are not yet loaded")
  If (bc = BC_LOAN_OTHER_H) Or bc = BC_ALL Then
    HasBenefit = AnyLoanBenefit
    If (bc = BC_ALL) Then
      If HasBenefit Then
        Exit Property
      Else
        GoTo van:
      End If
    Else
      Exit Property
    End If
  ElseIf bc = BC_nonSHAREDVAN_G Or bc = BC_SHAREDVAN_G Then
van:
    HasBenefit = Me.AnyVanBenefit
    If (bc = BC_ALL) Then
      If HasBenefit Then
        Exit Property
      Else
        GoTo other
      End If
    Else
      Exit Property
    End If
  Else
other:
    For i = 1 To benefits.Count
      Set ben = benefits(i)
      If (Not ben Is Nothing) Then
        If ben.BenefitClass <> BC_SHAREDVANs_G And ben.BenefitClass <> BC_NONSHAREDVANS_G And ben.BenefitClass <> BC_LOANS_H Then
          If ben.BenefitClass = bc Or bc = BC_ALL Then
            HasBenefit = True
            Exit Property
          End If
        End If
      End If
    Next
  End If
End Property

Public Property Get HasAddress() As Boolean
  HasAddress = m_bHasAddress
End Property
Public Property Let HasAddress(NewValue As Boolean)
  m_bHasAddress = NewValue
End Property
Public Property Get CalculatingDependantEmployeeCars() As Boolean
  CalculatingDependantEmployeeCars = m_bCalculatingDependantEmployeeCars
End Property
Public Property Let CalculatingDependantEmployeeCars(NewValue As Boolean)
  m_bCalculatingDependantEmployeeCars = NewValue
End Property
Public Function GetLoan(ByVal Index As Long) As IBenefitClass
  Dim loans As loans
  Dim lLoansIndex As Long

  On Error GoTo GetLoan_ERR
  
  Call xSet("GetLoan")
  
  If AnyLoanBenefit Then
    lLoansIndex = GetLoansBenefitIndex()
    If lLoansIndex <= benefits.Count Then
      Set loans = benefits(lLoansIndex)
      If loans Is Nothing Then
        Call Err.Raise(ERR_IS_NOTHING, "GetLoan", "The loans collection for " & FullName & " is nothing but requested loan.")
      Else
        Set GetLoan = loans.loans(Index)
      End If
    End If
  End If
  
GetLoan_END:
  Call xReturn("GetLoan")
  Exit Function
GetLoan_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetLoan", "Get Loan", "Error getting loan, loan index = " & Index)
  Resume GetLoan_END
End Function
Public Function RemoveBenefit(ibf As IBenefitForm2, ben As IBenefitClass, ByVal BenefitIndex As Long, Optional bToScreen As Boolean = True) As Boolean
  Dim NextBenefitIndex As Long
  Dim loans As loans
  Dim li As ListItem
  
  On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  'need special treatement for loans as they are collections within collections
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "RemoveBenefit", "The benefit is nothing.")
  ben.NeedToCalculate = True
  Call NeedToCalculatePropogate(ben)
  If Not ben.CompanyDefined Or IsCBDEmployee(PersonnelNumber) Then
    'need special treatement for loans as they are a collection within a collection
    Call ben.DeleteDB
    'get the nextbest listitem tag
    If bToScreen Then NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
    If BenefitIsLoan(ben.BenefitClass) Then
      'then my index is for the loans collection and not the benefits collection
      Set loans = ben.Parent
      ' check to see if my parent, the loans collection, has any items apart from me
      Call loans.Remove(BenefitIndex)
      If loans.LoansCount = 0 Then
        'get rid of the loans collection in the benefits collection
        Call benefits.Remove(benefits.ItemIndex(loans))
      End If
    Else
      Call benefits.Remove(BenefitIndex)
    End If
    If bToScreen Then
      For Each li In ibf.lv.listitems
        If li.Tag = BenefitIndex Then
          ibf.lv.listitems.Remove (li.Index)
          Exit For
        End If
      Next
      'need to calc as some are dependant on others
      Call CalculateBenefitsOfType(ibf.benclass)
      Call SelectBenefitByBenefitIndex(ibf, NextBenefitIndex)
    End If
    RemoveBenefit = True
  End If
    
RemoveBenefit_END:
  Set ibf = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
End Function
Public Function RemoveBenefitWithLinks(ibf As IBenefitForm2, ben As IBenefitClass, ByVal BenefitIndex As Long, benLinked As IBenefitClass, Optional bToScreen As Boolean = True) As Boolean
  
  On Error GoTo RemoveBenefit_ERR
  
  Call xSet("RemoveBenefit")
  
  Call benefits.Remove(benefits.ItemIndex(benLinked))
  
  RemoveBenefitWithLinks = RemoveBenefit(ibf, ben, BenefitIndex, bToScreen)
  
RemoveBenefit_END:
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  RemoveBenefitWithLinks = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing the selected benefit.")
  Resume RemoveBenefit_END
  Resume
End Function
Public Sub CalculateBenefitsOfType(ByVal bc As BEN_CLASS)
  Dim i As Long
  Dim ben As IBenefitClass
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = bc Then
        ben.NeedToCalculate = True
        Call ben.Calculate
        
      End If
    End If
  Next
End Sub

Public Function IterateBenefits(IT As ITERATION_TYPE) As Boolean
  Dim l As Long
  Dim ben As IBenefitClass
  Dim lDirtyBenefits As Long, lBenefitsWritten As Long
  
On Error GoTo IterateBenefits_ERR
  
  Call xSet("IterateBenefits")
  Set ben = Me
  Call ben.Calculate
     
  For l = 1 To benefits.Count
    Set ben = benefits.Item(l)
    If Not (ben Is Nothing) Then
      Select Case IT
        Case ITERATION_TYPE.AnyDirty
          If ben.Dirty Then
            IterateBenefits = True
            Exit For
          End If
        Case ITERATION_TYPE.AnyReportable
          If ben.value(ITEM_BENEFIT_REPORTABLE) Then
            IterateBenefits = True
            Exit For
          End If
        Case ITERATION_TYPE.AnyInvalid
          If ben.InvalidFields > 0 Or (StrComp(ben.value(ITEM_BENEFIT), S_ERROR) = 0) Then
            IterateBenefits = True
            Exit For
          End If
        Case ITERATION_TYPE.WriteDirtyAndAreAllWritten
          If ben.Dirty Then
            lDirtyBenefits = lDirtyBenefits + 1
            If ben.InvalidFields = 0 Or IsBenefitACollection(ben) Then
              lBenefitsWritten = lBenefitsWritten + Abs(ben.writeDB)
            End If
          End If
          
      End Select
    End If
  Next l
  
  If IT = WriteDirtyAndAreAllWritten Then
    IterateBenefits = IIf((lDirtyBenefits - lBenefitsWritten) > 0, False, True)
  End If

  
IterateBenefits_END:
  Call xReturn("IterateBenefits")
  Exit Function
IterateBenefits_ERR:
  IterateBenefits = False
  Call ErrorMessage(ERR_ERROR, Err, "IterateBenefits", "Iterate Current Employee Benefits", "Error iterating the current employees benefits, iteration type of " & IT & ".")
  Resume IterateBenefits_END
  Resume
End Function

Private Sub Class_Initialize()
  Dim i As Long
  Dim ben As IBenefitClass
  
  On Error GoTo Employee_Initialize_err
  Call xSet("Employee_Initialize")
  
  Set ben = Me

  Call IBenefitClass_SetBenItemsInformation
  Set benefits = New ObjectList
  m_initOK = True
  
Employee_Initialize_end:
  Call xReturn("Employee_Initialize")
  Exit Sub
  
Employee_Initialize_err:
  m_initOK = False
  Resume Employee_Initialize_end
End Sub

Public Sub KillBenefitsEx(Optional ByVal CurrentEmployee As Employee = Nothing)
  If Not CurrentEmployee Is Nothing Then
    If (Not CurrentEmployee Is Me) And p11d32.KillBenefits Then Call KillBenefits
  Else
    If (Not p11d32.CurrentEmployer.CurrentEmployee Is Me) And p11d32.KillBenefits Then Call KillBenefits
  End If
  Call DoDBEvents(FREE_LOCKS)
End Sub

Public Function KillBenefits() As Boolean
  Dim i As Long
  Dim ben As IBenefitClass
  
  Call xSet("KillBenefits")
  
 'Debug.Print"Employee " & m_BenItems(ee_Surname) & " - kill benefits"
  
  If Not benefits Is Nothing Then
    For i = 1 To benefits.Count
      If Not (benefits(i) Is Nothing) Then
        Set ben = benefits(i)
        Call ben.Kill
        Set ben = Nothing
      End If
    Next i
  End If
  
  benefits.RemoveAll
  BenefitsLoaded = False
  
  Call xReturn("KillBenefits")
  
End Function
Public Property Get InitOK() As Boolean
  InitOK = m_initOK
End Property

Public Function LoadBenefits(ben As BENEFIT_TABLES, Optional bToScreen As Boolean = True, Optional bForceLoad As Boolean = False, Optional ByVal bForceLoaded As Boolean = False) As Boolean
  Dim benobj As IBenefitClass
  Dim reloc As Relocation
  Dim i As Long
  Dim iFirstItem As Long, iLastItem As Long
  On Error GoTo LoadBenefits_Err
  
  Call xSet("LoadBenefits")
  
  If m_bBenefitsLoaded And Not bForceLoad Then GoTo LoadBenefits_End
  
  Call KillBenefits
  
  If (ben = TBL_ALLBENEFITS) Then
    iFirstItem = TBL_FIRST_ITEM
    iLastItem = TBL_LASTITEM
  Else
    iFirstItem = ben
    iLastItem = ben
  End If
  
  
  If bToScreen Then Call PrgStart(iLastItem)
  
  For i = iFirstItem To iLastItem
    Set benobj = Nothing
    Select Case i
      Case TBL_VANS
        Set benobj = New nonSharedVans
      Case TBL_COMPANY_CARS
        Set benobj = New Fuel
        Set benobj = New CompanyCar
        
      Case TBL_SHARED_VANS
        'nothing as in employer level
      Case TBL_OTHER
        Set benobj = New other
        If Me Is p11d32.CurrentEmployer.CDBEmployee Then
          benobj.CompanyDefined = True
        End If
      Case TBL_CDB_LINKS
        'nothing as used in CreateCDBLinks
      Case TBL_BENLOAN
        Set benobj = New loans
      Case TBL_CDB_LINKS
      Case TBL_PHONES
        Set benobj = New phone
      Case TBL_ASSETSATDISPOSAL
        Set benobj = New AssetsAtDisposal
      Case TBL_SERVICESPROVIDED
        Set benobj = New ServicesProvided
      Case TBL_ACCOM
        Set benobj = New Accommodation
      Case TBL_EMPLOYEE_CARS
        Set benobj = New EmployeeCar
      Case TBL_RELOC
        Set benobj = New Relocation
      Case TBL_TRANS
        Set benobj = New AssetsTransferred
    End Select
    
    If Not benobj Is Nothing Then
      Set benobj.Parent = Me
      'this add then possible remove is to keep the ordering of loans and cars correct
      Call benefits.Add(benobj)
      If benobj.ReadDB = 0 Then Call benefits.Remove(benobj)
    End If
    If bToScreen Then Call PrgStep
  Next i
  
  'clear up
  For i = 1 To benefits.Count
    If Not benefits(i) Is Nothing Then
      Set benobj = benefits(i)
      Call benobj.ReadDB
    End If
  Next
  
  
  Call CreateCDBLinks
  If (ben = TBL_ALLBENEFITS) Or bForceLoaded Then
    m_bBenefitsLoaded = True
  End If
  LoadBenefits = True
  
LoadBenefits_End:
  If bToScreen Then Call PrgStop
  Call DBEngine.Idle(dbFreeLocks)
  Call xReturn("LoadBenefits")
  Exit Function

LoadBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadBenefits", "Load Benefit", "Unable to load the benefit requested")
  Resume LoadBenefits_End
  Resume
End Function
Public Property Get BenefitsLoaded() As Boolean
  BenefitsLoaded = m_bBenefitsLoaded
End Property
Public Property Let BenefitsLoaded(NewValue As Boolean)
  m_bBenefitsLoaded = NewValue
End Property
Public Property Get InvalidVanFields() As Long
  InvalidVanFields = m_InvalidVanFields
End Property

Public Property Let InvalidVanFields(ByVal l As Long)
  m_InvalidVanFields = l
End Property
Public Function ReportFooter(Description As String) As String
  ReportFooter = "{Arial=8,nb}{x=0}P11D " & p11d32.Rates.value(TaxFormYear) & Description & "(" & app.companyName & "){Arial=8,i}{x=50}Page {PAGE}" & " -  printed at {TIME} on {DATE}"
End Function
Public Function SumBenefit(ByRef Description As Variant, ByRef value As Variant, ByRef MadeGood As Variant, ByRef benefit As Variant, BenArr() As BEN_CLASS, Optional VALUE_ENUM As Long = ITEM_VALUE, Optional MADEGOOD_ENUM As Long = ITEM_MADEGOOD_NET, Optional BENEFIT_ENUM As Long = ITEM_BENEFIT, Optional sIRDesc As String) As Boolean
  Dim ben As IBenefitClass
  Dim i As Long, j As Long, lLoansIndex As Long, k As Long
  Dim loans As loans
  Dim bFoundBen As Boolean
  Dim OL As ObjectList
  
  On Error GoTo SumBenefit_Err
  Call xSet("SumBenefit")
  
  value = 0
  MadeGood = 0
  benefit = 0
  Description = ""
    
    
  For i = 1 To UBound(BenArr)
    If BenArr(i) = BC_NONSHAREDVANS_G Then
      If Not AnyVanBenefit Then GoTo NEXT_ARRAY_BEN
    End If
    If BenefitIsLoan(BenArr(i)) Then
      lLoansIndex = GetLoansBenefitIndex
      If lLoansIndex < 1 Then ECASE ("Loans not found when loan type passed in Sum Benefit.")
      Set loans = benefits(lLoansIndex)
      If loans Is Nothing Then ECASE ("Loans not found when loan type passed in Sum Benefit.")
      Set OL = loans.loans
    Else
      Set OL = benefits
    End If
    
    For j = 1 To OL.Count
      Set ben = OL(j)
      If Not ben Is Nothing Then
        If ben.BenefitClass = BenArr(i) Then
          If ben.LinkBen And (Not ben.value(ITEM_BENEFIT_REPORTABLE)) Then
            GoTo NEXT_BEN
          End If
          value = value + ben.value(VALUE_ENUM)
          MadeGood = MadeGood + ben.value(MADEGOOD_ENUM)
          benefit = benefit + ben.value(BENEFIT_ENUM)
          k = k + 1
          bFoundBen = True
          If BenefitIsLoan(BenArr(i)) Then
            Description = p11d32.Rates.BenClassTo(BenArr(i), BCT_FORM_CAPTION)
          Else
            If k > 1 Then
              Description = "Multiple items"
              sIRDesc = "Multiple items"
            Else
              Description = ben.value(ITEM_DESC) & " "
              'CAD CHECK
              If (HasIRDescription(ben.BenefitClass)) Then
                sIRDesc = ben.value(IRDescriptionBenItem(ben.BenefitClass))
              End If
            End If
          End If
        End If
      End If
NEXT_BEN:
    Next
NEXT_ARRAY_BEN:
  Next
  
  SumBenefit = bFoundBen
  If Not SumBenefit Or benefit < 0 Then 'EK Or added to filter out neg section E benefit
    value = ""
    MadeGood = ""
    benefit = ""
  End If
  
SumBenefit_End:
  Call xReturn("SumBenefit")
  Exit Function
  
SumBenefit_Err:
  SumBenefit = False
  Description = S_ERROR
  value = S_ERROR
  MadeGood = S_ERROR
  benefit = S_ERROR
  Resume SumBenefit_End
  Resume
End Function
Private Function AddRelevantBenClass(ByVal bc As BEN_CLASS, BenClasses() As BEN_CLASS, lUboundBenClasses As Long)
  Dim j As Long
  
  On Error GoTo AddRelevantBenClass_ERR
  Call xSet("AddRelevantBenClass")
    
  For j = 1 To lUboundBenClasses
    If bc = BenClasses(j) Then
      AddRelevantBenClass = True
      Exit For
    End If
  Next
  
  If Not AddRelevantBenClass Then
    lUboundBenClasses = lUboundBenClasses + 1
    ReDim Preserve BenClasses(1 To lUboundBenClasses)
    BenClasses(lUboundBenClasses) = bc
  Else
    AddRelevantBenClass = False
  End If
  
AddRelevantBenClass_END:
  Call xSet("AddRelevantBenClass")
  Exit Function
AddRelevantBenClass_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddRelevantBenClass", "Add Relevant Ben Class", "Error adding a relevant ben class.")
  Resume AddRelevantBenClass_END
End Function
Public Function GetRelevantBenClasses(BenClasses() As BEN_CLASS) As Long
  Dim ben As IBenefitClass
  Dim i As Long, j As Long
  Dim bAdded As Boolean
  
  On Error GoTo GetRelevantBenClasses_ERR
  
  Call xSet("GetRelevantBenClasses")
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If (2 ^ (p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION)) And p11d32.ReportPrint.HMITSections_PRINT) Then
        'special for loans
        If ben.BenefitClass = BC_LOANS_H Then
          'Call AddRelevantBenClass(BC_LOAN_HOME_H, BenClasses, j) 'RK redundant 18/03/03
          Call AddRelevantBenClass(BC_LOAN_OTHER_H, BenClasses, j)
        ElseIf ben.BenefitClass = BC_NONSHAREDVANS_G Then
          If AnyVanBenefit Then Call AddRelevantBenClass(ben.BenefitClass, BenClasses, j)
        Else
          Call AddRelevantBenClass(ben.BenefitClass, BenClasses, j)
        End If
      End If
    End If
  Next
  
  GetRelevantBenClasses = j
  
GetRelevantBenClasses_END:
  Call xReturn("GetRelevantBenClasses")
  Exit Function
GetRelevantBenClasses_ERR:
  GetRelevantBenClasses = 0
  Call ErrorMessage(ERR_ERROR, Err, "GetRelevantBenClasses", "Get Relevant Ben Classes", "Error getting the relevant benefit classes for employee " & FullName & ".")
  Resume GetRelevantBenClasses_END
  Resume
End Function

Public Function GetRelevantBenClassesSorted(RelevantBenClasses() As BEN_CLASS) As Long

  GetRelevantBenClassesSorted = GetRelevantBenClasses(RelevantBenClasses)
          
  If GetRelevantBenClassesSorted > 0 Then
    Call QSortEx(RelevantBenClasses, 1, GetRelevantBenClassesSorted, Me)
  End If
  
End Function
Public Function CalculateBenefits(Optional lRelevantHMITSections As Long = -1, Optional ByVal HSC As HMIT_SELECTION_CHOICE = HMIT_SC_ALL) As Boolean
  Dim i As Long
  Dim ben As IBenefitClass
  Dim bNonzero As Boolean
  Dim HMITS As HMIT_SECTIONS
  
  On Error GoTo CalculateBenefits_ERR
  Call xSet("CalculateBenefits")
  
  CalculateBenefits = 0
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      ben.Calculate
      HMITS = 2 ^ p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION)
      If ben.value(ITEM_BENEFIT_REPORTABLE) = True Then lRelevantHMITSections = lRelevantHMITSections Or HMITS
      
      If Not bNonzero Then
        Select Case HSC
          Case HMIT_SC_ALL
            If -1 And HMITS Then bNonzero = ben.value(ITEM_VALUE) > 0
          Case HMIT_SC_RELEVANT
            If lRelevantHMITSections And HMITS Then bNonzero = ben.value(ITEM_VALUE) > 0
            'I know it is as above?
          Case HMIT_SC_SELECTED
            If p11d32.ReportPrint.HMITSections And HMITS Then bNonzero = ben.value(ITEM_VALUE) > 0
        End Select
      End If
    End If
  Next
  
  CalculateBenefits = bNonzero
  
CalculateBenefits_END:
  Call xReturn("CalculateBenefits")
  Exit Function
CalculateBenefits_ERR:
  
  Call ErrorMessage(ERR_ERROR, Err, "CalculateBenefits", "Calculate Benefits", "Error calculating all the benefits for " & FullName & ".")
  Resume CalculateBenefits_END
  Resume
End Function
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_EMPLOYEE
End Property
Private Function IBenefitClass_CalculateBody() As Variant
  Dim i As Long, j As Long
  Dim ben As IBenefitClass
  Dim benEmployee As IBenefitClass
  Dim benefit As Long
  Dim bc As BEN_CLASS
  Dim OL As ObjectList
  Dim loans As loans
  Dim NIC As Double
  
  On Error GoTo Employee_Calculate_ERR
  Call xSet("Employee_Calculate")
  
  Set benEmployee = Me
  Call benEmployee.SetCalcDefaults
  benEmployee.value(ee_NIInvalid) = ValidateNI(benEmployee.value(ee_NINumber_db), True) = INVALID
  Call LoadBenefits(TBL_ALLBENEFITS, False)
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = BC_LOANS_H Then
        Set loans = ben
        Set OL = loans.loans
        For j = 1 To OL.Count
          Call Employee_CalculateEx(benEmployee, OL(j))
        Next
      Else
        Call Employee_CalculateEx(benEmployee, ben)
      End If
    End If
  Next
  
  Dim a As Long
  
  a = benEmployee.value(ee_NIC_BENEFIT_SUBJECT_TO_CLASS1A) - benEmployee.value(ITEM_CLASS1A_ADJUSTMENT)
  a = benEmployee.value(ee_TotalBenefitsPotentiallySubjectToClass1A) - a
  benEmployee.value(ee_NIC_AdjustmentDeduct) = a
  
Employee_Calculate_END:
  IBenefitClass_CalculateBody = benEmployee.value(ITEM_BENEFIT)
  Call xSet("Employee_Calculate")
  Exit Function
Employee_Calculate_ERR:
  benEmployee.value(ITEM_BENEFIT) = S_ERROR
  Resume Next
End Function
Private Function IBenefitClass_Calculate() As Variant
  IBenefitClass_Calculate = CalculateHelper(Me)
End Function
'FIX XX
Private Function Employee_CalculateEx(ByVal benEmployee As IBenefitClass, ByVal ben As IBenefitClass)
  Dim bc As BEN_CLASS
  Dim vBenefit As Variant
  Dim vValue As Variant
  On Error GoTo Employee_CalculateEx_ERR
  Call xSet("Employee_CalculateEx")

  If Not ben Is Nothing Then
    bc = ben.BenefitClass
    vBenefit = ben.Calculate
    Employee_CalculateEx = vBenefit
    vValue = ben.value(ITEM_VALUE)
    benEmployee.value(ITEM_BENEFIT) = benEmployee.value(ITEM_BENEFIT) + Employee_CalculateEx
    benEmployee.value(ITEM_VALUE) = benEmployee.value(ITEM_VALUE) + vValue
    benEmployee.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A) = benEmployee.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A) + ben.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A)
    Debug.Print (benEmployee.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A))
       
    If ben.value(ITEM_NIC_CLASS1A_ABLE) Then
      If (ben.BenefitClass = BC_LIVING_ACCOMMODATION_D) Then
        vBenefit = ben.value(accom_temp_nic_calc_benefit)
      End If
      benEmployee.value(ee_TotalBenefitsPotentiallySubjectToClass1A) = benEmployee.value(ee_TotalBenefitsPotentiallySubjectToClass1A) + vBenefit
      'this is the benefit that is subject ot 12.5%
      benEmployee.value(ITEM_NIC_CLASS1A_BENEFIT) = benEmployee.value(ITEM_NIC_CLASS1A_BENEFIT) + ben.value(ITEM_NIC_CLASS1A_BENEFIT)
      benEmployee.value(ITEM_CLASS1A_ADJUSTMENT) = benEmployee.value(ITEM_CLASS1A_ADJUSTMENT) + ben.value(ITEM_CLASS1A_ADJUSTMENT)
    End If
  End If
   
   
Employee_CalculateEx_END:
  Call xSet("Employee_CalculateEx")
  Exit Function
Employee_CalculateEx_ERR:
  benEmployee.value(ITEM_BENEFIT) = S_ERROR
  Resume Next
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE ("CompanyDefined")
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
  
End Property
Public Property Get DateEdit() As Date
  DateEdit = m_dateedit_db
End Property
Public Property Let DateEdit(NewValue As Date)
  m_dateedit_db = NewValue
End Property

Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo DeleteDB_Err
  
  Call xSet("DeleteDB")
  
  Call SetPanel2("Deleting employee: " & FullName)
  
  Call LoadBenefits(TBL_ALLBENEFITS, False)
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If Not ben.CompanyDefined Then
        Call ben.DeleteDB
        Call ben.Kill
      End If
    End If
  Next i
  
  benefits.RemoveAll
  
  If m_BenItems(ee_Selected) Then
    Set ben = m_Parent
    ben.value(employer_NoOfSelectedEmployees) = ben.value(employer_NoOfSelectedEmployees) - 1
  End If
  
  Set rs = Me.Parent.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEE_READ, PersonnelNumber), dbOpenDynaset)
  If rs.BOF And rs.EOF Then GoTo DeleteDB_End
  rs.MoveFirst
  rs.Delete
  Set rs = Nothing
  'delete address
  Me.Parent.db.Execute (sql.Queries(DELETE_ADDRESS, PersonnelNumber))
  Me.Parent.db.Execute (sql.Queries(DELETE_CDB_LINKS_EMPLOYEE, PersonnelNumber))
  
'MP DB - data entered into tbl, but never used
'  Set rs = Me.Parent.db.OpenRecordset(sql.Queries(SELECT_AUDIT), dbOpenDynaset)
'  rs.AddNew
'  rs.Fields("Action").value = "Delete"
'  rs.Fields(S_FIELD_PERSONEL_NUMBER).value = Me.PersonnelNumber
'  rs.Fields("UserName").value = p11d32.UserName
'  rs.Fields("ActionDate").value = Now()
'  rs.Update

DeleteDB_End:
  Call SetPanel2("")
  Call xReturn("DeleteDB")
  Exit Function

DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DeleteDB", "Delete DB", "Error deleting the employee.")
  Resume DeleteDB_End
  Resume
End Function
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Property Let IBenefitClass_LinkBen(RHS As Boolean)

End Property

Private Property Get IBenefitClass_LinkBen() As Boolean

End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
  'ben.value(ee_OneOrMoreSharedVanAvailable_db) = False REMOVED MY AM
  ben.value(ITEM_BENEFIT) = 0
  ben.value(ITEM_NIC_CLASS1A_BENEFIT) = 0
  ben.value(ITEM_CLASS1A_ADJUSTMENT) = 0
  ben.value(ITEM_NIC_CLASS1A_ABLE) = False
  ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) = False
  ben.value(ITEM_ERROR) = ""
  ben.value(ee_TotalBenefitsPotentiallySubjectToClass1A) = 0
  ben.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A) = 0
  
End Sub

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  If Item = ee_FullName Then
    IBenefitClass_value = FullName
  Else
    IBenefitClass_value = m_BenItems(Item)
  End If
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, value As Variant)
  Dim benEmployer As IBenefitClass
  
  Select Case Item
    Case ee_OneOrMoreSharedVanAvailable_db
      Set benEmployer = m_Parent
      With p11d32.CurrentEmployer
        If value Then
          If m_BenItems(ee_OneOrMoreSharedVanAvailable_db) <> value Then
            benEmployer.value(employer_EmployeesWithSharedVan) = benEmployer.value(employer_EmployeesWithSharedVan) + 1
          End If
        Else
          If m_BenItems(ee_OneOrMoreSharedVanAvailable_db) <> value Then
            benEmployer.value(employer_EmployeesWithSharedVan) = benEmployer.value(employer_EmployeesWithSharedVan) - 1
          End If
        End If
        m_BenItems(ee_OneOrMoreSharedVanAvailable_db) = value
      End With
    Case ee_FullName
      Call ECASE("ee_FullName has no property let")
    Case Else
      m_BenItems(Item) = CorrectBenValue(BC_EMPLOYEE, Item, value)
  End Select
  
End Property
Private Function IsNamePart(ByVal Item As Long) As Boolean
  IsNamePart = (Item = ee_Title_db) Or (Item = ee_Surname_db) Or (Item = ee_Firstname_db) Or (Item = ee_Initials_db)
End Function
Public Property Get ForeNames() As String
  Dim ben As IBenefitClass
  Dim s As String, s1 As String
  
  Set ben = Me
  
  s = Trim$(ben.value(ee_Firstname_db))
  s1 = Trim$(ben.value(ee_Initials_db))

  s = s & " " & s1
  s = Trim(s)

  ForeNames = s
  

End Property

Public Property Get FullName() As String
  
  Dim s As String
  'construct the name with
  'Surname Title firstname Initials
  Select Case p11d32.NameOrder
      Case NO_SURNAME_TITLE_FN_INITIALS
        s = Trim(m_BenItems(ee_Surname_db))
        s = Trim$(s & " " & m_BenItems(ee_Title_db))
        s = Trim$(s & " " & m_BenItems(ee_Firstname_db))
        s = Trim$(s & " " & m_BenItems(ee_Initials_db))
      Case NO_SURNAME_INITIALS
        s = Trim(m_BenItems(ee_Surname_db))
        s = Trim$(s & " " & m_BenItems(ee_Initials_db))
      Case NO_TITLE_FN_INITIALS_SURNAME
        s = Trim(m_BenItems(ee_Title_db))
        s = Trim$(s & " " & m_BenItems(ee_Firstname_db))
        s = Trim$(s & " " & m_BenItems(ee_Initials_db))
        s = Trim$(s & " " & m_BenItems(ee_Surname_db))
      Case NO_FN_INITIALS_SURNAME
        s = Trim(m_BenItems(ee_Firstname_db))
        s = Trim$(s & " " & m_BenItems(ee_Initials_db))
        s = Trim$(s & " " & m_BenItems(ee_Surname_db))
      Case NO_FN_SURNAME
        s = Trim(m_BenItems(ee_Firstname_db))
        s = Trim$(s & " " & m_BenItems(ee_Surname_db))
      Case NO_SURNAME_FN_INITIALS
        s = Trim(m_BenItems(ee_Surname_db))
        s = Trim$(s & " " & m_BenItems(ee_Firstname_db))
        s = Trim$(s & " " & m_BenItems(ee_Initials_db))
      Case NO_SURNAME_FN
        s = Trim(m_BenItems(ee_Surname_db))
        s = Trim$(s & " " & m_BenItems(ee_Firstname_db))
      Case NO_TITLE_INITIALS_SURNAME
        s = Trim(m_BenItems(ee_Title_db))
        s = Trim$(s & " " & m_BenItems(ee_Initials_db))
        s = Trim$(s & " " & m_BenItems(ee_Surname_db))
      Case NO_INITIALS_SURNAME
        s = Trim(m_BenItems(ee_Initials_db))
        s = Trim$(s & " " & m_BenItems(ee_Surname_db))
  End Select
  
  FullName = s
  
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = m_ReadFromDB
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property

Private Sub IBenefitClass_Kill()
  Dim i As Long
  Dim ben As IBenefitClass
  
On Error GoTo Employee_Kill_Err

  Call xSet("Employee Kill")
  
  Call KillBenefits
  Set m_Parent = Nothing
  
Employee_Kill_End:
  Call xReturn("Employee Kill")
  Exit Sub
Employee_Kill_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Employee_Kill", "ERR_DELETE_EMPLOYEE", "Unable to remove the employee from memory.")
  Resume Employee_Kill_End

End Sub
Public Property Set Parent(NewValue As Employer)
  Set m_Parent = NewValue
End Property
Public Property Get Parent() As Employer
  Set Parent = m_Parent
End Property
Public Property Get InvalidFields() As Long
  InvalidFields = m_InvalidFields
End Property
Public Property Let InvalidFields(NewValue As Long)
  m_InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = FullName
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Property Get IBenefitClass_PrintHeader() As String
  ECASE ("PrintHeader")
End Property
Public Function PrintPrepare() As Boolean
  Dim HMITS As HMIT_SECTIONS
  Dim lRelevantHMITSections As Long
  On Error GoTo PrintPrepare_ERR
  
  Call xSet("PrintPrepare")
  
  Call LoadBenefits(TBL_ALLBENEFITS, False)
  
  HMITS = p11d32.ReportPrint.HMITSelectionChoice
  
  'if add to select case then add to calculate benefits
  PrintPrepare = CalculateBenefits(lRelevantHMITSections, HMITS)
  Select Case HMITS
    Case HMIT_SC_RELEVANT
      p11d32.ReportPrint.HMITSections_PRINT = lRelevantHMITSections
    Case HMIT_SC_ALL
      p11d32.ReportPrint.HMITSections_PRINT = -1
    Case HMIT_SC_SELECTED
      p11d32.ReportPrint.HMITSections_PRINT = p11d32.ReportPrint.HMITSections
  End Select
  
  
PrintPrepare_END:
  Call xReturn("PrintPrepare")
  Exit Function
PrintPrepare_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PrintPrepare", "Print Prepare", "Error preparing the emloyee " & FullName & " for printing.")
  Resume PrintPrepare_END
  Resume
End Function
Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Dim lRelevantHMITSections As Long
  
  On Error GoTo IBenefitClass_PrintWk_ERR
  
  Call xSet("IBenefitClass_PrintWk")
  
  Call PrintPrepare
  IBenefitClass_PrintWk = Report_HMIT(rep, Me)
  
IBenefitClass_PrintWk_END:
  Call xReturn("IBenefitClass_PrintWk")
  Exit Function
IBenefitClass_PrintWk_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Employee_HMIT_PrintWk", "Employee HMIT Print Wk", "Error printing the employee HMIT return")
  Resume IBenefitClass_PrintWk_END
End Function
Public Property Get PersonnelNumber() As String
  PersonnelNumber = m_BenItems(ee_PersonnelNumber_db)
End Property
Private Function IBenefitClass_ReadDB() As Long
  Dim ben As IBenefitClass
  Dim vancol As SharedVans
  Dim bSharedVan As Boolean
  Dim rs As Recordset
  
  On Error GoTo Employee_ReadDB_Err
  Call xSet("Employee ReadDB")
  
  If m_initOK Then
    Set ben = Me
    
    Set rs = p11d32.CurrentEmployer.EmployeesRecordSet
    
    If rs Is Nothing Then
      Set rs = p11d32.CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEE_READ, m_BenItems(ee_PersonnelNumber_db)), dbOpenDynaset, dbFailOnError)
    End If
    
    With rs
      ben.SetCalcDefaults     'MP DB ToDo - why SetCalcDefaults here in ReadDB? Also done in Calculate
      ben.value(ee_Surname_db) = "" & .Fields("Surname")
      ben.value(ee_Initials_db) = "" & .Fields("Initials")
      ben.value(ee_Firstname_db) = "" & .Fields("Firstname")
      ben.value(ee_Title_db) = "" & .Fields("Title")
      ben.value(ee_Salutation_db) = "" & .Fields("Salutation")
      
      ben.value(ee_Email_db) = "" & .Fields("Email")
      ben.value(ee_PersonnelNumber_db) = "" & .Fields(S_FIELD_PERSONEL_NUMBER)
      
      If p11d32.BringForward.Yes And p11d32.AppYear = 2000 Then
        ben.value(ee_Password_db) = ""
        ben.value(ee_Username_db) = ""
      Else
        ben.value(ee_Password_db) = IIf(IsNull(.Fields("EEPassword")), "", (.Fields("EEPassword")))
        ben.value(ee_Username_db) = IIf(IsNull(.Fields("EEUsername")), "", (.Fields("EEUsername")))
      End If
      ben.value(ee_Group1_db) = "" & .Fields("Group1")
      ben.value(ee_Group2_db) = "" & .Fields("Group2")
      ben.value(ee_Group3_db) = "" & .Fields("Group3")
      ben.value(ee_NINumber_db) = "" & .Fields("NI")
      
      ben.value(ee_joined_db) = IIf(IsNull(.Fields("Joined")), UNDATED, .Fields("Joined"))
      ben.value(ee_left_db) = IIf(IsNull(.Fields("Left")), UNDATED, .Fields("Left"))
      
      If p11d32.BringForward.Yes And p11d32.AppYear = 2006 Then
        ben.value(ee_DOB_db) = UNDATED
        ben.value(ee_Gender_db) = S_GENDER_NA
      Else
        If FieldPresent(rs.Fields, "DOB") Then
          ben.value(ee_DOB_db) = IIf(IsNull(.Fields("DOB")), UNDATED, (.Fields("DOB")))
          ben.value(ee_Gender_db) = IIf(IsNull(.Fields("Gender")), S_GENDER_NA, (.Fields("Gender")))
        End If
      End If
      
      ben.value(ee_Comments_db) = "" & .Fields("Comments")
      ben.value(ee_Director_db) = IIf(.Fields("Status") = "Director", True, False)
      'van details
      ben.value(ee_RelevantDaysForDailySharedVanCalc_db) = .Fields("RELEVANTDAYS")
      ben.value(ee_PaymentsForPrivateUseOfSharedVans_db) = .Fields("PRIVATECONTRIB")
      'AM Fix (set to "SharedVans") below) ben.value(ee_OneOrMoreSharedVanAvailable) = .Fields("TWOPLUSVANS")
      ben.value(ee_ReportyDailyCalculationOfSharedVans_db) = .Fields("DAILYCALC")
      ben.value(ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db) = .Fields("TwoPlusVans")
      
      'address details
      'read from joined separate table
      m_bHasAddress = Not (IsNull(.Fields("Address1")) And IsNull(.Fields("Address2")) And IsNull(.Fields("Address3")) And IsNull(.Fields("City")) And IsNull(.Fields("County")) And IsNull(.Fields("PostCode")) And IsNull(.Fields("Country")))
      ben.value(ee_AddressLine1_db) = "" & .Fields("Address1")
      ben.value(ee_AddressLine2_db) = "" & .Fields("Address2")
      ben.value(ee_AddressLine3_db) = "" & .Fields("Address3")
      ben.value(ee_City_db) = "" & .Fields("City")
      ben.value(ee_County_db) = "" & .Fields("County")
      ben.value(ee_PostCode_db) = "" & .Fields("PostCode")
      ben.value(ee_Country_db) = "" & .Fields("Country")
      
      m_dateedit_db = IIf(IsNull(.Fields("DateEdit")), UNDATED, .Fields("DateEdit"))
      'need to initialise set employee Value
'MP DB ToDo why setting ee_OneOrMoreSharedVanAvailable twice below?
      m_BenItems(ee_OneOrMoreSharedVanAvailable_db) = False
      ben.value(ee_OneOrMoreSharedVanAvailable_db) = .Fields("SHAREDVANS")
      
      If (p11d32.BringForward.Yes) Then
        ben.value(ee_VansOPRAAmountForegone_db) = 0
        ben.value(ee_VansFuelOPRAAmountForegone_db) = 0
      Else
        ben.value(ee_VansOPRAAmountForegone_db) = IIf(IsNull(.Fields("VansOPRAAmountForegone")), 0, (.Fields("VansOPRAAmountForegone")))
        ben.value(ee_VansFuelOPRAAmountForegone_db) = IIf(IsNull(.Fields("VansFuelOPRAAmountForegone")), 0, (.Fields("VansFuelOPRAAmountForegone")))
      End If
      
      ben.BenefitClass = BC_EMPLOYEE
      ben.ReadFromDB = True
      
      'KM - Added if statement, as this field doesn't exist in P11D2000 (blnk00.mdb)
'      If p11d32.AppYear > 2000 Then
'AM Removed as should bring forward
'        If p11d32.BringForward.Yes Then
'          ben.value(ee_Class1AEmployeeIsNotSubjectTo) = False
'        Else
          ben.value(ee_Class1AEmployeeIsNotSubjectTo_db) = .Fields("Class1AEmployeeIsNotSubjectTo")
'        End If
'      End If
      
    End With
    Set ben = Nothing
  End If
  
Employee_ReadDB_End:
  Set rs = Nothing
  Call xReturn("Employee ReadDB")
  Exit Function
  
Employee_ReadDB_Err:
  m_initOK = False
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "Employee_ReadDB", "Read Employee", "Error reading in the employee from the database.")
  Resume Employee_ReadDB_End
  Resume
End Function
Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property
Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property
Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  ECASE ("RSBookMark")
End Property
Private Property Get IBenefitClass_RSBookMark() As String
  ECASE ("RSBookMark")
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
   
  
  On Error GoTo SetBenItemsInformation_err
  bc = BC_EMPLOYEE
  m_NeedToCalculate = True
    
  
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  With p11d32
    .BenDataLinkDataType(bc, ITEM_NIC_CLASS1A_ABLE) = TYPE_BOOL
    
    .BenDataLinkDataType(bc, ee_PersonnelNumber_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_PersonnelNumber_db) = S_PNUM
    .BenDataLinkMMFieldSize(bc, ee_PersonnelNumber_db) = 17
    .BenDataLinkDataType(bc, ee_FullName) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_FullName) = "Full name"
    .BenDataLinkDataType(bc, ee_Surname_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Surname_db) = S_SURNAME
    .BenDataLinkMMFieldSize(bc, ee_Surname_db) = 35
    .BenDataLinkDataType(bc, ee_Firstname_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Firstname_db) = S_FIRSTNAME
    .BenDataLinkMMFieldSize(bc, ee_Firstname_db) = 7
    .BenDataLinkDataType(bc, ee_Director_db) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, ee_Director_db) = "Director"
    .BenDataLinkMMFieldSize(bc, ee_Director_db) = 1
    
    .BenDataLinkDataType(bc, ee_NINumber_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_NINumber_db) = "NI Number"
    .BenDataLinkMMFieldSize(bc, ee_NINumber_db) = 9
    .BenDataLinkDataType(bc, ee_Selected) = TYPE_BOOL
    .BenDataLinkDataType(bc, ee_Email_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Email_db) = "Email"
    
    .BenDataLinkDataType(bc, ee_Salutation_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Salutation_db) = "Salutation"
    .BenDataLinkDataType(bc, ee_Title_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Title_db) = S_TITLE
    .BenDataLinkDataType(bc, ee_Initials_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Initials_db) = "Initials"
    
    .BenDataLinkDataType(bc, ee_Group1_db) = TYPE_STR
    .BenDataLinkDataType(bc, ee_Group2_db) = TYPE_STR
    .BenDataLinkDataType(bc, ee_Group3_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Group1_db) = "Group 1"
    .BenDataLinkUDMDescription(bc, ee_Group2_db) = "Group 2"
    .BenDataLinkUDMDescription(bc, ee_Group3_db) = "Group 3"
    
    .BenDataLinkDataType(bc, ee_Comments_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Comments_db) = "Comments"
    .BenDataLinkDataType(bc, ee_joined_db) = TYPE_DATE
    .BenDataLinkUDMDescription(bc, ee_joined_db) = "Joined"
    .BenDataLinkDataType(bc, ee_left_db) = TYPE_DATE
    .BenDataLinkUDMDescription(bc, ee_left_db) = "Left"
        
    .BenDataLinkDataType(bc, ee_DOB_db) = TYPE_DATE
    .BenDataLinkUDMDescription(bc, ee_DOB_db) = "D.O.B"
    
    .BenDataLinkDataType(bc, ee_Gender_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Gender_db) = "Gender"
      
    .BenDataLinkDataType(bc, ee_Username_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Username_db) = "Intranet Username"
    .BenDataLinkDataType(bc, ee_Password_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Password_db) = "Intranet Password"
    
    
    .BenDataLinkDataType(bc, ee_PaymentsForPrivateUseOfSharedVans_db) = TYPE_LONG
    .BenDataLinkDataType(bc, ee_OneOrMoreSharedVanAvailable_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db) = TYPE_BOOL
    .BenDataLinkDataType(bc, ee_ReportyDailyCalculationOfSharedVans_db) = TYPE_BOOL
    'FC - Class1A
    .BenDataLinkDataType(bc, ee_Class1AEmployeeIsNotSubjectTo_db) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, ee_Class1AEmployeeIsNotSubjectTo_db) = "Not subject to Class 1A?"
    
    .BenDataLinkDataType(bc, ee_RelevantDaysForDailySharedVanCalc_db) = TYPE_LONG
  
    
    .BenDataLinkDataType(bc, ee_AddressLine1_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_AddressLine1_db) = "Address Line 1"
    .BenDataLinkDataType(bc, ee_AddressLine2_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_AddressLine2_db) = "Address Line 2"
    .BenDataLinkDataType(bc, ee_AddressLine3_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_AddressLine3_db) = "Address Line 3"
    .BenDataLinkDataType(bc, ee_City_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_City_db) = "City"
    .BenDataLinkDataType(bc, ee_County_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_County_db) = "County"
    .BenDataLinkDataType(bc, ee_PostCode_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_PostCode_db) = "Postcode"
    .BenDataLinkDataType(bc, ee_Country_db) = TYPE_STR
    .BenDataLinkUDMDescription(bc, ee_Country_db) = "Country"
    
    .BenDataLinkUDMRequiresCalculate(bc, ee_Benefit) = True
    .BenDataLinkUDMRequiresCalculate(bc, ee_NIC_Class1A_Benefit) = True
    .BenDataLinkUDMRequiresCalculate(bc, ee_Class1A_Adjustment) = True
    
    .BenDataLinkDataType(bc, ee_Benefit) = TYPE_LONG
    .BenDataLinkDataType(bc, ITEM_NIC_CLASS1A_BENEFIT) = TYPE_DOUBLE
    .BenDataLinkDataType(bc, ITEM_BENEFIT_SUBJECT_TO_CLASS1A) = TYPE_LONG
    .BenDataLinkDataType(bc, ITEM_MADEGOOD_IS_TAXDEDUCTED) = TYPE_BOOL
    .BenDataLinkDataType(bc, ITEM_ERROR) = TYPE_STR
    
    .BenDataLinkDataType(bc, ee_Class1A_Adjustment) = TYPE_LONG
    .BenDataLinkDataType(bc, ee_VansOPRAAmountForegone_db) = TYPE_LONG
    .BenDataLinkDataType(bc, ee_VansFuelOPRAAmountForegone_db) = TYPE_LONG
    
    
    
    .BenDataLinkUDMDescription(bc, ee_Benefit) = "Total benefit"
    .BenDataLinkUDMDescription(bc, ITEM_BENEFIT_SUBJECT_TO_CLASS1A) = S_UDM_IR_BENEFIT_SUBJECT_TO_CLASS1A
    .BenDataLinkUDMDescription(bc, ITEM_NIC_CLASS1A_BENEFIT) = S_UDM_NIC_CLASS1A_BENEFIT
    .BenDataLinkUDMDescription(bc, ee_Class1A_Adjustment) = "Class1A additions"
     .BenDataLinkDataType(bc, ee_NIC_AdjustmentDeduct) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, ee_NIC_AdjustmentDeduct) = "Class1A deductions"
    
    .BenDataLinkDataType(bc, ee_NIInvalid) = TYPE_BOOL
    .BenDataLinkUDMDescription(bc, ee_NIInvalid) = "NI Number Invalid" ' EK added 1/04 TTP#51
    
    .BenDataLinkDataType(bc, ee_TotalBenefitsPotentiallySubjectToClass1A) = TYPE_LONG
    .BenDataLinkUDMDescription(bc, ee_TotalBenefitsPotentiallySubjectToClass1A) = S_UDM_BENEFITS_POTENTIALLY_SUBJECT_TO_CLASS1A
    
    
  End With

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
  Resume
End Sub
Private Function IsAlphaStrEx(s As String, ByVal CharPos As Long)
  Dim l As Long
  
  l = Asc(Mid$(UCASE(s), CharPos, 1))
  IsAlphaStrEx = (l >= 65) And (l <= 90)
  
End Function
Public Property Get NINumberValid() As Boolean
  NINumberValid = ValidateNI(IBenefitClass_value(ee_NINumber_db), False) = STANDARD
End Property

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  ECASE ("TABLE")
End Property
Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim ee As Employee
  Dim bWriteOK As Boolean
  
  Dim dDate As Date
  Dim i As Long

  On Error GoTo WriteDB_Err
  Call xSet("WriteDB")
  
  Set ben = Me
  Set ee = Me
  
  bWriteOK = True
  
  If Not p11d32.BringForward.Yes Then Call SetPanel2("Writing benefits and employee information to Database for: " & FullName)
  
  If ben.Dirty Then
    If Not MultiUserCheck() Then
      IBenefitClass_WriteDB = p11d32.CurrentEmployer.ReloadCurrentEmployee
      GoTo WriteDB_End
    ElseIf ben.InvalidFields = 0 Then
      Set rs = Me.Parent.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEE_WRITE, PersonnelNumber), dbOpenDynaset)
      If rs.EOF And rs.BOF Then
        rs.AddNew
        rs.Fields(S_FIELD_PERSONEL_NUMBER).value = PersonnelNumber
        rs.Fields("DateEdit").value = Now
        DateEdit = rs.Fields("DateEdit").value
      Else
        rs.Edit
      End If
      With ben
        rs.Fields("Surname").value = .value(ee_Surname_db)
        rs.Fields("Title").value = .value(ee_Title_db)
        rs.Fields("Initials").value = .value(ee_Initials_db)
        rs.Fields("Firstname").value = .value(ee_Firstname_db)
        rs.Fields("Salutation").value = .value(ee_Salutation_db)
        rs.Fields("Email").value = .value(ee_Email_db)
        rs.Fields("Group1").value = .value(ee_Group1_db)
        rs.Fields("Group2").value = .value(ee_Group2_db)
        rs.Fields("Group3").value = .value(ee_Group3_db)
        
        rs.Fields("NI").value = .value(ee_NINumber_db)
        If p11d32.BringForward.Yes Then
          If p11d32.BringForward.Comments Then
            rs.Fields("Comments").value = .value(ee_Comments_db)
          Else
            rs.Fields("Comments").value = ""
          End If
        Else
          rs.Fields("Comments").value = .value(ee_Comments_db)
        End If
        rs.Fields("Status").value = IIf(ben.value(ee_Director_db), "Director", "Staff")
        rs.Fields("SHAREDVANS").value = .value(ee_OneOrMoreSharedVanAvailable_db)
        rs.Fields("TWOPLUSVANS").value = .value(ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db)
        rs.Fields("DAILYCALC").value = .value(ee_ReportyDailyCalculationOfSharedVans_db)
        
        rs.Fields("Class1AEmployeeIsNotSubjectTo").value = .value(ee_Class1AEmployeeIsNotSubjectTo_db)
        rs.Fields("Joined").value = IIf(Len(.value(ee_joined_db)), .value(ee_joined_db), UNDATED)
        rs.Fields("Left").value = IIf(Len(.value(ee_left_db)), .value(ee_left_db), UNDATED)
        rs.Fields("EEPassword").value = .value(ee_Password_db)
        rs.Fields("EEUsername").value = .value(ee_Username_db)
        
        rs.Fields("Gender").value = .value(ee_Gender_db)
        rs.Fields("DOB").value = IIf(Len(.value(ee_DOB_db)), .value(ee_DOB_db), UNDATED)
        
        
        'RK These fields should not roll forward 18/06/03
        If p11d32.BringForward.Yes Then
          rs.Fields("RELEVANTDAYS").value = 0
          rs.Fields("PRIVATECONTRIB").value = 0
          rs.Fields("VansOPRAAmountForegone").value = 0
          rs.Fields("VansFuelOPRAAmountForegone").value = 0
        Else
          rs.Fields("RELEVANTDAYS").value = .value(ee_RelevantDaysForDailySharedVanCalc_db)
          rs.Fields("PRIVATECONTRIB").value = .value(ee_PaymentsForPrivateUseOfSharedVans_db)
          rs.Fields("VansOPRAAmountForegone").value = .value(ee_VansOPRAAmountForegone_db)
          rs.Fields("VansFuelOPRAAmountForegone").value = .value(ee_VansFuelOPRAAmountForegone_db)
        End If
        
        If Not WriteAddress Then GoTo WriteDB_End
        rs.Fields("DateEdit").value = Now
        ee.DateEdit = rs.Fields("DateEdit").value
        ben.ReadFromDB = True
        rs.Update
        ben.Dirty = False
      End With
    Else
      bWriteOK = False
    End If
  End If
  
  If bWriteOK Then
    bWriteOK = WriteBenefits()
  Else
    Call WriteBenefits
  End If
  
  If bWriteOK And p11d32.BringForward.Yes Then MDIMain.ClearConfirmUndo
  IBenefitClass_WriteDB = bWriteOK
  
WriteDB_End:
  If p11d32.BringForward.Yes Then Call SetPanel2("")
  Set ee = Nothing
  Set ben = Nothing
  Set rs = Nothing
  Call xReturn("WriteDB")
  Exit Function
WriteDB_Err:
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "WriteDB", "Writing employee details", "Error writing the employee details to the database.")
  Resume WriteDB_End
  Resume
  
End Function
Public Function nonSharedVans() As IBenefitClass

  On Error GoTo NonSharedVans_Err
  Call xSet("NonSharedVans")

  Set nonSharedVans = benefits(L_VANS_BENINDEX)

NonSharedVans_End:
  Call xReturn("NonSharedVans")
  Exit Function

NonSharedVans_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVans", "NonSharedVans", "Error returning the non shared vans benefit class.")
  Resume NonSharedVans_End
  Resume
End Function
Public Function AnyVanBenefit() As Boolean
  Dim vancol As nonSharedVans
  Dim ben As IBenefitClass
  Dim ey As Employer
  On Error GoTo AnyVanBenefit_ERR
  
  Call xSet("AnyVanBenefit")
  
  Set vancol = nonSharedVans()
  
  If vancol Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "AnyVanBenefit", "Van collection is nothing.")
  
  If vancol.Vans.CountValid = 0 Then
    If m_BenItems(ee_OneOrMoreSharedVanAvailable_db) = False Then GoTo AnyVanBenefit_END
    Set ey = Me.Parent
    If ey.SharedVans.Vans.CountValid = 0 Then GoTo AnyVanBenefit_END
  End If
  AnyVanBenefit = True
    
AnyVanBenefit_END:
  Call xReturn("AnyVanBenefit")
  Exit Function
AnyVanBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnyVanBenefit", "Any Van Benefit", "Error assessing whether there is any van benefit for " & FullName & ".")
  Resume AnyVanBenefit_END
  Resume
End Function
Public Function AnyLoanBenefit() As Boolean
  Dim loans As loans
  Dim ben As IBenefitClass
  Dim l As Long
  
  On Error GoTo AnyLoanBenefit_ERR
  
  Call xSet("AnyLoanBenefit")
  
  l = GetLoansBenefitIndex
  If l > 0 Then
    Set loans = benefits(l)
    If loans Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "AnyLoanBenefit", "Any Loan Benefit", "The loans collection is nothing.")
    If loans.loans.CountValid > 0 Then AnyLoanBenefit = True
  End If
AnyLoanBenefit_END:
  Call xReturn("AnyLoanBenefit")
  Exit Function
AnyLoanBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnyLoanBenefit", "Any Loan Benefit", "Error assessing whether there is any loan benefit for " & FullName & ".")
  Resume AnyLoanBenefit_END
End Function
Public Function MultiUserCheck() As Boolean
  Dim d As Date
  Dim ben As IBenefitClass
  Dim rs As Recordset
  Dim sMsg As String
  Dim i As Long
  
  On Error GoTo MultiUserCheck_Err
  Call xSet("MultiUserCheck")

  Set ben = Me
  
  If ben.ReadFromDB And (Not p11d32.BringForward.Yes) Then
    Set rs = p11d32.CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEECHECK, ben.value(ee_PersonnelNumber_db)), dbOpenDynaset, dbFailOnError)
    If Not (rs.EOF And rs.BOF) Then
      d = IIf(IsNull(rs.Fields("DateEdit").value), UNDATED, rs.Fields("DateEdit").value)
      If d <> DateEdit Then
        sMsg = "The employee you are trying to update has been changed by " & IIf(Len(rs.Fields("UserEdit").value & ""), rs.Fields("UserEdit").value, "UNKNOWN") & _
              " on " & Format$(d, "dd/mm/yy hh:nn:ss") & _
               vbCrLf & "Please confirm overwrite of these changes"
        i = MultiDialog("Employee has changed", sMsg, "Confirm", "Cancel")
        If i = 1 Then
          MultiUserCheck = True
        Else
          MultiUserCheck = False
        End If
      Else
        MultiUserCheck = True
      End If
      Set rs = Nothing
      Call DBEngine.Idle(dbFreeLocks)
    Else
      Call Err.Raise(ERR_NORECORDS, "SaveEmployee", "There are no records for the query " & SELECT_EMPLOYEECHECK)
    End If
  Else
    MultiUserCheck = True
  End If

MultiUserCheck_End:
  Call xReturn("MultiUserCheck")
  Exit Function

MultiUserCheck_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MultiUserCheck", "Multi User Check", "Error checking to see if another user has edited the selected employee.")
  Resume MultiUserCheck_End
  Resume
End Function


Public Function WriteBenefits() As Boolean
  Dim i As Long, b As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo WriteBenefits_Err
  Call xSet("WriteBenefits")

  b = True
  
  If Not p11d32.BringForward.Yes Then Call PrgStart(benefits.Count, "Writing benefits", ValueOfMax)
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If Not p11d32.BringForward.Yes Then
        Call PrgStepCaption("Writing benefit " & ben.Name)
      Else
        ben.ReadFromDB = False
      End If
      If ben.Dirty Or p11d32.BringForward.Yes Then
        If ben.InvalidFields = 0 Or IsBenefitACollection(ben) Then
          If b = False Then
            Call WriteBenefitsSub(ben)
          Else
            b = WriteBenefitsSub(ben)
          End If
        Else
          b = False
        End If
      End If
    End If
  Next

  WriteBenefits = b
  
WriteBenefits_End:
  If Not p11d32.BringForward.Yes Then Call PrgStop
  Call xReturn("WriteBenefits")
  Exit Function

WriteBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteBenefits", "Write Benefits", "Error writing the employees benefits.")
  Resume WriteBenefits_End
  Resume
End Function
Private Function WriteBenefitsSub(ben As IBenefitClass) As Boolean
  Dim benEE As IBenefitClass
  
  On Error GoTo WriteBenefitsSub_ERR
  
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WriteBenefitsSub", "The benefit is nothing")

  Set benEE = Me
  
  If ben.CompanyDefined And Not IsCBDEmployee(benEE.value(ee_PersonnelNumber_db)) Then GoTo WriteBenefitsSub_END

  If p11d32.BringForward.Yes Then
    If p11d32.BringForward.HMITSChosen And (2 ^ p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION)) Then
      If ben.CanBringForward Then
        WriteBenefitsSub = ben.writeDB()
      Else
        WriteBenefitsSub = True
      End If
    Else
      WriteBenefitsSub = True
    End If
  Else
    WriteBenefitsSub = ben.writeDB()
  End If
  
WriteBenefitsSub_END:
  Exit Function
WriteBenefitsSub_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteBenefitsSub", "Write Benefits Sub", "Error writing benefit sub.")
  Resume WriteBenefitsSub_END
End Function
Public Function GetLoansBenefitIndex() As Long
  Dim i As Long
  Dim ben As IBenefitClass
  
  On Error GoTo GetLoansBenefitIndex_Err
  Call xSet("GetLoansBenefitIndex")
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BC_LOANS_H Then
        GetLoansBenefitIndex = i
        Exit For
      End If
    End If
  Next i
  
GetLoansBenefitIndex_End:
  Call xReturn("GetLoansBenefitIndex")
  Exit Function

GetLoansBenefitIndex_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetLoansBenefitIndex", "Get Loans Benefit Index", "Error getting the loans collection benefit index.")
  Resume GetLoansBenefitIndex_End
End Function
Public Function IsBenefitACollection(ben As IBenefitClass) As Boolean

  On Error GoTo IsBenefitACollection_Err
  Call xSet("IsBenefitACollection")

  If Not ben Is Nothing Then
    IsBenefitACollection = ben.BenefitClass = BC_LOANS_H Or ben.BenefitClass = BC_NONSHAREDVANS_G
  End If

IsBenefitACollection_End:
  Call xReturn("IsBenefitACollection")
  Exit Function

IsBenefitACollection_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsBenefitACollection", "Is Benefit A Collection", "Error determining if the benefit is a collection")
  Resume IsBenefitACollection_End
End Function
Public Function PersonnelNumberChangeEx(ByVal sNewPN As String, Optional ByVal bRaiseError As Boolean = False) As Boolean
  Dim ben As IBenefitClass
  Dim t As TableDef
  Dim ibf As IBenefitForm2
  On Error GoTo err_err
  
  Set ben = Me
  sNewPN = Trim$(sNewPN)
  If StrComp(sNewPN, ben.value(ee_PersonnelNumber_db), vbTextCompare) <> 0 Then
    If Not m_Parent.ValidatePersonnelNumber(sNewPN, bRaiseError) Then GoTo err_end
    For Each t In m_Parent.db.TableDefs
      If FieldPresent(t.Fields, S_FIELD_PERSONEL_NUMBER) Then
        Call m_Parent.db.Execute(sql.Queries(UPDATE_PERSONNEL_NUMBER, t.Name, sNewPN, ben.value(ee_PersonnelNumber_db)))
      End If
    Next
    ben.value(ee_PersonnelNumber_db) = sNewPN
    If CurrentForm Is F_Employees Then
      If p11d32.CurrentEmployer.CurrentEmployee Is Me Then
        Set ibf = CurrentForm
        Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, Me)
        Call ibf.BenefitOn
      End If
    End If
    
  End If
  PersonnelNumberChangeEx = True
err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "PersonnelNumberChangeEx"), Err.Description)
End Function
Public Sub PersonnelNumberChange()
  Dim ben As IBenefitClass
  Dim sNewPN As String
  Dim t As TableDef
  
  On Error GoTo PersonnelNumberChange_ERR
  Call xSet("PersonnelNumberChange")
    
  Set ben = Me
  F_Input.ValText.TypeOfData = VT_STRING
  F_Input.ValText.AllowEmpty = False
  F_Input.ValText.MaxLength = 50
Start:
  If F_Input.Start("Change personnel number", "Enter a new pesonnel number.", ben.value(ee_PersonnelNumber_db)) Then
    If Not PersonnelNumberChangeEx(F_Input.ValText.Text) Then GoTo Start:
  End If
  
PersonnelNumberChange_END:
  Call xSet("Personnel Number Change")
  Exit Sub
PersonnelNumberChange_ERR:
  Set F_Input = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "PersonnelNumberChange", "Personnel Number Change", "Error changing the personnel number of an employee.")
  Resume PersonnelNumberChange_END
  Resume
End Sub
Private Function P46ReplacedCarsExtract(ByVal bFromPayeOnline)
  P46ReplacedCarsExtract = (p11d32.ReportPrint.P46PrintReplacedP46s And (Not bFromPayeOnline)) Or (bFromPayeOnline And p11d32.AppYear > 2010)
End Function
Public Function GetP46Cars(P46Cars As ObjectList, ByVal dDateFrom As Date, ByVal dDateTo As Date, Optional bFromPayeOnline As Boolean = False) As Boolean
  Dim i As Long, j As Long, k As Long
  Dim CompanyCar As IBenefitClass, CompanyCarDateCheck As IBenefitClass
  Dim CompanyCars As ObjectList
  Dim bForcePrintingP46 As Boolean, bDateFromPass, bDateToPass As Boolean
  Dim ValidP46Car As Boolean
  Dim dTaxYearStart As Date, dTaxYearEnd As Date, dDateFromLess1Day As Date
  
  On Error GoTo GetP46Cars_Err
  Call xSet("GetP46Cars")
  dTaxYearStart = p11d32.Rates.value(TaxYearStart)
  dTaxYearEnd = p11d32.Rates.value(TaxYearEnd)
  
  Set CompanyCars = BenefitsOfType(Me, BC_COMPANY_CARS_F)
  
  Set P46Cars = New ObjectList
    
  For i = 1 To CompanyCars.Count
    Set CompanyCar = CompanyCars(i)
    
      With CompanyCar
          If IsNumeric(CompanyCar.Calculate) Then
            'set defaults, moved from car calc to enable CalculateHelper
            'no need to worry about the dirty flag and needtocalculate as will always calc these fields for p46 cars
            CompanyCar.value(car_P46CarProvidedReplaced) = False
            CompanyCar.value(car_P46FirstProvidedWithCar) = False
            CompanyCar.value(car_p46ReplacementForMultipleCarsMakeAndModel) = False
            CompanyCar.value(car_P46SecondCar) = False
            CompanyCar.value(car_P46WithdrawnWithoutReplacement) = False
            
            bForcePrintingP46 = .value(car_ForceP46_db)
            bDateFromPass = P46DateInRange(.value(Car_AvailableFrom_db), dDateFrom, dDateTo, bForcePrintingP46, dTaxYearStart, dTaxYearEnd)
            bDateToPass = P46DateInRange(.value(Car_AvailableTo_db), dDateFrom, dDateTo, bForcePrintingP46, dTaxYearStart, dTaxYearEnd)
                  
            If Not bDateFromPass And Not bDateToPass Then
              ValidP46Car = False
            ElseIf bDateToPass And Not bDateFromPass And .value(car_Replaced_db) Then
              ValidP46Car = False
            Else
              'valid p46 cars
              If bDateToPass Then
                If Not .value(car_Replaced_db) Then
                  .value(car_P46WithdrawnWithoutReplacement) = True
                  ValidP46Car = True
                End If
              End If
              If bDateFromPass Then
                If P46ReplacedCarsExtract(bFromPayeOnline) Then
                  If .value(car_Replacement_db) Then
                    .value(car_P46CarProvidedReplaced) = True
                  Else
                    .value(car_P46FirstProvidedWithCar) = Not .value(car_Second_db)
                  End If
                  ValidP46Car = True
                Else
                  ValidP46Car = False
                  
                  
                  If .value(car_Replacement_db) Then
                    ValidP46Car = False
                    'START REMOVED FOR 2011
                    'If .value(car_Replaced_db) Then
                    '  ValidP46Car = False
                    'End If
                    .value(car_P46CarProvidedReplaced) = True
                    'END REMOVED FOR 2011
                  Else
                    .value(car_P46FirstProvidedWithCar) = Not .value(car_Second_db)
                    ValidP46Car = True
                  End If
                  
                End If
                .value(car_P46SecondCar) = .value(car_Second_db)
              End If
            End If
            If ValidP46Car Then
              .value(car_P46PrintCarDetails) = True
              If .value(car_P46CarProvidedReplaced) Then
                'IF MORE THAN 2 CARS FROM - TO SPAN OR HIT MY AVAILABLE FROM - 1 DAY THEN NEED TO PRINT DETAILS
                dDateFromLess1Day = DateAdd("d", -1, .value(Car_AvailableFrom_db))
                For j = 1 To CompanyCars.Count
                  If j <> i Then
                    Set CompanyCarDateCheck = CompanyCars(j)
                    If DateInRange(dDateFromLess1Day, CompanyCarDateCheck.value(Car_AvailableFrom_db), CompanyCarDateCheck.value(Car_AvailableTo_db)) Then
                      k = k + 1
                      If k > 1 Then
                        .value(car_p46ReplacementForMultipleCarsMakeAndModel) = True
                        Exit For
                      End If
                    End If
                  End If
                Next
              End If
              
              Call P46Cars.Add(CompanyCar)
              GetP46Cars = True
            End If
          End If
      End With
  Next i

  


GetP46Cars_End:
  Call xReturn("GetP46Cars")
  Exit Function

GetP46Cars_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetP46Cars", "Get P46 Cars", "Error finding the P46 Cars.")
  Resume GetP46Cars_End
End Function

Private Function P46DateInRange(ByVal dDate As Date, ByVal dFrom As Date, ByVal dTo As Date, ByVal Overide As Boolean, ByVal TaxYearStart As Date, ByVal TaxYearEnd As Date) As Boolean
  If DateInRange(dDate, dFrom, dTo) Then
    P46DateInRange = True
    If (dDate = TaxYearStart) Or (dDate = TaxYearEnd) Then
      P46DateInRange = Overide
    End If
  End If
End Function
Private Function WriteAddress() As Boolean
  Dim ben As IBenefitClass
  
  On Error Resume Next
  
  Call xSet("WriteAddress")
  
  WriteAddress = True
  
  If Not p11d32.BringForward.Yes Then Call SetPanel2("Writing address for " & FullName)
  
  Set ben = Me
  If m_bHasAddress Then
    Call m_Parent.db.Execute(sql.Queries(UPDATE_ADDRESS, ben.value(ee_AddressLine1_db), ben.value(ee_AddressLine2_db), ben.value(ee_AddressLine3_db), ben.value(ee_City_db), ben.value(ee_County_db), ben.value(ee_PostCode_db), ben.value(ee_Country_db), ben.value(ee_PersonnelNumber_db)), dbFailOnError)
    If m_Parent.db.RecordsAffected > 0 Then Exit Function
    On Error GoTo WriteAddress_Err
    Call m_Parent.db.Execute(sql.Queries(ADD_ADDRESS, ben.value(ee_AddressLine1_db), ben.value(ee_AddressLine2_db), ben.value(ee_AddressLine3_db), ben.value(ee_City_db), ben.value(ee_County_db), ben.value(ee_PostCode_db), ben.value(ee_Country_db), ben.value(ee_PersonnelNumber_db)), dbFailOnError)
  End If
  
WriteAddress_End:
  If Not p11d32.BringForward.Yes Then Call SetPanel2("")
  Call xReturn("WriteAddress")
  Exit Function

WriteAddress_Err:
  WriteAddress = False
  Call ErrorMessage(ERR_ERROR, Err, "WriteAddress", "Write Address", "Error writing the address to the database.")
  Resume WriteAddress_End
  Resume
End Function
'cad cdb new stuff
Public Sub CDBBenefitAdd(ByVal ben As IBenefitClass)
  Dim benOther As other
  Dim benCopy As IBenefitClass
  Dim benEE As Employee
  
  If IsCBDEmployee(Me.PersonnelNumber) Then
    Call benefits.Add(ben)
  Else
    Set benCopy = ben.Copy(Me)
    
    benCopy.RSBookMark = ""
    benCopy.Dirty = False
     
    Set benOther = benCopy
    Set benOther.CDBMasterBenefitLink = ben
    Set benCopy.Parent = Me
    ben.NeedToCalculate = True
    ben.NeedToCalculate = True
  End If
End Sub
'cad cdb
Public Function CreateCDBLinks() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim sCriteria As String
  Dim benCopy As IBenefitClass
  Dim benOther As other
  
  On Error GoTo CreateCDBLinks_Err
  Call xSet("CreateCDBLinks")
 
  Set ben = Me
  If Not (Me Is p11d32.CurrentEmployer.CDBEmployee) Then
    'create links to objects in CDBEmpoyee if required
    sCriteria = "P_NUM = '" & m_BenItems(ee_PersonnelNumber_db) & "'"
    Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_CDB_LINKS)
    rs.FindFirst (sCriteria)
    If Not rs.NoMatch Then
      If GetBenefitFromCDBEmployee(ben, "" & rs.Fields("BenCode").value) Then
        Call CDBBenefitAdd(ben)
      End If
      rs.FindNext (sCriteria)
      Do While Not rs.NoMatch
        If GetBenefitFromCDBEmployee(ben, "" & rs.Fields("BenCode").value) Then
          'cad cdb fix
          'new stuff
          Call CDBBenefitAdd(ben)
        End If
        rs.FindNext (sCriteria)
      Loop
    End If
  End If

CreateCDBLinks_End:
  Set ben = Nothing
  Set rs = Nothing
  Call xReturn("CreateCDBLinks")
  Exit Function

CreateCDBLinks_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CreateCDBLinks", "Create CDB Links", "Error creating company defined benefit links.")
  Resume CreateCDBLinks_End
  Resume
End Function


Public Function GetBenefitFromCDBEmployee(ben As IBenefitClass, sBenCode As String) As Boolean
  Dim i As Long
  Dim oth As other
  
  On Error GoTo GetBenefitFromCDBEmployee_Err
  Call xSet("GetBenefitFromCDBEmployee")

  With p11d32.CurrentEmployer.CDBEmployee
    Call .LoadBenefits(TBL_OTHER, , , True)
    For i = 1 To .benefits.Count
      Set ben = .benefits(i)
      If Not ben Is Nothing Then
        If IsBenOtherClass(ben.BenefitClass) Then
          Set oth = ben
          If StrComp(oth.PersonnelNumber, sBenCode, vbTextCompare) = 0 Then
            GetBenefitFromCDBEmployee = True
            Exit For
          End If
        End If
      End If
      Set ben = Nothing
    Next
  End With
  
GetBenefitFromCDBEmployee_End:
  Call xReturn("GetBenefitFromCDBEmployee")
  Exit Function

GetBenefitFromCDBEmployee_Err:
  Set ben = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "GetBenefitFromCDBEmployee", "Get Benefit From CDB Employee", "Error retrieving the company defined benefits from the CDB employee.")
  Resume GetBenefitFromCDBEmployee_End
  Resume
End Function


Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim l0 As Long, l1 As Long
  
  l0 = v0
  l1 = v1
  
  If l0 = BC_LOANS_H Then
    l0 = BC_LOAN_OTHER_H
  ElseIf l1 = BC_LOANS_H Then
    l1 = BC_LOAN_OTHER_H
  End If
  
  If l0 < l1 Then
    ISortFunction_CompareItems = -1
  ElseIf l0 > l1 Then
    ISortFunction_CompareItems = 1
  Else
    ISortFunction_CompareItems = 1
  End If
  
End Function
Public Function BringForwardPrepare(ByVal HMITS As Long, ByVal BringForwardType As BRING_FORWARD_TYPE) As Boolean
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  Dim ey As Employer
  Dim sSQL As String, sWhere As String
  Const sDel As String = "Delete * from "
  
  On Error GoTo BringForwardPrepare_ERR
  
  Call xSet("BringForwardPrepare")
  
  If BringForwardType <> BFT_UPDATE Then GoTo BringForwardPrepare_END
  
  sWhere = " Where (" & S_FIELD_PERSONEL_NUMBER & " = " & StrSQL(PersonnelNumber) & ") "
  
  Set ben = Me
  Set ey = ben.Parent
  'basically we will get the section chosen by the user and
  'delete the benefits that are chosen to bring forward
  For bc = BC_FIRST_ITEM To BC_UDM_BENEFITS_LAST_ITEM
    If (2 ^ p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION)) And HMITS Then
      sSQL = sDel & p11d32.Rates.BenClassTo(bc, BCT_TABLE) & sWhere
      Select Case bc
        Case BC_COMPANY_CARS_F
          'delete miles
          Call ey.db.Execute(sSQL)
'MP DB          Call ey.db.Execute(sDel & "T_CARMILES" & sWhere)
        Case BC_QUALIFYING_RELOCATION_J
          Call ey.db.Execute(sSQL)
        Case BC_LOAN_OTHER_H
          Call ey.db.Execute(sSQL)
          Call ey.db.Execute(sDel & "T_LOANS" & sWhere)
        Case BC_EMPLOYEE_CAR_E
          Call ey.db.Execute(sSQL)
          Call ey.db.Execute(sDel & "T_BENEECARDETAILS" & sWhere)
        Case BC_SERVICES_PROVIDED_K
          Call ey.db.Execute(sSQL & " AND (UDBCode=" & StrSQL(S_SERVICESPROVIDED_UDBCODE) & ")")
        Case BC_ASSETSTRANSFERRED_A
          Call ey.db.Execute(sSQL)
        Case Else
          If IsBenOtherClass(bc) Then
              sSQL = sSQL & " AND " & BenOtherWhereClauseSub(bc)
              ey.db.Execute (sSQL)
            
          Else
            'STANDARD BEN IE PHONES
            ey.db.Execute (sSQL)
          End If
      End Select
    End If
  Next
  
  BringForwardPrepare = True
  
BringForwardPrepare_END:
  Call xReturn("BringForwardPrepare")
  Exit Function
BringForwardPrepare_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BringForwardPrepare", "Bring Forward Prepare", "Error preparing an employee for bring forward.")
  Resume BringForwardPrepare_END
  Resume
End Function
Private Sub EnumBenefitAction(ByVal ben As IBenefitClass, ByVal value As Variant, ByVal bet As BENEFITS_ENUM_TYPE)
  Select Case bet
    Case BET_NEED_TO_CALCULATE
      ben.NeedToCalculate = True
    Case Else
      ECASE ("Invalid Enum Benefit Action:" & bet)
  End Select
End Sub

Public Sub EnumBenefitsDoAction(ByVal value As Variant, ByVal bet As BENEFITS_ENUM_TYPE)
  Dim i As Long, j As Long
  Dim ben As IBenefitClass
  Dim benSub As IBenefitClass
  Dim lns As loans
  Dim Vans As nonSharedVans
  On Error GoTo err_err
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = BC_LOANS_H Then
        Set lns = ben
        For j = 1 To lns.loans.Count
          Set benSub = lns.loans(j)
          If Not benSub Is Nothing Then
            Call EnumBenefitAction(benSub, value, bet)
          End If
        Next
      ElseIf ben.BenefitClass = BC_NONSHAREDVANS_G Then
        Set Vans = ben
        For j = 1 To Vans.Vans.Count
          Set benSub = Vans.Vans(j)
          If Not benSub Is Nothing Then
            Call EnumBenefitAction(benSub, value, bet)
          End If
        Next
      End If
      Call EnumBenefitAction(ben, value, bet)
     
    End If
  Next
  
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "EnumBenefitsDoAction"), Err.Description)
  Resume
End Sub
