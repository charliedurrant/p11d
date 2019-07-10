VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBDEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Implements IBenefitClass

'Employee class module
Public Enum EmployeeItems
  ee_Firstname
  ee_Surname
  ee_Email
  ee_Salutation
  ee_Title
  ee_Initials
  ee_PersonnelNumber
  ee_Payref
  ee_Group1
  ee_Group2
  ee_Group3
  ee_Comments
  ee_NINumber
  ee_joined
  ee_LEFT
  ee_Status
  ee_TotalEmployeeCarMiles 'backwards
  ee_TotalEmployeeCarAllowance 'backwards
  ee_PaymentsForPrivateUseOfSharedVans
  ee_OneOrMoreSharedVanAvailable
  ee_NonSharedVanAvailableAtSameTimeAsSharedVan
  ee_ReportyDailyCalculationOfSharedVans
  ee_RelevantDaysForDailySharedVanCalc
  ee_LASTITEM

End Enum

Public benefits As ObjectList

Private m_Locked As Boolean
Private m_dateedit As Date
Private m_dirty As Boolean

' inititalized at -1, when loaded contains no of bens.


Private m_initOK As Boolean
Private m_Parent As Object
Private m_InvalidFields As Long
Private m_InvalidVanFields As Long
Private m_bCalculatingDependantEmployeeCars As Boolean
Private m_EmployeeItems(0 To ee_LASTITEM) As Variant
Private m_bentype As benClass
Private m_ReadFromDB As Boolean
Private m_bBenefitsLoaded As Boolean

Public Property Get CalculatingDependantEmployeeCars() As Boolean
  CalculatingDependantEmployeeCars = m_bCalculatingDependantEmployeeCars
End Property
Public Property Let CalculatingDependantEmployeeCars(NewValue As Boolean)
  m_bCalculatingDependantEmployeeCars = NewValue
End Property

Private Function BenefitIsLoan(benefit As IBenefitClass) As Boolean
  If Not benefit Is Nothing Then BenefitIsLoan = (benefit.BenefitClass = BC_HOMELOAN) Or (benefit.BenefitClass = BC_OTHERLOAN)
End Function

Public Function RemoveBenefit(ibf As IBenefitForm2, ben As IBenefitClass, ByVal BenefitIndex As Long) As Boolean
  Dim NextBenefitIndex As Long
  
  Dim loans As loans
  
  On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  
  'need special treatement for loans as they are collections within collections
  If Not ben.CompanyDefined Then
    'need special treatement for loans as they are a collection within a collection
    Call ben.DeleteDB
    'get the nextbest listitem tag
    NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
    If BenefitIsLoan(ben) Then
      'then my index is for the loans collection and not the benefits collection
      Set loans = ben.Parent
      ' check to see if my parent, the loans collection, has any items apart from me
      If loans.count > 1 Then
        'we have others so remove myself
        loans.Remove (BenefitIndex)
      Else
        'get rid of the loans collection in the benefits collection
        Call benefits.Remove(benefits.ItemIndex(loans))
      End If
    Else
      Call benefits.Remove(BenefitIndex)
    End If
    're-add the benefits to the list view
    Call ibf.BenefitsToListView
    'select an item
    Call SelectBenefitByBenefitIndex(ibf, NextBenefitIndex)
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

Public Function IterateBenefits(IT As ITERATION_TYPE) As Boolean
  Dim l As Long
  Dim ben As IBenefitClass
  Dim lDirtyBenefits As Long, lBenefitsWritten As Long
  
On Error GoTo IterateBenefits_ERR
  
  Call xSet("IterateBenefits")
  
  
  For l = 1 To benefits.count
    Set ben = benefits.Item(l)
    If Not (ben Is Nothing) Then
      Select Case IT
        Case ITERATION_TYPE.AnyDirty
          If ben.Dirty Then
            IterateBenefits = True
            Exit For
          End If
        Case ITERATION_TYPE.AnyInvalid
          If ben.InvalidFields > 0 Then
            IterateBenefits = True
            Exit For
          End If
        Case ITERATION_TYPE.WriteDirtyAndAreAllWritten
          If ben.Dirty Then
            lDirtyBenefits = lDirtyBenefits + 1
            If ben.InvalidFields = 0 Or IsBenefitACollection(ben) Then
              lBenefitsWritten = lBenefitsWritten + Abs(ben.WriteDB)
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

Public Property Get FullName() As String

  Dim s As String
  'construct the name with
  'Surname Title firstname Initials
  If Len(m_EmployeeItems(ee_Surname)) Then s = s & m_EmployeeItems(ee_Surname)
  s = Trim$(s & " " & m_EmployeeItems(ee_Title))
  s = Trim$(s & " " & m_EmployeeItems(ee_Firstname))
  s = Trim$(s & " " & m_EmployeeItems(ee_Initials))
  FullName = s
    
End Property
Public Property Get Dirty() As Boolean
  Dirty = m_dirty
End Property
Public Property Let Dirty(NewValue As Boolean)
  m_dirty = NewValue
End Property
Private Sub Class_Initialize()
  Dim i As Long
  
  On Error GoTo Employee_Initialize_err
  Call xSet("Employee_Initialize")
  
  Set benefits = New ObjectList
  m_initOK = True
  
Employee_Initialize_end:
  Call xReturn("Employee_Initialize")
  Exit Sub
  
Employee_Initialize_err:
  m_initOK = False
  Resume Employee_Initialize_end
End Sub
Private Function KillBenefits() As Boolean
  Dim i As Long
  Dim ben As IBenefitClass
  
  Call xSet("KillBenefits")
  
  If Not benefits Is Nothing Then
    For i = 1 To benefits.count
      If Not (benefits(i) Is Nothing) Then
        Set ben = benefits(i)
        Call ben.Kill
        Call benefits.Remove(i)
        Set ben = Nothing
      End If
    Next i
    Call benefits.Compact
  End If
  
  Call xReturn("KillBenefits")
  
End Function
Public Property Get InitOK() As Boolean
  InitOK = m_initOK
End Property

Public Function LoadBenefits(ben As BENEFIT_TABLES) As Boolean
  Dim benobj As IBenefitClass
  Dim i As Long
  
  On Error GoTo LoadBenefits_Err
  
  Call xSet("LoadBenefits")
  
  If m_bBenefitsLoaded Then GoTo LoadBenefits_End
  
  Call KillBenefits
  
  
  
  For i = 0 To TBL_LASTITEM - 1
    Set benobj = Nothing
    Select Case i
      Case TBL_VANS
        If benefits.count = 0 Then
          Set benobj = New NonSharedVans
        Else
          ECASE ("NonSharedVans are not the first benefit")
        End If
      Case TBL_COMPANYCAR
        Set benobj = New CompanyCar
      Case TBL_SHARED_VANS
        'nothing as in employer level
      Case TBL_OTHER
        Set benobj = New Other
      Case TBL_P
        Set benobj = New P
      Case TBL_HOMELOAN, TBL_BENLOAN
        'If (m_BenefitLoaded(TBL_BENLOAN) = -1) Or (m_BenefitLoaded(TBL_HOMELOAN) = -1) Then
        Set benobj = New loans
        'End If
      Case TBL_COMPANYDEFINED
        Call LoadCDB
        Set benobj = Nothing
      Case TBL_PHONE
        Set benobj = New Phone
      Case TBL_ASSETSATDISPOSAL
        Set benobj = New AssetsAtDisposal
      Case TBL_SERVICESPROVIDED
        Set benobj = New ServicesProvided
      Case TBL_P
        Set benobj = New P
      Case TBL_ACCOM
        Set benobj = New Accomodation
      Case TBL_EECAR
        Set benobj = New EmployeeCar
      Case TBL_RELOC
        Set benobj = New Relocation
      Case TBL_TRANS
        Set benobj = New AssetsTransferred
      Case TBL_ALLBENEFITS
        
      Case TBL_CDRNEW
      
      Case Else
        Call ECASE(" NOT A RECOGNIZED BENEFIT")
    End Select
    
    If Not benobj Is Nothing Then
      Set benobj.Parent = Me
      If benobj.ReadDB Then
        Call benefits.Add(benobj)
      End If
    End If
  Next i
  
  For i = 1 To benefits.count
    If Not benefits(i) Is Nothing Then
      Set benobj = benefits(i)
      Call benobj.ReadDB
    End If
    
  Next
  
  
  m_bBenefitsLoaded = True
  LoadBenefits = True
  
LoadBenefits_End:
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
Private Function LoadCDB() As Long
  Dim ben As IBenefitClass
  Dim rs As Recordset
  Dim s As String, q As String
  Dim i As Long, RefCount As Long
  On Error GoTo LoadCDB_err
  
  Set rs = m_Parent.rsBenTables(TBL_COMPANYDEFINED)
  s = Me.PersonnelNumber
  Call rs.FindFirst("P_NUM = '" & s & "'")
  Do Until rs.NoMatch
    q = rs.Fields("BENCODE") 'apf possible null if no benefit for this person
    For i = 1 To m_Parent.SysEmployee.benefits.count
      Set ben = m_Parent.SysEmployee.benefits(i)
      If ben.Reference = q Then
        Call Me.benefits.Add(ben)
        RefCount = RefCount + 1
      End If
    Next i
    Set ben = Nothing
    rs.FindNext ("P_NUM = '" & s & "'")
  Loop
  
LoadCDB_end:
  
  Set rs = Nothing
  Exit Function
LoadCDB_err:
  ErrorMessage ERR_ERROR, Err, "LoadCDB", "ERR_READDB", "Unable to load in the employees company defined benefits"
  Resume LoadCDB_end
End Function


Public Property Get InvalidVanFields() As Long
  InvalidVanFields = m_InvalidVanFields
End Property

Public Property Let InvalidVanFields(ByVal l As Long)
  m_InvalidVanFields = l
End Property

Public Function ReportHeader(ReportType As P11DREPORT_TYPES) As String
'Dim q As String
'On Error GoTo rptheader_Err
'  Call xSet("rptheader")
'
'    q = "{Arial=7,n}{X=0}{BOX=100,6,f}" & "{Arial=6}" & STR_CR
'    q = q & "{Arial=13,BI}{X=10}{BW}{YREL=200}Return of Expenses payments and benefits not covered by dispensation." & "{YREL=-100}" & STR_CR
'    q = q & "{x=10}{YREL=200}Year Ended 5 April 1998 " & "{YREL=-100}" & STR_CR & STR_CR
'    q = q & "{Arial=7,NB}{WB}" & STR_CR
'    q = q & "{Arial=7,B}Employer's details"
'    q = q & "{x=50}Employee's details" & STR_CR
'    q = q & "{Arial=7}Employer's name"
'    q = q & "{Arial=7}{x=50}Employee's name" & STR_CR & STR_CR
'    q = q & "{Times=10,bi}{WBTEXTBOXL=40,2,""" & Me.Parent.Name & """}"
'    q = q & "{x=50}{WBTEXTBOXL=25,2,""" & Me.Name & """}" & STR_CR & STR_CR
'    q = q & "{Arial=7}PAYE tax reference"
'    q = q & "{Arial=7}{x=50}National Insurance number" & STR_CR
'    q = q & "{Times=10,nbi}{WBTEXTBOXL=25,2,""" & Me.Parent.Payeref & """}"
'    q = q & "{x=50}{WBTEXTBOXL=23,2,""" & Me.NINumber & """}" & STR_CR & STR_CR & STR_CR
'    ReportHeader = q
'rptheader_End:
'  Call xReturn("rptheader")
'  Exit Function
'
'rptheader_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "rptheader", "ERR_UNDEFINED", "Undefined error.")
'  Resume rptheader_End
End Function

Public Function ReportFooter(Description As String) As String
  ReportFooter = "{Arial=8,nb}{x=0}P11D " & P11d32.Rates.FormYear & Description & "(Arthur Andersen){Arial=8,i}{x=50}Page {PAGE}" & " -  printed at {TIME} on {DATE}"
End Function

Public Sub PrintOther(bt As benClass, rep As Reporter)
  Dim ben As IBenefitClass
  Dim i As Long
  On Error GoTo PrintOther_Err
  Call xSet("PrintOther")
  
  If rep.InitReport("Other benefits", gtarget, PORTRAIT) Then
    rep.ReportHeader = Me.ReportHeader("(Working paper)")
    rep.PageFooter = Me.ReportFooter("(Working paper)")
    For i = 1 To benefits.count
      Set ben = benefits(i)
      If Not (ben Is Nothing) Then
        If ben.BenefitClass = bt Then
          ben.PrintWk (rep)
        End If
      End If
    Next i
  End If
  Call rep.EndReport
  
PrintOther_End:
  Call xReturn("PrintOther")
  Exit Sub

PrintOther_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PrintOther", "ERR_UNDEFINED", "Undefined error.")
  Resume PrintOther_End
End Sub
Public Sub PrintHMIT(MainRep As Reporter)
  Dim rep As Reporter
  
  On Error GoTo printwk_Err
  Call xSet("PrintHMIT")
  If MainRep Is Nothing Then
    Set rep = New Reporter
  Else
    Set rep = MainRep
  End If
  If rep.InitReport("HMIT form", gtarget, PORTRAIT) Then
    rep.ReportHeader = ""
    rep.ReportFooter = ""
    rep.PageHeader = ""
    rep.PageFooter = Me.ReportFooter("")
    Call HMITPage1(rep)
    Call HMITPage2(rep)
  End If
  Call rep.EndReport
  If MainRep Is Nothing Then
    If gtarget = PREPARE_REPORT Then Call rep.PreviewReport
  End If
printwk_End:
  Call xReturn("PrintHMIT")
  Exit Sub

printwk_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PrintWK", "ERR_UNDEFINED", "Undefined error.")
  Resume printwk_End
End Sub
Public Function DumpBens() As Boolean
  Dim i As Long, ben As IBenefitClass
  On Error GoTo DumpBens_Err
  Call xSet("DumpBens")
  For i = 1 To benefits.count
    Set ben = benefits(i)
    If Not (ben Is Nothing) Then
      Debug.Print i & vbTab & ben.Name
    End If
  Next i
DumpBens_End:
  Call xReturn("DumpBens")
  Exit Function

DumpBens_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DumpBens", "ERR_UNDEFINED", "Undefined error.")
  Resume DumpBens_End
End Function
Private Sub HMITPage1(rep As Reporter)
'  Dim ben As IBenefitClass
'  Dim BenIndex As Long
'
'  Dim i As Long
'
'  Dim BusMiles1 As String
'  Dim BusMiles2 As String
'  Dim Diesel1 As Boolean
'  Dim Diesel2 As Boolean
'
'  Dim sDesc As Variant
'  Dim lval As Variant
'  Dim lmg As Variant
'  Dim lben As Variant
'
'  BusMiles1 = "0 TO 2499"
'  BusMiles2 = "0 TO 2499"
'
' On Error GoTo Page1_Err
'  Call xSet("Page1")
'
'  'Employee and employer details
'  rep.Out "{Arial=7,n}{X=0}{BOX=100,4,f}" & "{Arial=6}" & vbCrLf
'  rep.Out "{Arial=13,BI}{X=50}{BW}{YREL=200}P11D EXPENSES AND BENEFITS 1997-98" & "{YREL=-100}" & vbCrLf & vbCrLf
'  rep.Out "{Arial=13,BI}{x=100}{BW}"
'  rep.Out "{Arial=7,B}{WB}" & vbCrLf
'  rep.Out "{Arial=7,B}{WB}Note to employer"
'  rep.Out "{Arial=7,B}{x=50}Note to employee" & vbCrLf
'  rep.Out "{Arial=7}Complete this return for a director, or an employee who earned at a rate of "
'  rep.Out "{x=50}Your employer has filled in this form.  Keep it in a safe place as you may not" & vbCrLf
'  rep.Out "£8,500 a year or more during the year 6 April " & Format$(P11d32.Rates.GetItem(taxyearstart), "YYYY") & " to 5 April " & Format$(P11d32.Rates.GetItem(taxyearend), "YYYY") & ".  Do not"
'  rep.Out "{x=50}be able to get a duplicate. You will need it for your tax records and to" & vbCrLf
'  rep.Out "include expenses and benefits covered by a dispensation or PAYE settlement"
'  rep.Out "{x=50}complete your " & S_THISYEAR & "  Tax Return if you get one.  Your tax code may need" & vbCrLf
'  rep.Out "agreement. Read the P11D Guide and booklet 480, Chapter 24, before you"
'  rep.Out "{x=50}to be adjusted to take account of the information given on this P11D." & vbCrLf
'  rep.Out "complete the form. Send the completed P11D and form P11D(b) to the Tax"
'  rep.Out "{x=50}The box numbers on this P11D have the same numbering as the" & vbCrLf
'  rep.Out "Office by 6 July " & Format$(P11d32.Rates.GetItem(taxyearend), "YYYY") & ". You must give a copy of this information to the director or"
'  rep.Out "{x=50}Employement Pages of the Tax Return, for example, 1.12. Include the total" & vbCrLf
'  rep.Out "employee by the same date.  The term employee is used to cover both directors"
'  rep.Out "{x=50}figures in the corresponding box on the Tax Return, unless you think some" & vbCrLf
'  rep.Out "and employees throughout the rest of this form."
'  rep.Out "{x=50}other figure is more appropriate." & vbCrLf & vbCrLf
'  rep.Out "{x=0}{LINE=100}" & vbCrLf
'  rep.Out "{Arial=7,B}Employer's details"
'  rep.Out "{x=50}Employee's details" & vbCrLf
'  rep.Out "{Arial=7}Employer's name"
'  rep.Out "{Arial=7}{x=50}Employee's name" & vbCrLf & vbCrLf
'  rep.Out "{Times=10,bi}{WBTEXTBOXL=40,2," & m_Parent.Name & "}"
'  rep.Out "{x=50}{WBTEXTBOXL=30,2," & Name & "}"
'  If Status Then
'    rep.Out "{Wingdings=12,b}{x=96}{WBTEXTBOXC=3,2,ü}"
'  Else
'    rep.Out "{x=96}{BOX=3,2}"
'  End If
'  rep.Out "{Arial=6}{x=88}Tick here" & vbCrLf
'  rep.Out "{Arial=6}{x=88}if a director" & vbCrLf & vbCrLf & vbCrLf
'  rep.Out "{Arial=7}PAYE tax reference"
'  rep.Out "{x=50}Works number or department"
'  rep.Out "{Arial=7,R}{x=99}National Insurance number" & vbCrLf
'  rep.Out "{Times=10,nbi}{WBTEXTBOXL=25,2," & m_Parent.Payeref & "}"
'  rep.Out "{x=50}{WBTEXTBOXL=23,2," & PersonnelNumber & "}"
'  rep.Out "{x=76}{WBTEXTBOXL=23,2," & NINumber & "}" & vbCrLf & vbCrLf & vbCrLf
'
'  ' Assets transferred
'  BenIndex = HasBenefit(BC_ASSETSTRANSFERRED, "")
'  If BenIndex > 0 Then
'    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_ASSETSTRANSFERRED, "")
'  End If
'
'  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,A}"
'  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
'  rep.Out "{Arial=7,B}Assets transferred (cars, property, goods or other assets) " & vbCrLf
'  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
'  rep.Out "{x=" & cost & "}Cost / Market value"
'  rep.Out "{x=" & madeg - 1 & "}from which tax deducted"
'  rep.Out "{x=" & cash + 5 & "}Cash equivalent" & vbCrLf
'  rep.Out "{Times=10,bi}{x=25}{WBTEXTBOXR=30,2," & sDesc & "}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
'  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.12}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
'  rep.Out "{Arial=7,B}{wb}{x=71} - "
'  rep.Out "{Arial=7,B}{x=83} = "
'  rep.Out "{Arial=7}{x=5}Description of asset      " & vbCrLf
'
'' B Payments made
'  BenIndex = HasBenefit(BC_MISC, S_PAYMENTS)
'  If BenIndex > 0 Then
'    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_MISC, S_PAYMENTS)
'  End If
'
'  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,B}"
'  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
'  rep.Out "{Arial=7,B}Payments made on behalf of employee" & vbCrLf & "{Arial=6}" & vbCrLf
'
'  rep.Out "{Times=10,bi}{x=25}{WBTEXTBOXL=55,2," & sDesc & "}"
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.12}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
'  rep.Out "{Arial=7}{x=5}Description of payment " & vbCrLf & vbCrLf
'
'  BenIndex = HasBenefit(BC_MISC, S_PAYMENTS)
'  If BenIndex > 0 Then
'    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_MISC, S_NOTIONAL)
'  End If
'
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.12}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}"
'  rep.Out "{Arial=7,n}{wb}{x=5}Tax on notional payments not borne by employee within 30 days of receipt of each notional payment" & vbCrLf
'  rep.Out "{Arial=7}" & vbCrLf & "{Arial=6}" & vbCrLf
'
'' Credit Cards and Vouchers
'
'  BenIndex = HasBenefit(BC_CREDITSUBS, S_CREDIT)
'  If BenIndex > 0 Then
'    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_CREDITSUBS, S_CREDIT)
'  End If
'
'  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,C}"
'  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=30}•  {YREL=30}"
'  rep.Out "{Arial=7,B}Vouchers or credit cards" & vbCrLf & "{Arial=6}" & vbCrLf
'  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
'  rep.Out "{x=" & cost + 1 & "}Gross amount"
'  rep.Out "{x=" & madeg - 1 & "}from which tax deducted"
'  rep.Out "{x=" & cash + 5 & "}Cash equivalent" & vbCrLf
'
'  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
'  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.13}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
'  rep.Out "{Arial=7,nB}{wb}{x=71} - "
'  rep.Out "{Arial=7,B}{x=83} = "
'  rep.Out "{Arial=7}{x=5}Value of vouchers and payments made using credit cards or tokens"
'  rep.Out "{Arial=7}{YREL=40}" & vbCrLf & "{Arial=6}" & vbCrLf
'
''Accommodation
'  BenIndex = HasBenefit(BC_ACCOM, "")
'  If BenIndex > 0 Then
'    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_ACCOM, "")
'  End If
'
'  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,D}"
'  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
'  rep.Out "{Arial=7,B}Living Accommodation" & vbCrLf & "{x=" & cash + 5 & "}{Arial=6}Cash equivalent" & "{YREL=40}" & "{Arial=6}" & vbCrLf
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.14}{x=" & cash + 4 & "}{TIMES=10,BI}{WBTEXTBOXR=10,2," & lben & "}"
'  rep.Out "{Arial=7}{x=5}{wb}Cash equivalent of accommodation provided for employee, or his/ her family or household" & vbCrLf & vbCrLf
'  rep.Out "{Arial=7}" & vbCrLf & "{Arial=6}" & vbCrLf
'
''Mileage Allowance
'  BenIndex = HasBenefit(BC_EECAR, "")
'  If BenIndex > 0 Then
'    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_EECAR, "")
'  End If
'
'  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,E}"
'  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
'  rep.Out "{Arial=7,B}Mileage allowance" & vbCrLf & "{Arial=6,n}" & vbCrLf
'  rep.Out "{Arial=6}{x=" & madeg & "}{wb}Amount made good or" & vbCrLf
'  rep.Out "{x=" & cost + 1 & "}Gross amount"
'  rep.Out "{x=" & madeg - 1 & "}from which tax deducted"
'  rep.Out "{Arial=6}{x=90}Taxable payment" & vbCrLf
'  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.15}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" 'SP & vbcrlf
'  rep.Out "{Arial=7,B}{wb}{x=71} - "
'  rep.Out "{Arial=7,B}{x=83} = "
'  rep.Out "{Arial=7}{x=5}Car and mileage allowances paid for employee's car" & vbCrLf
'
' 'Car and fuel benefits
'  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,F}"
'  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
'  rep.Out "{Arial=7,B}Cars and car fuel" & vbCrLf & vbCrLf
'  rep.Out "{Arial=7,i}{x=5}If more than two cars were made available, either at the same time or in succession, please give details on a separate sheet" & vbCrLf & vbCrLf
'  rep.Out "{Arial=6}{x=45}Car 1"
'  rep.Out "{Arial=6}{x=75}Car 2" & vbCrLf
'  rep.Out "{TIMES=10,bi}{x=40}{WBTEXTBOXR=25,2,SBNP-Make1}"
'  rep.Out "{Times=10,bi}{x=73}{WBTEXTBOXR=25,2,SBNP-Make2}" & vbCrLf
'
'  rep.Out "{Arial=7}{x=5}Make and Model" & vbCrLf & "{Arial=6}" & vbCrLf & vbCrLf
'      rep.Out "{Times=10,bi}{x=40}{x=40}{WBTEXTBOXR=10,2,SBNP-RegDate1}"
'      'dycar.FindNext criteria
'      rep.Out "{Times=10,bi}{x=73}{WBTEXTBOXR=10,2,SBNP-RegDate2}" & vbCrLf
'    'Case Else
'  'End Select
'  rep.Out "{Arial=7}{x=5}Dates first registered" & vbCrLf & "{Arial=6}" & vbCrLf
'
'  rep.Out "{Arial=6}{x=40}From"
'  rep.Out "{Arial=6}{x=52}To"
'  rep.Out "{Arial=6}{x=73}From"
'  rep.Out "{Arial=6}{x=85}To" & vbCrLf
'
'  rep.Out "{Times=10,bi}{x=40}{WBTEXTBOXR=10,2,SBNP-From1}"
'  rep.Out "{Times=10,bi}{x=52}{WBTEXTBOXR=10,2,SBNP-To1}"
'  rep.Out "{Times=10,bi}{x=73}{WBTEXTBOXR=10,2,SBNP-From2}"
'  rep.Out "{Times=10,bi}{x=85}{WBTEXTBOXR=10,2,SBNP-To2}" & vbCrLf
'  rep.Out "{Arial=7}{x=5}Dates car was available" & vbCrLf & "{Arial=6}" & vbCrLf
'
'  rep.Out "{Arial=6}" & vbCrLf
'  rep.Out "{Arial=6}{x=38}2,499 or less"
'  rep.Out "{Arial=6}{x=46}2,500 to 17,999"
'  rep.Out "{Arial=6}{x=54}18,000 or more"
'  rep.Out "{Arial=6}{x=72}2,499 or less"
'  rep.Out "{Arial=6}{x=79}2,500 to 17,999"
'  rep.Out "{Arial=6}{x=87}18,000 or more" & vbCrLf
'
'  rep.Out "{Arial=7}{x=40}{BOX=3,2}"
'  rep.Out "{x=48}{BOX=3,2}"
'  rep.Out "{x=56}{BOX=3,2}"
'  rep.Out "{x=73}{BOX=3,2}"
'  rep.Out "{x=81}{BOX=3,2}"
'  rep.Out "{x=89}{BOX=3,2}"
'
'  Select Case BusMiles1
'    Case "18000 PLUS"
'      rep.Out "{Wingdings=12,b}{x=56}{WBTEXTBOXL=3,2,ü}"
'    Case "2500 TO 17999"
'      rep.Out "{Wingdings=12,b}{x=48}{WBTEXTBOXL=3,2,ü}"
'    Case "0 TO 2499"
'      rep.Out "{Wingdings=12,b}{x=40}{WBTEXTBOXL=3,2,ü}"
'    Case Else
'      Call ECASE("Unknown mileage band")
'  End Select
'
'  Select Case BusMiles2
'    Case "18000 PLUS"
'      rep.Out "{Wingdings=12,b}{x=89}{WBTEXTBOXL=3,2,ü}"
'    Case "2500 TO 17999"
'      rep.Out "{Wingdings=12,b}{x=81}{WBTEXTBOXL=3,2,ü}"
'    Case "0 TO 2499"
'      rep.Out "{Wingdings=12,b}{x=73}{WBTEXTBOXL=3,2,ü}"
'    Case Else
'      Call ECASE("Unknown mileage band")
'  End Select
'
'  rep.Out vbCrLf
'  rep.Out "{Arial=7}{x=5}Business mileage used in calculation" & vbCrLf
'  rep.Out "{Arial=7}{x=5}for this car." & "{Arial=7,i}Tick only one box for each car." & vbCrLf
'  rep.Out "{Arial=6,i}{x=5}If the car was not available for part of the year, the business mileage limits are reduced proportionately." & vbCrLf
'  rep.Out "{Arial=6}" & vbCrLf
'  rep.Out "{Arial=6}{x=40}Engine size in cc"
'  rep.Out "{Arial=6}{x=56}Petrol"
'  rep.Out "{Arial=6}{x=60}Diesel"
'  rep.Out "{Arial=6}{x=73}Engine size in cc"
'  rep.Out "{Arial=6}{x=85}Petrol"
'  rep.Out "{Arial=6}{x=90}Diesel"
'  rep.Out "{Arial=7}{x=5}Tick type of fuel and enter engine size" & vbCrLf
'  rep.Out "{Arial=7,i}{x=5}if there is a car fuel scale charge"
'  rep.Out "{Arial=7}{x=40}{WBTEXTBOXR=10,2,SBNP-EngSize1}"
'  rep.Out "{Arial=7}{x=73}{WBTEXTBOXR=10,2,SBNP-EngSize2}"
'    If Diesel1 = False Then
'      rep.Out "{x=56}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
'      rep.Out "{Arial=7}{x=60}{BOX=3,2}"
'    Else
'      rep.Out "{Arial=7}{x=56}{BOX=3,2}"
'      rep.Out "{x=60}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
'    End If
'    If Diesel2 = False Then
'      rep.Out "{x=85}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
'      rep.Out "{Arial=7}{x=90}{BOX=3,2}"
'    Else
'      rep.Out "{Arial=7}{x=85}{BOX=3,2}"
'      rep.Out "{x=90}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
'    End If
'  rep.Out vbCrLf
'  rep.Out "{Arial=6,b}{x=48}cc"
'  rep.Out "{Arial=6,b}{x=81}cc"
'  rep.Out "{Arial=7,n}" & vbCrLf & vbCrLf
'  rep.Out "{Arial=6}{x=74}Car 1"
'  rep.Out "{Arial=6}{x=90}Car 2" & vbCrLf
'  'list price
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-Acc1}"
'  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-Acc2}" & vbCrLf
'
'  rep.Out "{Arial=7,N}{x=5}Price of optional accessories fitted when the car was first made available to the employee" & vbCrLf & vbCrLf
'
''ACCESSORIES 2
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-AccAfter1}"
'  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-AccAfter2}" & vbCrLf
'  rep.Out "{Arial=7}{x=5}Price of accessories added after the car was first made available to the employee" & vbCrLf & vbCrLf
'
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-CapCon1}"
'  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-CapCon2}" & vbCrLf
'  rep.Out "{x=0}{Arial=7}{x=5}Capital contributions (maximum £5,000) the employee made towards the cost of car or accessories" & vbCrLf & vbCrLf
'
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-Amt1}"
'  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-Amt2}" & vbCrLf
'  rep.Out "{Arial=7}{x=5}Amount paid by employee for private use of the car" & vbCrLf & vbCrLf
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-CashEquiv1}"
'  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-CashEquiv2}" & vbCrLf
'  rep.Out "{Arial=7,n}{x=5}Cash equivalent of each car" & vbCrLf & vbCrLf
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.16}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2,SBNP-CarCount}" & vbCrLf
'
'  rep.Out "{Arial=7,n}{x=5}Total cash equivalent of all cars available in  SBNP-STR_FORM_YEAR." & vbCrLf & vbCrLf
'
'  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-FuelEquiv1}"
'  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-Fuelquiv2}" & vbCrLf
'  rep.Out "{Arial=7,N}{x=5}Cash equivalent of fuel for each car" & vbCrLf & vbCrLf
'  rep.Out "{Arial=7,n}{x=5}Total cash equivalent of fuel for all cars available in  SBNP-STR_FORM_YEAR. "
'  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.17}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2,SBNP-BenFuel}"
'
'Page1_End:
'    Call xReturn("page1")
'    Exit Sub
'
'Page1_Err:
'    Call ErrorMessage(ERR_ERROR, Err, "page1", "ERR_HMIT_PAGE1", "Error printing page 1 of the HMIT return.")
'    Resume Page1_End
End Sub

Public Sub HMITPage2(rep As Reporter)
  Dim BenIndex As Long
  Dim sDesc, lval, lmg, lben
  On Error GoTo page2_Err
  Call xSet("page2")
  
  Dim Miras1 As Boolean
  Dim Miras2 As Boolean
  Dim Shares As Boolean
  Dim EnterTick As Boolean
  rep.Out "{NEWPAGE}"
'Vans
  rep.Out "{Arial=7,i}" & vbCrLf & vbCrLf
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,G}"
  rep.Out "{Arial=11,NB}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Vans " & vbCrLf
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.18}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2,SBNP-VanTot}" & vbCrLf
  rep.Out "{Arial=7,n}{wb}{x=5}Cash equivalent of all vans made available for private use"
  rep.Out "{Arial=7}" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,H}"
  rep.Out "{Arial=11,nB}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Interest-free and low interest loans" & vbCrLf
  rep.Out "{Arial=7,i}{x=5}If the total amount outstanding on all loans does not exceed £5,000 at any time in the year, there is no need for details in this section." & vbCrLf & vbCrLf
  rep.Out "{Arial=6,b}{x=" & madeg + 1 & "}Loan 1"
  rep.Out "{Arial=6,b}{x=" & cash + 5 & "}Loan 2" & vbCrLf
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=5,2,SBNP-LoanPur1}"
  rep.Out "{x=" & cash + 4 & "}{WBTEXTBOXR=5,2,SBNP-LoanPur2}" & vbCrLf
  rep.Out "{Arial=7,n}{x=5}Purpose of loan(s) - {Arial=7,i}please use the code shown in P11D Guide for employers" & "{Arial=6}" & vbCrLf & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=5,2,SBNP-LoanBorr1}"
  rep.Out "{x=" & cash + 4 & "}{WBTEXTBOXR=5,2,SBNP-LoanBorr2}" & vbCrLf
  rep.Out "{Arial=7,n}{x=5}Number of joint borrowers {Arial=7,i}(if applicable)" & "{Arial=6}" & vbCrLf & vbCrLf
  
  If Miras1 = False Then
    rep.Out "{x=" & madeg & "}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
  Else
    rep.Out "{Arial=7}{x=" & madeg & "}{BOX=3,2}"
  End If
  
  If Miras2 = False Then
    rep.Out "{x=" & cash + 4 & "}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
  Else
    rep.Out "{Arial=7}{x=" & cash + 4 & "}{BOX=3,2}"
  End If
  rep.Out vbCrLf
  rep.Out "{Arial=7}{x=5}Tick the box if the loan is within MIRAS" & "{Arial=6}" & vbCrLf & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-LoanOpBal1}"
  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-LoanOpBal2}" & vbCrLf
  rep.Out "{Arial=7,n}{x=5}Amount outstanding at 5 April SBNP-1STYEAR or at date loan was made if later" & "{Arial=6}" & vbCrLf & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-LoanClB1}"
  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-LoanClB2}" & vbCrLf
    
  rep.Out "{Arial=7,n}{x=5}Amount outstanding at 5 April SBNP-2NDYEAR or at date loan was discharged if earlier" & "{Arial=6}" & vbCrLf & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-LoanMax1}"
  rep.Out "{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-LoanMax2}" & vbCrLf
  rep.Out "{Arial=7,n}{x=5}Maximum amount outstanding at any time in the year" & "{Arial=6}" & vbCrLf & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-TotInt1}"
  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-TotInt2}" & vbCrLf
  rep.Out "{Arial=7}{x=5}Total amount of interest paid by the borrower in the year to 5 April SBNP-2NDYEAR"
  rep.Out "{Arial=6,i} - enter ""NIL"" if none was paid" & vbCrLf & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-LoanDate1}"
  rep.Out "{Times=10,bi}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2,SBNP-LoanDate2}" & vbCrLf
  rep.Out "{Arial=7}{x=5}Date loan was made or discharged in the year to 5 April SBNP-2NDYEAR if applicable" & "{Arial=6}" & vbCrLf & vbCrLf
  
  rep.Out "{Arial=10,nb}{x=" & madeg - 4 & "}{BWTEXTBOX=4,2,1.19}{x=" & madeg & "}{Times=10,bi}{WBTEXTBOXR=10,2,SBNP-LoansCashEqiv1}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.19}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2,SBNP-LoansCashEqiv2}" & vbCrLf
  
  rep.Out "{Arial=7}{wb}{x=5}Cash equivalent of loans after deducting any interest paid by the borrower" & vbCrLf & "{Arial=6}" & vbCrLf & vbCrLf
  
'Mobile Phones
  BenIndex = HasBenefit(BC_PHONE, S_MOBILE)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_PHONE, S_MOBILE)
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,I}"
  rep.Out "{Arial=11,B}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Mobile telephones" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.20}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7}{wb}{x=5}Cash equivalent of mobile telephones provided"
  rep.Out "{Arial=7,n}" & vbCrLf & "{Arial=6}" & vbCrLf
  
'medical
  BenIndex = HasBenefit(BC_MEDICAL, S_MEDICAL)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_MEDICAL, S_MEDICAL)
    Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,J}"
  rep.Out "{Arial=11,B}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Private medical treatment or insurance" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
  rep.Out "{Arial=6}{x=" & cost + 2 & "}Cost to you"
  rep.Out "{Arial=6}{x=" & madeg - 1 & "}from which tax deducted"
  rep.Out "{Arial=6}{x=90}Cash equivalent" & vbCrLf
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.21}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Private medical treatment or insurance"
  rep.Out "{Arial=7,n}" & vbCrLf & "{Arial=6}" & vbCrLf

'relocation
  BenIndex = HasBenefit(BC_RELOC, "")
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_RELOC, "")
  End If
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,K}"
  rep.Out "{Arial=11,B}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Qualifying relocation expenses payments and benefits (non-qualifying expenses go in P below)" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7}{wb}{x=5}Excess over £8,000 of all qualifying relocation expenses payments and benefits for each move"
  rep.Out "{Arial=7,n}" & vbCrLf & vbCrLf
  
'Services supplied
  BenIndex = HasBenefit(BC_SERVICESPROVIDED, "")
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_SERVICESPROVIDED, "")
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If

  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,L}"
  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Services supplied" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
  rep.Out "{Arial=6}{x=" & cost + 2 & "}Cost to you"
  rep.Out "{Arial=6}{x=" & madeg - 1 & "}from which tax deducted"
  rep.Out "{Arial=6}{x=90}Cash equivalent" & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Services supplied to the employee"
  rep.Out "{Arial=7,n}" & vbCrLf & "{Arial=6}" & vbCrLf

'Assets at disposal
  BenIndex = HasBenefit(BC_ASSETSATDISPOSAL, "")
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_ASSETSATDISPOSAL, "")
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,M}"
  rep.Out "{Arial=11,B}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Assets placed at the employee's disposal" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=6}{x=" & cost & "}Annual value plus"
  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
  rep.Out "{Arial=6}{x=" & cost & "}expenses incurred"
  rep.Out "{Arial=6}{x=" & madeg - 1 & "}from which tax deducted"
  rep.Out "{Arial=6}{x=90}Cash equivalent" & vbCrLf
  rep.Out "{Times=10,bi}{x=25}{WBTEXTBOXL=28,2," & sDesc & "}"
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Description of asset"
  rep.Out "{Arial=7,n}" & vbCrLf & "{Arial=6}" & vbCrLf

'Shares
  BenIndex = HasBenefit(BC_MISC, S_SHARES)
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,N}"
  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Shares" & vbCrLf & "{Arial=6}" & vbCrLf
  
  If BenIndex > 0 Then
    rep.Out "{x=" & madeg & "}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
  Else
    rep.Out "{Arial=7}{x=" & madeg & "}{BOX=3,2}"
  End If
  rep.Out vbCrLf
  rep.Out "{Arial=7}{x=5}Tick the box if during the year there have been share-related benefits for the employee"
  rep.Out "{Arial=7}" & vbCrLf & "{Arial=6}" & vbCrLf

'Other Items
  
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,O}"
  rep.Out "{Arial=11,B}{x=5}{WB}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Other items" & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
  rep.Out "{Arial=6}{x=" & cost + 2 & "}Cost to you"
  rep.Out "{Arial=6}{x=" & madeg - 1 & "}from which tax deducted"
  rep.Out "{Arial=6}{x=90}Cash equivalent" & vbCrLf
  
  BenIndex = HasBenefit(BC_CREDITSUBS, S_SUBSCRIPTION)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_CREDITSUBS, S_SUBSCRIPTION)
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Subscriptions and professional fees"
  rep.Out "{Arial=7,n}" & "{Arial=6}" & vbCrLf & vbCrLf
  
  BenIndex = HasBenefit(BC_CREDITSUBS, S_SUBSCRIPTION)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_CREDITSUBS, S_SUBSCRIPTION)
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Times=10,bi}{x=25}{WBTEXTBOXR=28,2,SBNP-OthItem}"
  rep.Out "{x=" & cost & "}{WBTEXTBOXR=10,2,SBNP-OthCost}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-OthMG}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2,SBNP-OthCashEquiv}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Description of other items"
  
'Income Tax Paid
  BenIndex = HasBenefit(BC_MISC, S_TAXPAID)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_MISC, S_TAXPAID)
    Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Arial=6}{x=90}Tax paid" & vbCrLf
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7}{wb}{x=5}Income tax paid but not deducted from director's remuneration"
  rep.Out "{Arial=7,n}" & vbCrLf & vbCrLf & vbCrLf
'Expenses
  BenIndex = HasBenefit(BC_GENERAL, S_TRAVEL)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_GENERAL, S_TRAVEL)
    Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Arial=10,nb}{X=0}{BWTEXTBOX=3,2,P}"
  rep.Out "{Arial=11,B}{x=5}{wb}{YREL=-30}•  {YREL=30}"
  rep.Out "{Arial=7,b}Expenses payments made to, or on behalf of, the employee " & vbCrLf & "{Arial=6}" & vbCrLf
  rep.Out "{Arial=6}{x=" & madeg & "}Amount made good or" & vbCrLf
  rep.Out "{Arial=6}{x=" & cost + 2 & "}Cost to you"
  rep.Out "{Arial=6}{x=" & madeg - 1 & "}from which tax deducted"
  rep.Out "{Arial=6}{x=90}Taxable payment" & vbCrLf
  
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.22}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,nb}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Travelling and subsistence payments"
  rep.Out "{Arial=7,n}" & vbCrLf & "{Arial=6}" & vbCrLf
  
'Entertainment
  BenIndex = HasBenefit(BC_GENERAL, S_ENTERTAINMENT)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_GENERAL, S_ENTERTAINMENT)
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If

  If m_Parent.CT Then
    rep.Out "{x=" & cost - 5 & "}{Wingdings=12,b}{WBTEXTBOXL=3,2,ü}"
  Else
    rep.Out "{Arial=7}{x=" & cost - 5 & "}{BOX=3,2}"
  End If
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.23}{x=" & cash + 4 & "}{times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Entertainment {Arial=7,i}(trading organisations read P11D Guide and then"
  rep.Out "{Arial=7,n}" & vbCrLf
  rep.Out "{Arial=7,i}{x=5}enter a tick or a cross as appropriate here)" & vbCrLf
  
'General Expenses
  BenIndex = HasBenefit(BC_GENERAL, S_GENERAL)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_GENERAL, S_GENERAL)
   Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.23}{x=" & cash + 4 & "}{Times=10,bi}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,nb}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}General expenses allowance for business travel"
  rep.Out "{Arial=6,n}" & vbCrLf & vbCrLf

'Home Phone
  BenIndex = HasBenefit(BC_PHONE, S_HOME)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_PHONE, S_HOME)
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Times=10,bi}{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.23}{x=" & cash + 4 & "}{Times=10,BI}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Payments for use of home telephone"
  rep.Out "{Arial=6,n}" & vbCrLf & vbCrLf
  
'Non-qual relocation
  rep.Out "{Times=10,BI}{x=" & cost & "}{WBTEXTBOXR=10,2,SBNP-NonQCost}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2,SBNP-NonQMG}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.23}{x=" & cash + 4 & "}{Times=10,BI}{WBTEXTBOXR=10,2,SBNP-NonQTP}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Non-qualifying relocation expenses {Arial=6,i}(those not shown in section K)"
  rep.Out "{Arial=7}" & vbCrLf & "{Arial=6}" & vbCrLf

'Other Expenses
  BenIndex = HasBenefit(BC_MISC, S_CHAUFFEUR)
  If BenIndex > 0 Then
    Call SumBenefit(sDesc, lval, lmg, lben, BenIndex, BC_MISC, S_CHAUFFEUR)
  Else
    sDesc = Empty
    lval = Empty
    lmg = Empty
    lben = Empty
  End If
  rep.Out "{Times=10,bi}{x=25}{WBTEXTBOXL=28,2," & sDesc & "}"
  rep.Out "{x=" & cost & "}{WBTEXTBOXR=10,2," & lval & "}"
  rep.Out "{x=" & madeg & "}{WBTEXTBOXR=10,2," & lmg & "}"
  rep.Out "{Arial=10,nb}{x=" & cash & "}{BWTEXTBOX=4,2,1.23}{x=" & cash + 4 & "}{WBTEXTBOXR=10,2," & lben & "}" & vbCrLf
  rep.Out "{Arial=7,b}{wb}{x=71} - "
  rep.Out "{Arial=7,b}{x=83} = "
  rep.Out "{Arial=7}{x=5}Description of other expenses"
  rep.Out "{Arial=8,n}" & vbCrLf

page2_End:
  Call xReturn("page2")
  Exit Sub

page2_Err:
  Call ErrorMessage(ERR_ERROR, Err, "page2", "ERR_UNDEFINED", "Undefined error.")
  Resume page2_End
End Sub
Private Function HasBenefit(BenefitClass As benClass, SpecificBenefit As String) As Long
  Dim i As Long
  Dim ben As IBenefitClass
  
  For i = 1 To benefits.count
    Set ben = benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BenefitClass Then
        If (Len(SpecificBenefit) > 0) And (ben.BenefitSubClass > 0) Then
          If StrComp(ben.GetItem(ben.BenefitSubClass), SpecificBenefit, vbBinaryCompare) <> 0 Then
            GoTo next_ben
          End If
        End If
        HasBenefit = i
        Exit Function
      End If
    End If
next_ben:
  Next i
  HasBenefit = 0
End Function
Private Function SumBenefit(ByRef sDescription As Variant, ByRef value As Variant, ByRef MadeGood As Variant, ByRef benefit As Variant, IndexStart As Long, BenefitClass As benClass, SpecificBenefit As String) As Boolean
  Dim lval As Long, lmg As Long, lben As Long
  Dim s As String
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo SumBenefit_Err
  Call xSet("SumBenefit")
  
  For i = IndexStart To benefits.count
    Set ben = benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BenefitClass Then
        If (Len(SpecificBenefit) > 0) And (ben.BenefitSubClass > 0) Then
          If StrComp(ben.GetItem(ben.BenefitSubClass), SpecificBenefit, vbBinaryCompare) <> 0 Then
            GoTo next_ben
          End If
        End If
        lval = lval + ben.GetItem(ITEM_VALUE)
        lmg = lmg + ben.GetItem(ITEM_MADEGOOD)
        lben = lben + ben.GetItem(ITEM_BENEFIT)
        s = s & ben.GetItem(ITEM_DESC) & " "
      End If
    End If
next_ben:
  Next i
  If Len(s) > 0 Then
    sDescription = s
  Else
    sDescription = ""
  End If
  value = lval
  MadeGood = lmg
  benefit = lben
  
SumBenefit_End:
  Call xReturn("SumBenefit")
  Exit Function
  
SumBenefit_Err:
  SumBenefit = False
  sDescription = S_ERROR
  value = S_ERROR
  MadeGood = S_ERROR
  benefit = S_ERROR
  Resume SumBenefit_End
End Function
Private Property Let IBenefitClass_BenefitClass(NewValue As benClass)
  m_bentype = NewValue
End Property
Private Property Get IBenefitClass_BenefitClass() As benClass
  IBenefitClass_BenefitClass = m_bentype
End Property
Private Property Get IBenefitClass_BenefitSubClass() As Long
  IBenefitClass_BenefitSubClass = 0
End Property
Private Function IBenefitClass_Calculate() As Variant
  ECASE ("Calculate")
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE ("CompanyDefined")
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
  
End Property
Public Property Get DateEdit() As Date
  DateEdit = m_dateedit
End Property
Public Property Let DateEdit(NewValue As Date)
  m_dateedit = NewValue
End Property
Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo DeleteDB_Err
  
  Call xSet("DeleteDB")
  
  Set rs = Me.Parent.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEE, PersonnelNumber), dbOpenDynaset)
  If rs.BOF And rs.EOF Then GoTo DeleteDB_End
  rs.MoveFirst
  rs.Delete
  Set rs = Nothing
  For i = 1 To benefits.count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      Call ben.DeleteDB
      benefits(i).DeleteDB
    End If
  Next i
  benefits.RemoveAll
  
  Set rs = Me.Parent.db.OpenRecordset(sql.Queries(SELECT_AUDIT), dbOpenDynaset)
  rs.AddNew
  rs.Fields("Action") = "Delete"
  rs.Fields(S_EMPLOYEE_NUM_FIELD) = Me.PersonnelNumber
  rs.Fields("UserName") = P11d32.UserName
  rs.Fields("ActionDate") = Now()
  rs.Update
  

DeleteDB_End:
  Call xReturn("DeleteDB")
  Exit Function

DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DeleteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume DeleteDB_End

End Function
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = NewValue
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property
Private Function IBenefitClass_GetItem(ByVal Item As Long) As Variant
  IBenefitClass_GetItem = m_EmployeeItems(Item)
End Function
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
  Call Kill
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
Private Function IBenefitClass_PrintWk(rep As TCSREP.Reporter) As Boolean
  ECASE ("PrintWk")
End Function
Public Sub Kill()
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
Public Property Get PersonnelNumber() As String
  PersonnelNumber = m_EmployeeItems(ee_PersonnelNumber)
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
    Set rs = P11d32.CurrentEmployer.EmployeesRecordSet
    If rs Is Nothing Then
      Set rs = P11d32.CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEE, m_EmployeeItems(ee_PersonnelNumber)), dbOpenDynaset, dbFailOnError)
    End If
    With rs
      Call ben.SetItem(ee_Surname, "" & .Fields("Surname"))
      Call ben.SetItem(ee_Initials, "" & .Fields("Initials"))
      Call ben.SetItem(ee_Firstname, "" & .Fields("Firstname"))
      Call ben.SetItem(ee_Title, "" & .Fields("Title"))
      Call ben.SetItem(ee_Salutation, "" & .Fields("Salutation"))
      Call ben.SetItem(ee_Email, "" & .Fields("Email"))
      Call ben.SetItem(ee_PersonnelNumber, "" & .Fields(S_EMPLOYEE_NUM_FIELD))
      Call ben.SetItem(ee_Group1, "" & .Fields("Group1"))
      Call ben.SetItem(ee_Group2, "" & .Fields("Group2"))
      Call ben.SetItem(ee_Group3, "" & .Fields("Group3"))
      Call ben.SetItem(ee_NINumber, "" & .Fields("NI"))
      Call ben.SetItem(ee_joined, "" & .Fields("Joined"))
      Call ben.SetItem(ee_LEFT, "" & .Fields("Left"))
      Call ben.SetItem(ee_Comments, "" & .Fields("Comments"))
      Call ben.SetItem(ee_Status, IIf(.Fields("Status") = "Director", True, False))
      'van details
      Call ben.SetItem(ee_RelevantDaysForDailySharedVanCalc, .Fields("RELEVANTDAYS"))
      Call ben.SetItem(ee_PaymentsForPrivateUseOfSharedVans, .Fields("PRIVATECONTRIB"))
      Call ben.SetItem(ee_OneOrMoreSharedVanAvailable, .Fields("TWOPLUSVANS"))
      Call ben.SetItem(ee_ReportyDailyCalculationOfSharedVans, .Fields("DAILYCALC"))
      Call ben.SetItem(ee_NonSharedVanAvailableAtSameTimeAsSharedVan, .Fields("TwoPlusVans"))
  
      m_dateedit = IIf(IsNull(.Fields("DateEdit")), UNDATED, .Fields("DateEdit"))
      'need to initialise set employee setitem
      m_EmployeeItems(ee_OneOrMoreSharedVanAvailable) = False
      Call ben.SetItem(ee_OneOrMoreSharedVanAvailable, IIf(.Fields("SHAREDVANS"), True, False))
      ben.BenefitClass = BC_EMPLOYEE
      ben.ReadFromDB = True
      
    End With
    Set ben = Nothing
    IBenefitClass_ReadDB = True
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
  m_ReadFromDB = NewValue 'ECASE ("ReadFromDB")
End Property
Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
  'ECASE ("ReadFromDB")
End Property
Private Property Get IBenefitClass_Reference() As String
  ECASE ("Reference")
End Property
Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  ECASE ("RSBookMark")
End Property
Private Property Get IBenefitClass_RSBookMark() As String
  ECASE ("RSBookMark")
End Property
Private Function IBenefitClass_SetItem(ByVal Item As Long, value As Variant) As Boolean
  If Item = ee_OneOrMoreSharedVanAvailable Then
    With P11d32.CurrentEmployer
      If value Then
        If m_EmployeeItems(ee_OneOrMoreSharedVanAvailable) <> value Then
          .EmployeesWithSharedVan = .EmployeesWithSharedVan + 1
        End If
      Else
        If m_EmployeeItems(ee_OneOrMoreSharedVanAvailable) <> value Then
          .EmployeesWithSharedVan = .EmployeesWithSharedVan - 1
        End If
      End If
      m_EmployeeItems(ee_OneOrMoreSharedVanAvailable) = value
    End With
  Else
    m_EmployeeItems(Item) = value
  End If
End Function
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
  
  Call MDIMain.sts.SetStatus(0, "Writing benefits and employee information to Database")
  If ben.Dirty Then
    If Not MultiUserCheck Then
      IBenefitClass_WriteDB = P11d32.CurrentEmployer.ReloadCurrentEmployee
      GoTo WriteDB_End
    ElseIf Not ben.InvalidFields Then
      Set rs = Me.Parent.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEE, PersonnelNumber), dbOpenDynaset)
      If rs.EOF And rs.BOF Then
        rs.AddNew
        rs.Fields(S_EMPLOYEE_NUM_FIELD) = PersonnelNumber
        rs.Fields("DateEdit") = Now
        DateEdit = rs.Fields("DateEdit")
      Else
        rs.Edit
      End If
      With ben
        rs.Fields("Name") = ee.FullName
        rs.Fields("Surname") = .GetItem(ee_Surname)
        rs.Fields("Title") = .GetItem(ee_Title)
        rs.Fields("Initials") = .GetItem(ee_Initials)
        rs.Fields("Firstname") = .GetItem(ee_Firstname)
        rs.Fields("Salutation") = .GetItem(ee_Salutation)
        rs.Fields("Email") = .GetItem(ee_Email)
        rs.Fields("Group1") = .GetItem(ee_Group1)
        rs.Fields("Group2") = .GetItem(ee_Group2)
        rs.Fields("Group3") = .GetItem(ee_Group3)
        rs.Fields("NI") = .GetItem(ee_NINumber)
        rs.Fields("Joined") = TryConvertDateDMY(.GetItem(ee_joined), UNDATED)
        rs.Fields("Left") = TryConvertDateDMY(.GetItem(ee_LEFT), UNDATED) 'IIf(dDate = UNDATED, "", dDate)
        rs.Fields("Comments") = .GetItem(ee_Comments)
        rs.Fields("Status") = .GetItem(ee_Status)
        rs.Fields("SHAREDVANS") = .GetItem(ee_OneOrMoreSharedVanAvailable)
        rs.Fields("RELEVANTDAYS") = .GetItem(ee_RelevantDaysForDailySharedVanCalc)
        rs.Fields("PRIVATECONTRIB") = .GetItem(ee_PaymentsForPrivateUseOfSharedVans)
        rs.Fields("TWOPLUSVANS") = .GetItem(ee_NonSharedVanAvailableAtSameTimeAsSharedVan)
        rs.Fields("DAILYCALC") = .GetItem(ee_ReportyDailyCalculationOfSharedVans)
        rs.Fields("DateEdit") = Now
        ee.DateEdit = rs.Fields("DateEdit")
        ben.ReadFromDB = True
        rs.Update
        Dirty = False
      End With
    Else
      bWriteOK = False
    End If
  End If
  
  If bWriteOK Then
    bWriteOK = WriteBenefits
  Else
    Call WriteBenefits
  End If
  
  If bWriteOK Then MDIMain.ClearConfirmUndo
  IBenefitClass_WriteDB = bWriteOK
  
WriteDB_End:
  Call MDIMain.sts.SetStatus(0, "")
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


Public Function NonSharedVans() As IBenefitClass

  On Error GoTo NonSharedVans_Err
  Call xSet("NonSharedVans")

  Set NonSharedVans = benefits(L_VANS_BENINDEX)

NonSharedVans_End:
  Call xReturn("NonSharedVans")
  Exit Function

NonSharedVans_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVans", "ERR_UNDEFINED", "Undefined error.")
  Resume NonSharedVans_End
  Resume
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
  
  If ben.ReadFromDB Then
    Set rs = P11d32.CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_EMPLOYEECHECK, ben.GetItem(ee_PersonnelNumber)), dbOpenDynaset, dbFailOnError)
    If Not (rs.EOF And rs.BOF) Then
      d = IIf(IsNull(rs.Fields("DateEdit")), UNDATED, rs.Fields("DateEdit"))
      If d <> DateEdit Then
        sMsg = "The employee you are trying to update has been changed by " & IIf(Len(rs.Fields("UserEdit") & ""), rs.Fields("UserEdit"), "UNKNOWN") & _
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
  Call ErrorMessage(ERR_ERROR, Err, "MultiUserCheck", "ERR_UNDEFINED", "Undefined error.")
  Resume MultiUserCheck_End
  Resume
End Function


Public Function WriteBenefits() As Boolean
  Dim i As Long, b As Boolean
  Dim ben As IBenefitClass
  On Error GoTo WriteBenefits_Err
  Call xSet("WriteBenefits")

  b = True
  
  For i = 1 To benefits.count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If ben.Dirty Then
        If ben.InvalidFields = 0 Or IsBenefitACollection(ben) Then
          If b = False Then
            Call ben.WriteDB
          Else
            b = ben.WriteDB
          End If
        Else
          b = False
        End If
      End If
    End If
  Next

  WriteBenefits = b
  
WriteBenefits_End:
  Call xReturn("WriteBenefits")
  Exit Function

WriteBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteBenefits", "ERR_UNDEFINED", "Undefined error.")
  Resume WriteBenefits_End
  Resume
End Function


Public Function IsBenefitACollection(ben As IBenefitClass) As Boolean
End Function



