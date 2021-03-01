VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Importing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IProcessImport
Implements IErrorFilter

Private m_dbP11dFile As Database
Private m_dbError As Database
Private m_rsError As Recordset
Public Employer As Employer

Private m_iw As ImportWizard
Private mSql As SQLQUERIES_IMP
Private m_TrackingPath As String
Public Tracking As Boolean

Private Sub Class_Initialize()
  Set mSql = New SQLQUERIES_IMP
End Sub

Private Sub Class_Terminate()
  Call KillImport(Nothing)
  Set mSql = Nothing
End Sub
Public Property Let TrackingPath(ByVal NewValue As String)
  m_TrackingPath = NewValue
End Property
Public Property Get TrackingPath() As String
  TrackingPath = m_TrackingPath
End Property
Private Sub KillImport(ic As ImportClass)
  Dim ben As IBenefitClass
  
  On Error Resume Next
  If Not ic Is Nothing Then ic.KillImporter
  Set m_iw = Nothing
  Set ic = Nothing
  
  Set ben = Employer
  Set m_rsError = Nothing
  Set m_dbError = Nothing
  Set m_dbP11dFile = Nothing
  If Not Employer Is Nothing Then
    ben.Kill
    Set Employer = Nothing
    Set ben = Nothing
  End If
  Set p11d32.CurrentEmployer = Nothing
End Sub
  
Private Property Get ImportErrorFile() As String
  ImportErrorFile = FullPath(AppPath) & "Import32.err"
End Property
Public Function PreImport(ByVal bExclusive As Boolean) As Employer
  Dim ey As Employer
  Dim ben As IBenefitClass
  
On Error GoTo err_err

  If F_Employers.LB.SelectedItem Is Nothing Then Call Err.Raise(ERR_EMPLOYER_INVALID, "PreImport", "There is no employer selected")
  
  Set ey = p11d32.Employers.Item(F_Employers.LB.SelectedItem.Tag)
  Set ben = ey
  Call ben.Kill
  
  Call ey.OpenEmployer(bExclusive)
  
  Set PreImport = ey

err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "PreImport"), "Failed to pre init employer for import")
  Resume
End Function
Public Sub Zip(ByVal ben As IBenefitClass, sDestFilename As String)
  Dim Zip As cZip
  Call ben.Kill
  Set Zip = New cZip
  Zip.BasePath = p11d32.WorkingDirectory
  Zip.AddFileSpec (ben.value(employer_FileName))
  Zip.RecurseSubDirs = False
  Zip.ZipFile = sDestFilename
  Call Zip.Zip
End Sub
Public Sub InitImport()
  Dim i As Long
  Dim ic As ImportClass
  Dim wsError As Workspace
  Dim ben As IBenefitClass
  Dim ifb As IBenefitForm2
  Dim sGuid As String
  Dim sTrackingDescription As String
  Dim sSourceFileName As String, sDestFilename As String
  Dim sTrackingMessage As String
  Dim dt As Date
  Dim Zip As cZip
  
  Const L_MAX_TRACKING_DESCRIPTION = 40
  Const S_TRACKING_MESSAGE As String = "Please enter a " & L_MAX_TRACKING_DESCRIPTION & " character description of the import." & vbCrLf & vbCrLf & "Pleae don't use any special characters like \/:" & vbCrLf & vbCrLf & "To bypass tracking press Cancel or enter a blank string"
  
  On Error GoTo InitImport_Err
  
  Call xSet("InitImport")
  
  'make sure the fixes have run
  Set Employer = PreImport(True)
  If (Employer Is Nothing) Then GoTo InitImport_End
  Set ben = Employer
  
  Set ben = Employer
  If Me.Tracking Then
    sTrackingMessage = S_TRACKING_MESSAGE
    Do While (Len(sTrackingDescription) = 0 Or Len(sTrackingDescription) > L_MAX_TRACKING_DESCRIPTION)
      sTrackingDescription = Trim$(InputBox(sTrackingMessage, "Tracking enabled"))
      If (Len(sTrackingDescription) = 0) Then
        Exit Do
      End If
      For i = 1 To Len(S_INVALID_FILE_CHARS)
        sTrackingDescription = Replace$(sTrackingDescription, Mid(S_INVALID_FILE_CHARS, i, 1), "")
      Next
      sTrackingDescription = Trim$(sTrackingDescription)
      i = Len(sTrackingDescription)
      If (i > L_MAX_TRACKING_DESCRIPTION) Then
        sTrackingMessage = "Your description was " & i & " character(s) long" & vbCrLf & vbCrLf & S_TRACKING_MESSAGE
      End If
    Loop
    If (Len(sTrackingDescription) > 0) Then
      'we are now tracking
      sGuid = ben.value(employer_GUID_db)
      dt = Now
      sDestFilename = sTrackingDescription & "_" & sGuid & "_"
      sDestFilename = sDestFilename & Year(dt) & Format$(Month(dt), "0#") & Format$(Day(dt), "0#") & Format$(Hour(dt), "0#") & Format$(Minute(dt), "0#") & Format$(Second(dt), "0#")
      sDestFilename = Me.TrackingPath & sDestFilename & S_FILE_EXTENSION_BAK
      sSourceFileName = p11d32.WorkingDirectory & ben.value(employer_FileName)
      Call Me.Zip(ben, sDestFilename)
      
    End If
  End If
  
  Set Employer = PreImport(True)
  If (Employer Is Nothing) Then GoTo InitImport_End
  Set ben = Employer
  Set p11d32.CurrentEmployer = Employer
  If Not Employer.db Is Nothing Then
    Set m_dbP11dFile = Employer.db
    Set wsError = DBEngine.CreateWorkspace("ErrorWS", p11d32.DBUserName, p11d32.DBPassword)
    Set m_dbError = InitDB(wsError, ben.value(employer_PathAndFile), ben.Name, , , False)
    If m_dbError Is Nothing Then Call Err.Raise(ERR_INITIMPORT, "InitImport", "Unable to open employer database." & vbCrLf & "Employer name: " & ben.Name)
    Set m_rsError = m_dbError.OpenRecordset(mSql.Queries(IMP_SELECT_ERRORS), dbOpenDynaset)
    If m_rsError Is Nothing Then Call Err.Raise(ERR_INITIMPORT, "InitImport", "Unable to open import errors table.")
    Do
      Set ic = New ImportClass
      Set ic.Workspace = p11d32.P11DWS
      Set ic.ErrorFilter = Me
      ic.RightTrimFields = True
      Set m_iw = ic.ImportWizard
      m_iw.AllowMultipleImports = True
      If Not ImportWizardSettings Then Exit Do
      m_iw.StartWizard
      If (m_iw.Imported) Then
        Call ben.ReadDB
        Set ifb = F_Employers
        Call ifb.UpdateBenefitListViewItem(F_Employers.LB.SelectedItem, ben)
      Else
        If (Len(sDestFilename) > 0) Then
          'Call MsgBox("Nothing imported, removing file stored by 'Import Tracking'", vbInformation)
          Call xKill(sDestFilename)
        End If
      End If
      Call ic.KillImporter
    Loop While m_iw.ImportAnother
  End If
  
InitImport_End:
  
  Call KillImport(ic)
  Call xReturn("InitImport")
  Exit Sub
InitImport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitImport", "Initialise import", "Error initialising the Importer.")
  Resume InitImport_End
  Resume
End Sub

Private Sub AddIRDescriptionImportData(qni As QUERY_NAMES_IMP, ParamArray Contraints())
  Dim i As Long
  
  'Category Description
  With m_iw
    Call .AddRequiredField(mSql.Name(qni), S_IMP_FIELD_CATEGORY_DESCRIPTION, False)
    Call .AddDefaultStaticLink(mSql.Name(qni), S_IMP_FIELD_CATEGORY_DESCRIPTION, TYPE_STR, S_IR_DESC_OTHER)
    Call .AddConstraint(mSql.Name(qni), S_IMP_FIELD_CATEGORY_DESCRIPTION, TYPE_STR, CONSTRAINT_INLIST, S_IR_DESC_OTHER)
    For i = LBound(Contraints) To UBound(Contraints)
      Call .AddConstraint(mSql.Name(qni), S_IMP_FIELD_CATEGORY_DESCRIPTION, TYPE_STR, CONSTRAINT_INLIST, CStr(Contraints(i)))
    Next
  End With
End Sub
Private Sub AddTaxYearConstraint(ImportName As String, sField As String)
  Call m_iw.AddConstraint(ImportName, sField, TYPE_DATE, CONSTRAINT_MIN, p11d32.Rates.value(TaxYearStart))
  Call m_iw.AddConstraint(ImportName, sField, TYPE_DATE, CONSTRAINT_MAX, p11d32.Rates.value(TaxYearEnd))
End Sub
Private Sub AddImportCompanyCars(ByVal bMatchOnAvailableFrom As Boolean, ByVal sNameSuffix As String, IPI As IProcessImport, dTaxYearStart As Date, dTaxYearEnd As Date)
  'COMPANY CARS
  Dim sName As String
  
  Set IPI = Me
  With m_iw
    sName = mSql.Name(IMP_COMPANYCARS) & sNameSuffix
    Call .AddSQL(mSql.Queries(IMP_COMPANYCARS), m_dbP11dFile, sName, IMP_COMPANYCARS, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(sName, "Personnel Number", True)
    Call .AddRequiredField(sName, "Registration number", True)
    Call .AddRequiredField(sName, "Make", True) 'so
    
    Call .AddDefaultStaticLink(sName, "Number of users", TYPE_LONG, 1)
    Call .AddConstraint(sName, "Number of users", TYPE_DATE, CONSTRAINT_MIN, 1)
    Call AddStandardDateFields(sName, dTaxYearStart, dTaxYearEnd)
    
    If bMatchOnAvailableFrom Then
      Call .AddRequiredField(sName, S_FIELD_AVAILABLE_FROM, True)
    End If
      
    Call AddTaxYearConstraint(sName, "Date car replaced")
    Call AddTaxYearConstraint(sName, S_FIELD_CAR_FUEL_WITHDRAWN_DATE)
    
    'need to adde constraints as strings as INLIST only accepts strings, should convery value on assigning to db as imp stores as variant
    
    Call .AddConstraint(sName, "Payment Frequency", TYPE_LONG, CONSTRAINT_INLIST, P46PF_ANNUALLY)   'so
    Call .AddConstraint(sName, "Payment Frequency", TYPE_LONG, CONSTRAINT_INLIST, P46PF_QUARTERLY)
    Call .AddConstraint(sName, "Payment Frequency", TYPE_LONG, CONSTRAINT_INLIST, P46PF_MONTHLY)
    Call .AddConstraint(sName, "Payment Frequency", TYPE_LONG, CONSTRAINT_INLIST, P46PF_WEEKLY)
    Call .AddConstraint(sName, "Payment Frequency", TYPE_LONG, CONSTRAINT_INLIST, P46PF_ACTUAL)
    Call .AddDefaultStaticLink(sName, "Payment Frequency", TYPE_LONG, P46PF_ACTUAL)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_PETROL)  'so
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_DIESEL)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_EUROIVDIESEL)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_BIFUEL_WITH_CO2_FOR_GAS)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_ELECTRIC)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_GAS_ONLY)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_HYBRID)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_E85_BIO_ENTHANOL_AND_PETROL)
    Call .AddConstraint(sName, "Type of fuel", TYPE_LONG, CONSTRAINT_INLIST, CCFT_RDE2_DIESEL)
    
    Call .AddDefaultStaticLink(sName, "Type of fuel", TYPE_LONG, CCFT_PETROL)

    Call .AddDefaultStaticLink(sName, "Amount is tax deducted", TYPE_BOOL, False)
  End With
End Sub
Private Sub AddImportQueries()
  Dim dTaxYearStart As Date, dTaxYearEnd As Date
  Dim IPI As IProcessImport
  On Error GoTo AddImportQueries_Err
  Call xSet("AddImportQueries")
  'format is
  '1 add query
  '2 Add keys
  '3 add defaults
  Set IPI = Me
  dTaxYearStart = p11d32.Rates.value(TaxYearStart)
  dTaxYearEnd = p11d32.Rates.value(TaxYearEnd)
  
  With m_iw
    
    'EMPLOYEES
    Call .AddSQL(mSql.Queries(IMP_EMPLOYEES), m_dbP11dFile, mSql.Name(IMP_EMPLOYEES), IMP_EMPLOYEES, , , IPI)
    
    Call .AddRequiredField(mSql.Name(IMP_EMPLOYEES), "Personnel Number", True)
    Call .AddDefaultValue(mSql.Name(IMP_EMPLOYEES), "Status/Position", TYPE_STR, S_STAFF)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Status/Position", TYPE_STR, CONSTRAINT_INLIST, S_STAFF)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Status/Position", TYPE_STR, CONSTRAINT_INLIST, S_DIRECTOR)
    Call .AddDefaultValue(mSql.Name(IMP_EMPLOYEES), "Date joined", TYPE_DATE, Null)
    Call .AddDefaultValue(mSql.Name(IMP_EMPLOYEES), "Date left", TYPE_DATE, Null)
    
            
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Date joined", TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Date left", TYPE_DATE, CONSTRAINT_MIN, dTaxYearStart)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Date left", TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
    Call .AddDefaultValue(mSql.Name(IMP_EMPLOYEES), "Not Subject to Class1a", TYPE_BOOL, False) 'FC - Class1A
    'cad p11d
    Call .AddFieldDisplayName(mSql.Name(IMP_EMPLOYEES), "Group1", p11d32.CurrentEmployer.GroupCode1Alias)
    Call .AddFieldDisplayName(mSql.Name(IMP_EMPLOYEES), "Group2", p11d32.CurrentEmployer.GroupCode2Alias)
    Call .AddFieldDisplayName(mSql.Name(IMP_EMPLOYEES), "Group3", p11d32.CurrentEmployer.GroupCode3Alias)
    
    Call .AddDefaultValue(mSql.Name(IMP_EMPLOYEES), "DOB", TYPE_DATE, Null)
    Call .AddDefaultValue(mSql.Name(IMP_EMPLOYEES), "Gender", TYPE_STR, S_GENDER_NA)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Gender", TYPE_STR, CONSTRAINT_INLIST, S_GENDER_MALE)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Gender", TYPE_STR, CONSTRAINT_INLIST, S_GENDER_FEMALE)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEES), "Gender", TYPE_STR, CONSTRAINT_INLIST, S_GENDER_NA)
    
    'Accommodation
    Call .AddSQL(mSql.Queries(IMP_ACCOMMODATION), m_dbP11dFile, mSql.Name(IMP_ACCOMMODATION), IMP_ACCOMMODATION, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_ACCOMMODATION), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_ACCOMMODATION), "Address", True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ACCOMMODATION), "Property owned by employer", TYPE_BOOL, True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ACCOMMODATION), "Is amount paid by employee rent", TYPE_BOOL, True)
    Call AddStandardDateFields(mSql.Name(IMP_ACCOMMODATION), dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ACCOMMODATION), "Date first occupied", TYPE_DATE, p11d32.Rates.value(LastTaxYearEnd))
    Call .AddConstraint(mSql.Name(IMP_ACCOMMODATION), "Date first occupied", TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ACCOMMODATION), "Amount is tax deducted", TYPE_BOOL, False)
    
    'EMPLOYEE ADDRESSES
    Call .AddSQL(mSql.Queries(IMP_ADDRESSES), m_dbP11dFile, mSql.Name(IMP_ADDRESSES), IMP_ADDRESSES, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_ADDRESSES), "Personnel Number", True)
       
    'ASSETS AT THE DISPOSAL OF THE EMPLOYEE
    Call .AddSQL(mSql.Queries(IMP_ASSETS_AT_DISPOSAL), m_dbP11dFile, mSql.Name(IMP_ASSETS_AT_DISPOSAL), IMP_ASSETS_AT_DISPOSAL, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_ASSETS_AT_DISPOSAL), "Personnel Number", True)
    
    Call .AddRequiredField(mSql.Name(IMP_ASSETS_AT_DISPOSAL), "Asset description", True)
    
    Call .AddStaticLink(mSql.Name(IMP_ASSETS_AT_DISPOSAL), "UDBCode", TYPE_STR, "L", , True)
    Call AddStandardDateFields(mSql.Name(IMP_ASSETS_AT_DISPOSAL), dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ASSETS_AT_DISPOSAL), "Amount is tax deducted", TYPE_BOOL, False)
      'Category Description
      Call AddIRDescriptionImportData(IMP_ASSETS_AT_DISPOSAL, S_IR_DESC_L_AIRCRAFT, S_IR_DESC_L_BOAT, S_IR_DESC_L_CORPORATEHOSP, S_IR_DESC_L_HOLIDAYACCOM, S_IR_DESC_L_TIMESHAREACCOM)
      
    'ASSETS TRANSFERRED
    Call .AddSQL(mSql.Queries(IMP_ASSETS_TRANSEFERRED), m_dbP11dFile, mSql.Name(IMP_ASSETS_TRANSEFERRED), IMP_ASSETS_TRANSEFERRED, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_ASSETS_TRANSEFERRED), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_ASSETS_TRANSEFERRED), "Asset description", True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ASSETS_TRANSEFERRED), "Amount is tax deducted", TYPE_BOOL, False)
      'Category Description
      Call AddIRDescriptionImportData(IMP_ASSETS_TRANSEFERRED, S_IR_DESC_A_CARS, S_IR_DESC_A_PRECIOUSMETALS, S_IR_DESC_A_PROPERTY)
    
    
    
    'BENEFICIAL_LOANS
    Call .AddSQL(mSql.Queries(IMP_BENEFICIAL_LOANS), m_dbP11dFile, mSql.Name(IMP_BENEFICIAL_LOANS), IMP_BENEFICIAL_LOANS, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_BENEFICIAL_LOANS), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_BENEFICIAL_LOANS), "Loan description", True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_BENEFICIAL_LOANS), "Amount is tax deducted", TYPE_BOOL, False)
    Call .AddDefaultStaticLink(mSql.Name(IMP_BENEFICIAL_LOANS), S_FIELD_LOAN_CHEAP_TAXABLE, TYPE_BOOL, True)
      

    'BENEFICIAL LOAN BALANCES
    Call .AddSQL(mSql.Queries(IMP_BENEFICIAL_LOAN_BALANCES), m_dbP11dFile, mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), IMP_BENEFICIAL_LOAN_BALANCES, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), "Loan description", True)
    Call .AddRequiredField(mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), "Date from", True)
    Call .AddStaticLink(mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), "Loan type", TYPE_STR, S_BENEFICIAL_LOANS_TYPE, , True)
    Call .AddConstraint(mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), "Date from", TYPE_DATE, CONSTRAINT_MIN, dTaxYearStart)
    Call .AddConstraint(mSql.Name(IMP_BENEFICIAL_LOAN_BALANCES), "Date from", TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
    
    'Chauffeur
    Call AddGeneralSQL(BC_CHAUFFEUR_OTHERO_N, IMP_CHAUFFEUR, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_CHAUFFEUR), "Amount is tax deducted", TYPE_BOOL, False)
    Call .AddStaticLink(mSql.Name(IMP_CHAUFFEUR), S_IMP_FIELD_CATEGORY_DESCRIPTION, TYPE_STR, S_IR_DESC_OTHER, , True)
    
    'COMPANY CARS
    Call AddImportCompanyCars(False, "", IPI, dTaxYearStart, dTaxYearEnd)
    'COMPANY CARS, match on date
    Call AddImportCompanyCars(True, " (match on '" + S_FIELD_AVAILABLE_FROM & "')", IPI, dTaxYearStart, dTaxYearEnd)
    
    'Credit Cards and Vouchers
    Call AddGeneralSQL(BC_VOUCHERS_AND_CREDITCARDS_C, IMP_CREDITCARDS_AND_VOUCHERS, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_CREDITCARDS_AND_VOUCHERS), "Amount is tax deducted", TYPE_BOOL, False)
    
    'EMPLOYEE CARS
'MP DB ToDo del/change line below? - IMP_COMPANY_CAR_MILES ref to deleted T_CarMiles table
    Call .AddSQL(mSql.Queries(IMP_EMPLOYEE_CARS), m_dbP11dFile, mSql.Name(IMP_EMPLOYEE_CARS), IMP_EMPLOYEE_CAR_MILEAGE, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_EMPLOYEE_CARS), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_EMPLOYEE_CARS), "Car", True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_EMPLOYEE_CARS), "Amount is tax deducted", TYPE_BOOL, False)
    
    'EMPLOYEE CAR MILEAGE 'uses same post process as company cars as same table
    Call .AddSQL(mSql.Queries(IMP_EMPLOYEE_CAR_MILEAGE), m_dbP11dFile, mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), IMP_EMPLOYEE_CAR_MILEAGE, , UPDATE_FIRST, IPI)
    'Call .AddLockUpdate(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), NO_UPDATES)
    Call .AddRequiredField(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Car", True)
    Call .AddRequiredField(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Description of mileage", True)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Date of car usage", TYPE_DATE, CONSTRAINT_MIN, dTaxYearStart)
    Call .AddConstraint(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Date of car usage", TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Description of mileage", TYPE_STR, "Mileage")
    Call .AddDefaultStaticLink(mSql.Name(IMP_EMPLOYEE_CAR_MILEAGE), "Date of car usage", TYPE_DATE, p11d32.Rates.value(TaxYearStart))
    
    'Entertainment
    Call AddGeneralSQL(BC_ENTERTAINMENT_N, IMP_ENTERTAINMENT, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_ENTERTAINMENT), "Amount is tax deducted", TYPE_BOOL, False)
    
    'General
    'Call AddGeneralSQL(BC_GENERAL_EXPENSES_BUSINESS_N, IMP_GENERAL, IPI, dTaxYearStart, dTaxYearEnd) ' removed 2016/17
    'Call .AddDefaultStaticLink(mSql.Name(IMP_GENERAL), "Amount is tax deducted", TYPE_BOOL, False) ' removed 2016/17
    
    
    'HOME PHONES
    Call .AddSQL(mSql.Queries(IMP_HOMEPHONES), m_dbP11dFile, mSql.Name(IMP_HOMEPHONES), IMP_HOMEPHONES, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_HOMEPHONES), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_HOMEPHONES), "Description", True)
    Call .AddStaticLink(mSql.Name(IMP_HOMEPHONES), "Class", TYPE_STR, p11d32.Rates.BenClassTo(BC_PHONE_HOME_N, BCT_DBCLASS), , True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_HOMEPHONES), "Amount is tax deducted", TYPE_BOOL, False)
    
    'Medical
    Call AddGeneralSQL(BC_PRIVATE_MEDICAL_I, IMP_MEDICAL, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_MEDICAL), "Amount is tax deducted", TYPE_BOOL, False)
    
    'NON shared vans
    Call .AddSQL(mSql.Queries(IMP_NON_SHARED_VANS), m_dbP11dFile, mSql.Name(IMP_NON_SHARED_VANS), IMP_NON_SHARED_VANS, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_NON_SHARED_VANS), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_NON_SHARED_VANS), "Description", True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_NON_SHARED_VANS), "Substantially available for Ordinary Commuting only?", TYPE_BOOL, False)
    Call .AddDefaultStaticLink(mSql.Name(IMP_NON_SHARED_VANS), "Mainly available for business travel only?", TYPE_BOOL, False)
    Call AddStandardDateFields(mSql.Name(IMP_NON_SHARED_VANS), dTaxYearStart, dTaxYearEnd)
'    If p11d32.AppYear > 2000 Then
    Call .AddDefaultStaticLink(mSql.Name(IMP_NON_SHARED_VANS), "Registration date", TYPE_DATE, p11d32.Rates.value(VanRegDateNew))
    
    'Section M
      'M class 1A
      Call AddGeneralSQL(BC_CLASS_1A_M, IMP_N_CLASS_1A, IPI, dTaxYearStart, dTaxYearEnd)
      Call .AddDefaultStaticLink(mSql.Name(IMP_N_CLASS_1A), "Amount is tax deducted", TYPE_BOOL, False)
      'Category Description
      Call AddIRDescriptionImportData(IMP_N_CLASS_1A, S_IR_DESC_M_C1A_ED_ASS, S_IR_DESC_M_C1A_NON_QUAL_RELOC, S_IR_DESC_M_C1A_STOP_LOSS_CHARGES, S_IR_DESC_M_C1A_SUBS_AND_FEES)
    
      'M Non class 1A
      Call AddGeneralSQL(BC_NON_CLASS_1A_M, IMP_N_NON_CLASS_1A, IPI, dTaxYearStart, dTaxYearEnd)
      Call .AddDefaultStaticLink(mSql.Name(IMP_N_NON_CLASS_1A), "Amount is tax deducted", TYPE_BOOL, False)
      'Category Description
      Call AddIRDescriptionImportData(IMP_N_NON_CLASS_1A, S_IR_DESC_M_NC1A_ED_ASS, S_IR_DESC_M_NC1A_LOANS_WRIT_WAIV, S_IR_DESC_M_NC1A_NURSERY, S_IR_DESC_M_NC1A_SUBS_AND_FEES)
      
    
    'N OTHER EXPENSES
    Call AddGeneralSQL(BC_OOTHER_N, IMP_O_OTHER_EXPENSES, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_O_OTHER_EXPENSES), "Amount is tax deducted", TYPE_BOOL, False)
    Call .AddDefaultStaticLink(mSql.Name(IMP_O_OTHER_EXPENSES), S_IMP_FIELD_CATEGORY_DESCRIPTION, TYPE_STR, S_IR_DESC_OTHER)
    
    
    'SECTION b
      'PAYMENTS MADE ON BEHALF OF EMPLOYEE
      Call AddGeneralSQL(BC_PAYMENTS_ON_BEFALF_B, IMP_PAYMENTS_ON_BEHALF_OF_EMPLOYEE, IPI, dTaxYearStart, dTaxYearEnd)
      Call .AddDefaultStaticLink(mSql.Name(IMP_PAYMENTS_ON_BEHALF_OF_EMPLOYEE), "Amount is tax deducted", TYPE_BOOL, False)
        'category description
        Call AddIRDescriptionImportData(IMP_PAYMENTS_ON_BEHALF_OF_EMPLOYEE, S_IR_DESC_B_ACCOUNTANCY_FEES, S_IR_DESC_B_DOMESTIC_BILLS, S_IR_DESC_B_PRIVATE_CAR_EX, S_IR_DESC_B_PRIVATE_CAR_EX, S_IR_DESC_B_PRIVATE_ED, S_IR_DESC_B_SEASON_TICKET)
        
      'Notional Payments
      Call AddGeneralSQL(BC_TAX_NOTIONAL_PAYMENTS_B, IMP_NOTIONAL_PAYMENTS, IPI, dTaxYearStart, dTaxYearEnd)
      Call .AddDefaultStaticLink(mSql.Name(IMP_NOTIONAL_PAYMENTS), "Amount is tax deducted", TYPE_BOOL, False)
    
    'RELOCATION ITEM
    Call .AddSQL(mSql.Queries(IMP_RELOCATION_ITEM), m_dbP11dFile, mSql.Name(IMP_RELOCATION_ITEM), IMP_RELOCATION_ITEM, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_RELOCATION_ITEM), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_RELOCATION_ITEM), "Address", True)
    Call .AddConstraint(mSql.Name(IMP_RELOCATION_ITEM), "Tax Year of move", TYPE_STR, CONSTRAINT_INLIST, p11d32.Rates.value(RelocationThisYear))
    Call .AddConstraint(mSql.Name(IMP_RELOCATION_ITEM), "Tax Year of move", TYPE_STR, CONSTRAINT_INLIST, p11d32.Rates.value(RelocationLastYear))
    Call .AddStaticLink(mSql.Name(IMP_RELOCATION_ITEM), "Tax Year of move", TYPE_STR, p11d32.Rates.value(RelocationThisYear))
    
    'RELOCATION DETAIL
    Call .AddSQL(mSql.Queries(IMP_RELOCATION_DETAIL), m_dbP11dFile, mSql.Name(IMP_RELOCATION_DETAIL), IMP_RELOCATION_DETAIL, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_RELOCATION_DETAIL), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_RELOCATION_DETAIL), "Item", True)
    Call .AddRequiredField(mSql.Name(IMP_RELOCATION_DETAIL), "Address", True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_RELOCATION_DETAIL), "Amount is tax deducted", TYPE_BOOL, False)
    
    'SERVICES PROVIDED
    Call .AddSQL(mSql.Queries(IMP_SERVICES_PROVIDED), m_dbP11dFile, mSql.Name(IMP_SERVICES_PROVIDED), IMP_SERVICES_PROVIDED, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_SERVICES_PROVIDED), "Personnel Number", True)
    Call .AddRequiredField(mSql.Name(IMP_SERVICES_PROVIDED), "Service description", True)
    Call .AddStaticLink(mSql.Name(IMP_SERVICES_PROVIDED), "UDBCode", TYPE_STR, "K", , True)
    Call .AddDefaultStaticLink(mSql.Name(IMP_SERVICES_PROVIDED), "Amount is tax deducted", TYPE_BOOL, False)
    
    'Shared Vans
    Call .AddSQL(mSql.Queries(IMP_SHARED_VANS), m_dbP11dFile, mSql.Name(IMP_SHARED_VANS), IMP_SHARED_VANS, , UPDATE_FIRST, IPI)
    Call .AddRequiredField(mSql.Name(IMP_SHARED_VANS), "Description", True)
    
    Call .AddDefaultStaticLink(mSql.Name(IMP_SHARED_VANS), "Amount is tax deducted", TYPE_BOOL, False)
    Call .AddDefaultStaticLink(mSql.Name(IMP_SHARED_VANS), "Personnel Number", TYPE_STR, S_SHAREDVAN_KEY)
    Call .HideField(mSql.Name(IMP_SHARED_VANS), "Personnel Number")
    Call AddStandardDateFields(mSql.Name(IMP_SHARED_VANS), dTaxYearStart, dTaxYearEnd)
    'Tax Paid
    Call AddGeneralSQL(BC_INCOME_TAX_PAID_NOT_DEDUCTED_M, IMP_TAXPAID, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_TAXPAID), "Amount is tax deducted", TYPE_BOOL, False)
    
    'Travel and Subsistence
    Call AddGeneralSQL(BC_TRAVEL_AND_SUBSISTENCE_N, IMP_TRAVEL_AND_SUBSISTENCE, IPI, dTaxYearStart, dTaxYearEnd)
    Call .AddDefaultStaticLink(mSql.Name(IMP_TRAVEL_AND_SUBSISTENCE), "Amount is tax deducted", TYPE_BOOL, False)
    
  End With


AddImportQueries_End:
  Set IPI = Nothing
  Call xReturn("AddImportQueries")
  Exit Sub

AddImportQueries_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddImportQueries", "Add Import Queries", "Error setting up the import queries for the import wizard")
  Resume AddImportQueries_End
  Resume
End Sub

Private Function AddGeneralSQL(ByVal bc As BEN_CLASS, ByVal QueryID As QUERY_NAMES_IMP, ByVal IPI As IProcessImport, ByVal dTaxYearStart As Date, ByVal dTaxYearEnd As Date) As Boolean
  On Error GoTo AddGeneralSQL_ERR
  
  Call xSet("AddGeneralSQL")
  
  Call m_iw.AddSQL(mSql.Queries(QueryID), m_dbP11dFile, mSql.Name(QueryID), QueryID, , UPDATE_FIRST, IPI)
  Call m_iw.AddRequiredField(mSql.Name(QueryID), "Personnel Number", True)
  Call m_iw.AddRequiredField(mSql.Name(QueryID), "Description", True)
  Call m_iw.AddRequiredField(mSql.Name(QueryID), "Class", True)
  Call m_iw.AddStaticLink(mSql.Name(QueryID), "Class", TYPE_STR, p11d32.Rates.BenClassTo(bc, BCT_DBCLASS), , True)
  Call m_iw.AddStaticLink(mSql.Name(QueryID), "UDBCode", TYPE_STR, p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING), , True)
  Call m_iw.AddStaticLink(mSql.Name(QueryID), "Category", TYPE_STR, p11d32.Rates.BenClassTo(bc, BCT_DBCATEGORY), , True)
  Call AddStandardDateFields(mSql.Name(QueryID), dTaxYearStart, dTaxYearEnd)
  
AddGeneralSQL_END:
  Call xReturn("AddGeneralSQL")
  Exit Function
AddGeneralSQL_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddGeneralSQL", "Add General SQL", "Error adding general sql, query id = " & QueryID & ".")
  Resume AddGeneralSQL_END
  Resume
End Function

Private Sub AddStandardDateFields(ByVal ImportName As String, ByVal dTaxYearStart As Date, ByVal dTaxYearEnd As Date)
  Call m_iw.AddDefaultStaticLink(ImportName, S_FIELD_AVAILABLE_FROM, TYPE_DATE, dTaxYearStart, S_FIELD_AVAILABLE_FROM)
  Call m_iw.AddDefaultStaticLink(ImportName, "Available to", TYPE_DATE, dTaxYearEnd, "Available to")
  Call m_iw.AddConstraint(ImportName, S_FIELD_AVAILABLE_FROM, TYPE_DATE, CONSTRAINT_MIN, dTaxYearStart)
  Call m_iw.AddConstraint(ImportName, S_FIELD_AVAILABLE_FROM, TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
  Call m_iw.AddConstraint(ImportName, "Available to", TYPE_DATE, CONSTRAINT_MAX, dTaxYearEnd)
  Call m_iw.AddConstraint(ImportName, "Available to", TYPE_DATE, CONSTRAINT_MIN, dTaxYearStart)
End Sub


Public Function GetStandardPostProcessSQL(sSelectSQL As String, ByVal sSourceTable As String, ByVal sSourceItem As String, ByVal sSourceItemName As String, Optional ByVal sSourceField As String = "P_NUM", Optional ByVal sParentTable As String = "T_EMPLOYEES", Optional ByVal sParentField As String = "P_NUM") As Long
  Dim sql As String
    
  Call xSet("GetStandardPostProcessSQL")
  sql = " FROM " & sSourceTable & " LEFT JOIN " & sParentTable & " ON " & sSourceTable & "." & sSourceField & " = " & sParentTable & "." & sParentField
  'are we testing on P_NUM in the  source table, if so then must not do any CDB_people
  If StrComp("P_NUM", sSourceField) Then
    sql = sql & " WHERE (" & sParentTable & "." & sParentField & " Is Null )"
  Else
    sql = sql & " WHERE (( " & sSourceTable & "." & sSourceField & " NOT Like " & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX_LIKE) & ") AND (" & sParentTable & "." & sParentField & " Is Null ))"
  End If
  sSelectSQL = "SELECT DISTINCTROW " & sSourceTable & "." & sSourceField & " AS " & S_IMP_PNUM_ALIAS & ", " & sSourceTable & "." & sSourceItem & " AS [" & sSourceItemName & "]" & sql
  
  Call xReturn("GetStandardPostProcessSQL")
End Function

Private Function IErrorFilter_FilterErrorMessage(ByVal Username As String, ByVal DateTime As Date, ByVal ErrorNumber As Long, ByVal ErrorName As String, ByVal ErrorText As String, ByVal SourceFunction As String) As Boolean
  Dim s() As String
  Dim lEqualsPos As Long
  On Error GoTo FilterError_err
   
  lEqualsPos = InStr(1, ErrorName, "=")
  If lEqualsPos > 1 Then
    ErrorName = Left$(ErrorName, lEqualsPos - 1)
  ElseIf lEqualsPos = 1 Then
    ErrorName = ""
  End If
   
  m_rsError.AddNew
  m_rsError.Fields("UserName") = Username
  m_rsError.Fields("DateTime") = DateTime
  m_rsError.Fields("ErrorNumber") = ErrorNumber
  m_rsError.Fields("ErrorDescription") = ErrorText
  
  Call GetDelimitedValues(s, ErrorName)
  m_rsError.Fields("ErrorItem") = s(1)
  m_rsError.Fields("Row") = s(2)
  m_rsError.Fields("Column") = s(3)
  m_rsError.Fields(S_IMP_PARENT) = s(4)
  m_rsError.Fields("Import") = m_iw.CurrentDest.Name
  m_rsError.Update
    
  
    
FilterError_end:
  Exit Function
FilterError_err:
  Resume Next

End Function

Private Function IErrorFilter_PrintAvailable() As Boolean
  IErrorFilter_PrintAvailable = True
End Function
Private Function IErrorFilter_PrintErrors() As Boolean
  Call ReportErrors(m_rsError, ReportHeader(Employer), PRINT_REPORT)
End Function
Private Function IErrorFilter_ViewErrors() As Boolean
  Call ReportErrors(m_rsError, ReportHeader(Employer), PREPARE_REPORT)
End Function
Private Function IErrorFilter_ViewAvailable() As Boolean
  IErrorFilter_ViewAvailable = True
End Function

Private Function IProcessImport_PostProcess(ic As ImportClass, RollBack As Boolean, LinesOK As Long, AddedRecordCount As Long, UpdatedRecordCount As Long) As Long
  Dim rs As Recordset
  Dim ErrorCount As Long, lTag As QUERY_NAMES_IMP, sImport As String
  Dim sSelectSQL As String
  Dim bPostProcessed As Boolean
  
  On Error GoTo PostProcessImport_ERR
  Call xSet("PostProcessImport")
  If Not ic Is Nothing Then
    If Not ic.ImportWizard Is Nothing Then
      If Not ic.ImportWizard.CurrentDest Is Nothing Then
        sImport = ic.ImportWizard.CurrentDest.Name
        lTag = ic.ImportWizard.CurrentDest.Tag
        Call GetPostProcessDataSQL(sSelectSQL, lTag, sImport)
        bPostProcessed = Len(sSelectSQL) > 0
        If bPostProcessed Then
          ErrorCount = 0
          Set rs = m_dbP11dFile.OpenRecordset(sSelectSQL, dbOpenSnapshot)
          If Not (rs.EOF And rs.BOF) Then
            rs.MoveLast
            ErrorCount = rs.RecordCount
          End If
          If ErrorCount > 0 Then
            Call MultiDialog("Errors in import", "Importing " & sImport & vbCrLf & CStr(ErrorCount) & " error(s) encountered in post processing", "&Ok")
            rs.MoveFirst
            Do
              Err.Number = ERR_POST_PROCESS
              Err.Description = "The " & rs.Fields(1).Name & " did not have a corresponding parent " & rs.Fields(0).Name
              Call ErrorMessage(ERR_ERROR, Err, "PostProcess", rs.Fields(1).value & ",,," & rs.Fields(0).value, Err.Description)
             
              rs.MoveNext
            Loop Until rs.EOF
            LinesOK = LinesOK - ErrorCount
            
          End If
        End If
        If ErrorCount = 0 Then ErrorCount = BackwardsCompatibilityPostProcessing(lTag)
      End If
      
    End If
  End If
      
PostProcessImport_END:
  IProcessImport_PostProcess = ErrorCount
  If ErrorCount > 0 Then RollBack = True
  Set rs = Nothing
  Call xReturn("PostProcessImport")
  Exit Function
PostProcessImport_ERR:
  ErrorCount = ErrorCount + 1
  Call ErrorMessage(ERR_ERROR, Err, "PostProcessImport", "Post Process Import", "Error in post processing the import for '" & sImport & "'.")
  Resume PostProcessImport_END
  Resume
End Function
Private Function QuestionRollBack(ic As ImportClass, ByVal bPostProcessed As Boolean, ByVal lImportTag As Long)
  On Error GoTo QuestionRollBack_ERR
  
  Call xSet("QuestionRollBack")
  
  
  If ic Is Nothing Then ECASE ("Import class is nothing")
  If (lImportTag = IMP_EMPLOYEES) And Not bPostProcessed Then
    If ic.ProcessingLinesInError > 0 Then
      QuestionRollBack = Not (MsgBox("There were " & ic.ProcessingLinesInError & " line(s) in error." & vbCrLf & vbCrLf & "Do you wish to import the other records regardless?", vbOKCancel, "QuestionRollBack") = vbOK)
    End If
  End If
  
QuestionRollBack_END:
  Call xReturn("QuestionRollBack")
  Exit Function
QuestionRollBack_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "QuestionRollBack", "Question Roll Back", "Error question the roll back of an import.")
  Resume QuestionRollBack_END
End Function

Private Function IProcessImport_PreProcess(ic As ImportClass) As Long
  On Error GoTo PreProcessImport_ERR
  Call xSet("PreProcessImport")
  
  If Not m_dbError Is Nothing Then
    Call m_dbError.Execute(mSql.Queries(IMP_DELETE_ERRORS))
    m_rsError.Requery
  End If

PreProcessImport_END:
  Call xReturn("PreProcessImport")
  Exit Function
  
PreProcessImport_ERR:
  Resume PreProcessImport_END
End Function

Private Function ImportWizardSettings() As Boolean
  On Error GoTo ImportWizardSettings_Err
  Call xSet("ImportWizardSettings")

  With m_iw
    .PreviewLines = 6
    Call AddImportQueries
    Call .AddFileExtension("Comma delimeted files (*.csv)", "*.csv", True, True)
    Call .AddFileExtension("Text files (*.txt)", "*.txt")
    Call .AddFileExtension("All files (*.*)", "*.*")
  
    .AutoOpenSourceFile = True
    .SourcePath = p11d32.WorkingDirectory
    .SpecPath = p11d32.WorkingDirectory
    .TitlePrefix = GetStatic("ApplicationName") & " Importer  -  "
    .AllowChangeUpdateType = True
  End With
  ImportWizardSettings = True
  
ImportWizardSettings_End:
  Call xReturn("ImportWizardSettings")
  Exit Function

ImportWizardSettings_Err:
  ImportWizardSettings = False
  'cd apf
  Call ErrorMessage(ERR_ERROR, Err, "ImportWizardSettings", "Add Import Wizard settings", "Error adding default setting to the import wizard.")
  Resume ImportWizardSettings_End
  Resume
End Function
Public Property Get ReportHeader(emp As Employer) As String
  Dim ben As IBenefitClass
  
  On Error GoTo ReportHeader_ERR
  
  Call xSet("ReportHeader")
  
  
  Set ben = Employer
  ReportHeader = "Import errors for: " & GetCanonicalPathName(ben.value(employer_PathAndFile)) & vbCrLf & vbCrLf
  
ReportHeader_END:
  Call xReturn("ReportHeader")
  Exit Property
ReportHeader_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportHeader", "Report Header", "Error getting the report header for import errors report.")
  Resume ReportHeader_ERR
  Resume
End Property
Private Sub UpdateGUIDKeys(ByVal QueryID As QUERY_NAMES_IMP)
  Dim rs As Recordset
  Dim sSelectSQL As String
  Dim sFieldName As String
  
  Select Case QueryID
    Case IMP_BENEFICIAL_LOANS
      sSelectSQL = "SELECT T_BenLoan.BalKey FROM T_BenLoan"
      sFieldName = "BalKey"
    Case IMP_EMPLOYEE_CAR_MILEAGE
      sSelectSQL = "SELECT T_BenEeCar.EeCarKey FROM T_BenEeCar"
      sFieldName = "EECarKey"
    Case IMP_RELOCATION_DETAIL
      sSelectSQL = "SELECT T_BenReloc.RelocKey FROM T_BenReloc"
      sFieldName = "RelocKey"
  End Select
  
  Set rs = m_dbP11dFile.OpenRecordset(sSelectSQL, dbOpenDynaset)
  Do While Not rs.EOF
    If IsNullEx(rs.Fields(sFieldName).value, "") = "" Then
      rs.Edit
      rs.Fields(sFieldName) = GenerateGUID
      rs.Update
    End If
    rs.MoveNext
  Loop
End Sub


Private Sub GetPostProcessDataSQL(sSelectSQL As String, ByVal ImportSpecTag As Long, ByVal ImportName As String)
  On Error GoTo GetPostProcessDataSQL_Err
  Call xSet("GetPostProcessDataSQL")
  sSelectSQL = ""
  Select Case ImportSpecTag
    Case IMP_EMPLOYEES
    Case IMP_HOMEPHONES
        Call GetStandardPostProcessSQL(sSelectSQL, "T_BENPHONE", "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_CHAUFFEUR, IMP_CREDITCARDS_AND_VOUCHERS, IMP_EDUCATION, IMP_ENTERTAINMENT, IMP_GENERAL, IMP_MEDICAL, IMP_NOTIONAL_PAYMENTS, IMP_N_NON_CLASS_1A, IMP_N_CLASS_1A, IMP_TAXPAID, IMP_TRAVEL_AND_SUBSISTENCE
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_VOUCHERS_AND_CREDITCARDS_C, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_O_OTHER_EXPENSES
        'NEEDS NOT TO QUERY CDB_ EMPLOYEE BENEFITS
        sSelectSQL = mSql.Queries(POST_PROCESS_OOTHER_O)
    Case IMP_ACCOMMODATION
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_LIVING_ACCOMMODATION_D, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_COMPANYCARS
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_COMPANY_CARS_F, BCT_TABLE), S_FIELD_CAR_REGISTRATION, Left$(ImportName, Len(ImportName) - 1))
    Case IMP_EMPLOYEE_CARS
      Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_EMPLOYEE_CAR_E, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
      'CHANGE FUEL TYPE FROM D TO L if reg after 01/01/2006
      
    Case IMP_BENEFICIAL_LOANS
        Call UpdateGUIDKeys(IMP_BENEFICIAL_LOANS)
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_LOAN_OTHER_H, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_ADDRESSES
        Call GetStandardPostProcessSQL(sSelectSQL, "T_ADDRESSES", "Address1", Left$(ImportName, Len(ImportName) - 2))
    Case IMP_ASSETS_TRANSEFERRED
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_ASSETSTRANSFERRED_A, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_ASSETS_AT_DISPOSAL
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_ASSETSATDISPOSAL_L, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_SERVICES_PROVIDED
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_SERVICES_PROVIDED_K, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_RELOCATION_ITEM
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_QUALIFYING_RELOCATION_J, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_NON_SHARED_VANS
        sSelectSQL = mSql.Queries(POST_PROCESS_NON_SHARED_VANS_SELECT)
    Case IMP_SHARED_VANS
        sSelectSQL = mSql.Queries(POST_PROCESS_SHARED_VANS_SELECT)
    Case IMP_PAYMENTS_ON_BEHALF_OF_EMPLOYEE
        Call GetStandardPostProcessSQL(sSelectSQL, p11d32.Rates.BenClassTo(BC_PAYMENTS_ON_BEFALF_B, BCT_TABLE), "Item", Left$(ImportName, Len(ImportName) - 1))
    Case IMP_EMPLOYEE_CAR_MILEAGE
        Call UpdateGUIDKeys(IMP_EMPLOYEE_CAR_MILEAGE)
        sSelectSQL = mSql.Queries(POST_PROCESS_EMPLOYEECAR_MILEAGE_SELECT)
    Case IMP_BENEFICIAL_LOAN_BALANCES
        sSelectSQL = mSql.Queries(POST_PROCESS_BENEFICIAL_LOAN_BALANCES_SELECT)
    Case IMP_RELOCATION_DETAIL
        Call UpdateGUIDKeys(IMP_RELOCATION_DETAIL)
        sSelectSQL = mSql.Queries(POST_PROCESS_RELOCATION_DETAIL_SELECT)
    Case Else
      Call ECASE("GetPostProcessDataSQL")
      GoTo GetPostProcessDataSQL_End
  End Select

GetPostProcessDataSQL_End:
  Call xReturn("GetPostProcessDataSQL")
  Exit Sub

GetPostProcessDataSQL_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetPostProcessDataSQL", "Get postprocess sql string", "Unable to retrieve post process sql string. Import number: " & CStr(ImportSpecTag) & ".")
  Resume GetPostProcessDataSQL_End
  Resume
End Sub


Private Function BackwardsCompatibilityPostProcessing(ByVal ImportTag As QUERY_NAMES_IMP) As Long

  On Error GoTo BackwardsCompatibilityPostProcessing_Err
  Call xSet("BackwardsCompatibilityPostProcessing")
  If Not m_dbP11dFile Is Nothing Then
    Select Case ImportTag
      Case IMP_COMPANYCARS
        m_dbP11dFile.Execute (mSql.Queries(UPDATE_COMPANY_CAR_FUEL_AVAILABLE_TO))
      Case IMP_COMPANY_CAR_MILES
        'm_dbP11dFile.Execute (mSql.Queries(UPDATE_COMPANY_CAR_KEYS))
      Case IMP_EMPLOYEE_CAR_MILEAGE
        m_dbP11dFile.Execute (mSql.Queries(UPDATE_EMPLOYEE_CAR_KEYS))
      Case IMP_BENEFICIAL_LOAN_BALANCES
        m_dbP11dFile.Execute (mSql.Queries(UPDATE_BENEFICIAL_LOAN_BALANCES_KEYS))
      Case IMP_RELOCATION_DETAIL
        m_dbP11dFile.Execute (mSql.Queries(UPDATE_RELOCATION_KEYS))
    End Select
  End If
  
BackwardsCompatibilityPostProcessing_End:
  Call xReturn("BackwardsCompatibilityPostProcessing")
  Exit Function

BackwardsCompatibilityPostProcessing_Err:
  BackwardsCompatibilityPostProcessing = 1
  Call ErrorMessage(ERR_ERROR, Err, "BackwardsCompatibilityPostProcessing", "Backwards Compatibility Post Processing", "Error processing imports for backwards compatability.")
  Resume BackwardsCompatibilityPostProcessing_End
  Resume
End Function
Private Function LogPostProcessingErrors(rsRecordsWhereParentNotAvailable As Recordset, sPostProcessingErrorMessage As String, sSrcItemField As String) As Boolean

  On Error GoTo LogPostProcessingErrors_Err
  Call xSet("LogPostProcessingErrors")

  With rsRecordsWhereParentNotAvailable
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        Call Err.Raise(ERR_POST_PROCESS, "LogPostProcessingErrors", sPostProcessingErrorMessage)
        .MoveNext
      Loop
    End If
  End With

LogPostProcessingErrors_End:
  Call xReturn("LogPostProcessingErrors")
  Exit Function

LogPostProcessingErrors_Err:
  If Not (rsRecordsWhereParentNotAvailable.EOF And rsRecordsWhereParentNotAvailable.BOF) Then
    Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "LogPostProcessingErrors", rsRecordsWhereParentNotAvailable.Fields(sSrcItemField), Err.Description)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "LogPostProcessingErrors", "Log Post Processing Errors", "Error loging a post processing error.")
  End If
  Resume Next
End Function
