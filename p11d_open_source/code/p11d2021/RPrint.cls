VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEnumFiles


Public Enum RPT_TYPE
  RPTT_STANDARD = 1
  RPTT_MANAGEMENT
  RPTT_ABACUSUDM
  RPTT_OTHER
  RPTT_MM
  RPTT_USER
  RPTT_UNKNOWN
End Enum


Public Enum P11D_REPORTS
  'this enum determines the order they are placed to the screen
  [RPT_START] = 1
  [RPT_FIRST_INI] = [RPT_START]
  RPT_HMIT = [RPT_START]
  [RPT_FIRST_STANDARD] = [RPT_START]
  RPT_HMIT_PLUS_WORKING_PAPERS
  RPT_EMPLOYEE_LETTER
  RPT_EMPLOYEE_LETTER_EMAIL
  RPT_EMPLOYEE_LETTER_WORKING_PAPERS
  ' add new standard reports here
  RPT_WORKINGPAPERS
  RPT_P46CAR
  
  RPT_PRINTEDEMPLOYEES  'jn code
  [RPT_LAST_STANDARD] = RPT_PRINTEDEMPLOYEES
   RPT_P11Db
  [RPT_LAST_INI] = RPT_P11Db
  [RPT_FIRST_OTHER_REPORT] = RPT_P11Db
  RPT_MM_SUBMISSION
  [RPT_LAST_OTHER_REPORT] = RPT_MM_SUBMISSION
  
  RPT_MANAGEMENT
    
  RPT_MANAGEMENT_EMPLOYEE_LIST
  [RPT_FIRST_MANAGEMENT] = RPT_MANAGEMENT_EMPLOYEE_LIST
  RPT_MANAGEMENT_BENEFIT_SUMMARY
  RPT_MANAGEMENT_NIC_CLASS1A_EMPLOYER
  RPT_MANAGEMENT_NIC_CLASS1A_EMPLOYEE
    
  RPT_MANAGEMENT_ASSETS_TRANSFERRED_A
  RPT_MANAGEMENT_PAYMENTS_ON_BEHALF_B
  RPT_MANAGEMENT_TAX_ON_NOTIONAL_PAYMENTS_B
  RPT_MANAGEMENT_VOUCHERS_AND_CREDIT_CARDS_C
  RPT_MANAGEMENT_LIVING_ACCOMODATION_D
  RPT_MANAGEMENT_EMPLOYEE_CARS_E
  'company cars
  RPT_MANAGEMENT_COMPANY_CAR
  RPT_MANAGEMENT_COMPANY_CAR_SUMMARY_F
  RPT_MANAGEMENT_COMPANY_CAR_DETAILS_F 'EK NEW CO CAR REPORT
  RPT_MANAGEMENT_COMPANY_CAR_AND_FUEL_F
  RPT_MANAGEMENT_COMPANY_CAR_CLASS1A_F
  
  RPT_MANAGEMENT_COMPANY_CAR_FUEL_REP_F 'EK new co car report, 1/04 TTP#113
  RPT_MANAGEMENT_COMPANY_CAR_FUEL_VAT_F
  RPT_MANAGEMENT_COMPANY_CAR_MISSING_MAKE_F
'AM To be removed  RPT_MANAGEMENT_COMPANY_CAR_MILES
  RPT_MANAGEMENT_COMPANY_CAR_NO_EMMISIONS_F             'RH
  RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_1_SUMMARY_F
  RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_2_SUMMARY_F
  RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_3_SUMMARY_F
  RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_4_SUMMARY_F
  'end cars
  RPT_MANAGEMENT_VANS_G
'AM To be removed RPT_MANAGEMENT_LOANS_HOME_H
  RPT_MANAGEMENT_LOANS_OTHER_H
  RPT_MANAGEMENT_PRIVATE_MEDICAL_I
  RPT_MANAGEMENT_QUALIFYING_RELOCATION_J
  RPT_MANAGEMENT_SERVICES_PROVIDED_K
  RPT_MANAGEMENT_ASSETS_AT_DISPOSAL_L
  RPT_MANAGEMENT_SUBSCRIPTION_AND_PROFESSIONAL_M
'AM To be removed  RPT_MANAGEMENT_EDUCATION_OTHER_N
  RPT_MANAGEMENT_NURSERY_OTHER_M
  RPT_MANAGEMENT_INCOME_TAX_PAID_NOT_DEDUCTED_M
  RPT_MANAGEMENT_O
  RPT_MANAGEMENT_TRAVEL_AND_SUBSISTENCE_N
  RPT_MANAGEMENT_ENTERTAINMENT_N
  'RPT_MANAGEMENT_GENERAL_EXPENSES_BUSINESS_N
  RPT_MANAGEMENT_PHONE_HOME_N
  RPT_MANAGEMENT_NON_QUALIFYING_RELOCATION_N
  RPT_MANAGEMENT_CHAUFFEUR_N
  RPT_MANAGEMENT_OOTHER_N
  
  'add new management reports here
  [RPT_LAST_MANAGEMENT] = RPT_MANAGEMENT_OOTHER_N
  
  'Abacus Export
  RPT_ABACUSUDM
  
  RPT_ABACUSUDM_ASSETS_TRANSFERRED_A
  [RPT_FIRST_ABACUSUDM] = RPT_ABACUSUDM_ASSETS_TRANSFERRED_A
  RPT_ABACUSUDM_PAYMENTS_ON_BEHALF_B
  RPT_ABACUSUDM_TAX_ON_NOTIONAL_PAYMENTS_B
  RPT_ABACUSUDM_VOUCHERS_AND_CREDIT_CARDS_C
  RPT_ABACUSUDM_LIVING_ACCOMODATION_D
  RPT_ABACUSUDM_EMPLOYEE_CARS_E
  RPT_ABACUSUDM_COMPANY_CAR_F
  RPT_ABACUSUDM_FUEL_F
  RPT_ABACUSUDM_VANS_G
'AM To be removed  RPT_ABACUSUDM_LOANS_HOME_H
  RPT_ABACUSUDM_LOANS_OTHER_H
  RPT_ABACUSUDM_PRIVATE_MEDICAL_I
  RPT_ABACUSUDM_QUALIFYING_RELOCATION_J
  RPT_ABACUSUDM_SERVICES_PROVIDED_K
  RPT_ABACUSUDM_ASSETS_AT_DISPOSAL_L
  
  RPT_ABACUSUDM_NURSERY_OTHER_M
  RPT_ABACUSUDM_SUBSCRIPTIONS_AND_PROFESSIONAL_M
  RPT_ABACUSUDM_INCOME_TAX_PAID_NOT_DEDUCTED_M
  RPT_ABACUSUDM_TRAVEL_AND_SUBSISTENCE_N
  RPT_ABACUSUDM_ENTERTAINMENT_N
  'RPT_ABACUSUDM_GENERAL_EXPENSES_BUSINESS_N
  RPT_ABACUSUDM_PHONE_HOME_N
  RPT_ABACUSUDM_NON_QUALIFYING_RELOCATION_N
  'RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_O          'rh
  'RPT_ABACUSUDM_NOTHER_O                      'rh
  RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_N_OTHER
  RPT_ABACUSUDM_OOTHER_N
  [RPT_LAST_ABACUSUDM] = RPT_ABACUSUDM_OOTHER_N
  
  RPT_USER
  [RPT_END] = RPT_USER
End Enum

Public Enum REPORT_STANDARD_DATA
  RSD_LETTERFILE
  RSD_OUTPUTREPORT
End Enum

Public Enum REPORT_DESTINATION
  REPD_PRINTER_FIRST_ITEM = 0
  REPD_PRINTER_PREVIEW = REPD_PRINTER_FIRST_ITEM
  REPD_FILE_HTML
  REPD_FILE_PRN
  
  REPD_PRINTER_LAST_ITEM = REPD_FILE_PRN
End Enum

Public Enum REPORT_EXPORT_EMPLOYEE_FILE_NAME_FIELDS
  REEFNF_FIRST_ITEM = 0
  REEFNF_PERSONNEL_NUMBER = REEFNF_FIRST_ITEM
  REEFNF_SURNAME_FIRTNAME_PERSONNEL_NUMBER = 1
  REEFNF_LAST_ITEM = REEFNF_SURNAME_FIRTNAME_PERSONNEL_NUMBER
End Enum


Public Enum HMIT_SELECTION_CHOICE
  HMIT_SC_ALL
  HMIT_SC_RELEVANT
  HMIT_SC_SELECTED
End Enum

Public Enum P11D_TOTAL_VALUE
  PTV_NONZERO_AND_ZERO
  PTV_NONZERO
  PTV_ZERO
End Enum

Public Enum P46_RANGE
  P46_CURRENTQUARTER = -1
  P46_QUARTER1 = 0
  P46_QUARTER2
  P46_QUARTER3
  P46_QUARTER4
  P46_USERRANGE
End Enum

Public Enum SYSTEM_LETTER_FILE
  SLF_HMIT
  SLF_WORKING_PAPERS
  SLF_HMIT_PLUS_WORKING_PAPERS
End Enum

Public Enum HMIT_SECTIONS
  HMIT_A = 1
  [HMIT_FIRST_ITEM] = HMIT_A
  HMIT_B
  HMIT_C
  HMIT_D
  HMIT_E
  HMIT_F
  HMIT_G
  HMIT_H
  HMIT_I
  HMIT_J
  HMIT_K
  HMIT_L
  HMIT_M
  HMIT_N
  [HMIT_LAST_ITEM] = HMIT_N
  HMIT_NOT
  [_HMIT_COUNT] = HMIT_NOT
End Enum

Private m_PrinterAvailable As Boolean
Private m_ReportPathUser As String
Private m_UsersLettersPath As String

Public SelectedEmployees As ObjectList

'F_Print data
Public P46Range As P46_RANGE
Public P46PrintReplacedP46s As Boolean
Public EmployeeSelection As EMPLOYEE_SELECTION
Public HMITSelectionChoice As HMIT_SELECTION_CHOICE
Public P11DTotalValue As P11D_TOTAL_VALUE
'F_PrintOptions data
Public ExportNoSpaces As Boolean
Public UseNotesForEmail As Boolean
Public SortOtherTypeBenefitsAlphabetically  As Boolean
Public NotesIniFile As String
Public P46DateFrom As Date
Public P46DateTo As Date
Public P46UserDateFrom As Variant
Public P46UserDateTo As Variant
Public HMITSections As Long

Public HMITSections_PRINT As Long
Public PrintedEmployees As Boolean 'JN CODE
Public TimeStamp As Boolean
Public EmployeeLetterMargin As Long
Public Destination As REPORT_DESTINATION
Public PRNFileName As String
Public PRNFileExtension As String
Public EmployeeLetterFontName As String
Private m_EmployeeLetterFontSize As Long
Public EnableEmailReports As Boolean
Public EmailDebug As Boolean
Public AllReports As Boolean
Public AbacusUDM As Boolean
Public UserReportFileLessExtension As String
Public UserReportSelectEmployeeFileLessExtension As String

Public CheckOptions As CHECK_BEFORE_PRINT
Public CHECKORDERBY As CHECKORDERBY
Public ChecksAutoRefresh As Boolean
Public A4ForcePrint As Boolean
Public ChecksSelected As Long

Private m_ExportOption As REPORT_EXPORTS

Public ExportAllowUserNameHTML As Boolean
Public ExportEmployeeFileNameFields As REPORT_EXPORT_EMPLOYEE_FILE_NAME_FIELDS
Public EmailSubject As String
Public SeparatePrintJobs As Boolean

Public EmailTextAsHTML As Boolean
Public DatesOnWorkingPaper As Boolean
Public HMITFieldTrim As Boolean
Public CancelPrint As Boolean
Public PRNLoopLength As Long
Public EmployeeLetterFile As String
Public RemeberEmployeeSelection As Boolean
Private m_ATCMail As Mail
Private m_ExportDirectory As String
Private m_DefaultSelectEmployeeReportIndex As Long
Private m_DefaultReportIndex As Long
Private m_RepSettings(RPT_START To RPT_END) As REPORT_SETTINGS
Private m_EmployeeLetterPath As String
Public Property Get ReportSettingsIgnoreZeroOnly(ByVal rpt As P11D_REPORTS) As Boolean
  ReportSettingsIgnoreZeroOnly = m_RepSettings(rpt).IgnoreZeroOnly
End Property

Public Property Get EmployeeLetterPath() As String
  EmployeeLetterPath = m_EmployeeLetterPath
End Property
Public Property Let EmployeeLetterPath(ByVal NewValue As String)
  m_EmployeeLetterPath = NewValue
End Property
' apf can check if I have tried to access this before
Public Property Get ATCMail() As Mail
  Set ATCMail = m_ATCMail
End Property
Public Property Get DefaultReportIndex() As Long
  DefaultReportIndex = m_DefaultReportIndex
End Property
Public Property Get DefaultSelectEmployeeReportIndex() As Long
  DefaultSelectEmployeeReportIndex = m_DefaultSelectEmployeeReportIndex
End Property
Public Property Let DefaultSelectEmployeeReportIndex(ByVal NewValue As Long)
  m_DefaultSelectEmployeeReportIndex = LetDefaultReportIndex(NewValue)
End Property
Public Property Let DefaultReportIndex(ByVal NewValue As Long)
  m_DefaultReportIndex = LetDefaultReportIndex(NewValue)
End Property
Private Function LetDefaultReportIndex(ByVal NewValue As Long) As Long
  If NewValue = 0 Then NewValue = RPT_FIRST_INI 'cos some prat changed the enum values
  LetDefaultReportIndex = NewValue
End Function

Public Function ReportPrintPrapare(ByVal ReportIndex As Long) As RPT_TYPE
  Dim rt As RPT_TYPE
  
  rt = p11d32.ReportPrint.ReportType(ReportIndex)
  If rt < 0 Then Call Err.Raise(ERR_REPORT_INVALID, ErrorSource(Err, "ReportPrintPrapare"), Err.Description)
    
  If rt = RPTT_MANAGEMENT Or rt = RPTT_STANDARD Or rt = RPTT_USER Or rt = RPTT_ABACUSUDM Then
    If p11d32.ReportPrint.NoOfEmployeesSelected(p11d32.CurrentEmployer) = 0 Then Call Err.Raise(ERR_NO_EMPLOYEES_SELECTED, "DoPrint", "No employees selected for printing (see print dialogue 'Print Range').")
  End If
  
  ReportPrintPrapare = rt
  
End Function
Private Function ManagementReportCaption(ByVal bc As BEN_CLASS, Optional sCaption As String = "", Optional bCaptionOnly As Boolean = False)
  If Len(sCaption) > 0 Then
    If bCaptionOnly Then
      ManagementReportCaption = sCaption
    Else
      ManagementReportCaption = p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING) & ": " & sCaption
    End If
  Else
    ManagementReportCaption = p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING) & ": " & p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION)
  End If
End Function

Public Sub Initialise()
  Dim dev As String, port As String
  Dim i As Long
  Dim sParentName As String, sParentName2 As String
  Dim sParentName3 As String
  
  On Error GoTo Initialise_ReportPrint_ERR
  Call xSet("Initialise_ReportPrint")
  
  'standard reports
  m_PrinterAvailable = GetDefaultPrinter(dev, port)
  m_RepSettings(RPT_P46CAR).Name = "P46(Car)"
  
'  If p11d32.AppYear > 2000 Then
  m_RepSettings(RPT_P46CAR).IgnoreZeroOnly = True 'km
  
  m_RepSettings(RPT_HMIT).Name = "P11D return"
  m_RepSettings(RPT_EMPLOYEE_LETTER).Name = "Employee letter"
  m_RepSettings(RPT_EMPLOYEE_LETTER_EMAIL).Name = "Employee letter - 'e-mail'"
  m_RepSettings(RPT_HMIT_PLUS_WORKING_PAPERS).Name = "P11D return + working papers"
  m_RepSettings(RPT_WORKINGPAPERS).Name = "Working papers"
  m_RepSettings(RPT_P11Db).Name = "P11D(b)"
  m_RepSettings(RPT_MM_SUBMISSION).Name = "Expenses and Benefits Return Submission Document"
  m_RepSettings(RPT_PRINTEDEMPLOYEES).Name = "Printed Employees Report"
  'm_RepSettings(RPT_INFO_SHEET_PLUS_HMIT).Name = "Info sheet +  P11D return"
  'm_RepSettings(RPT_INFO_SHEET_PLUS_HMIT_PLUS_WORKING_PAPERS).Name = "Info sheet + P11D return + Working papers"
  m_RepSettings(RPT_EMPLOYEE_LETTER_WORKING_PAPERS).Name = "Employee letter + Working papers"
  
  'management reports
  'header
  m_RepSettings(RPT_MANAGEMENT).Name = "Management reports"
  m_RepSettings(RPT_MANAGEMENT).GroupHeader = True
  'reports
  sParentName = m_RepSettings(RPT_MANAGEMENT).Name
  
  Call AddParentedReport(RPT_MANAGEMENT_BENEFIT_SUMMARY, sParentName, 0, "Benefit summary", True)
  Call AddParentedReport(RPT_MANAGEMENT_EMPLOYEE_LIST, sParentName, 0, "Employee list", True)
  Call AddParentedReport(RPT_MANAGEMENT_NIC_CLASS1A_EMPLOYER, sParentName, 0, "Class 1A Summary for employer", True)
  Call AddParentedReport(RPT_MANAGEMENT_NIC_CLASS1A_EMPLOYEE, sParentName, 0, "Class 1A Summary for employees", True)
  
  Call AddParentedReport(RPT_MANAGEMENT_ASSETS_TRANSFERRED_A, sParentName, BC_ASSETSTRANSFERRED_A)
  Call AddParentedReport(RPT_MANAGEMENT_PAYMENTS_ON_BEHALF_B, sParentName, BC_PAYMENTS_ON_BEFALF_B)
  Call AddParentedReport(RPT_MANAGEMENT_TAX_ON_NOTIONAL_PAYMENTS_B, sParentName, BC_TAX_NOTIONAL_PAYMENTS_B)
  Call AddParentedReport(RPT_MANAGEMENT_VOUCHERS_AND_CREDIT_CARDS_C, sParentName, BC_VOUCHERS_AND_CREDITCARDS_C)
  Call AddParentedReport(RPT_MANAGEMENT_LIVING_ACCOMODATION_D, sParentName, BC_LIVING_ACCOMMODATION_D)
  Call AddParentedReport(RPT_MANAGEMENT_EMPLOYEE_CARS_E, sParentName, BC_EMPLOYEE_CAR_E)
  'compay cars header
  'all headers to summerise benefit summaries must use menu caption as Keys will get mixed up
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR, sParentName, BC_COMPANY_CARS_F, p11d32.Rates.BenefitMenuCaption(BC_COMPANY_CARS_F), True, True)
  'reports
  'chilren of RPT_MANAGEMENT_COMPANY_CAR
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_SUMMARY_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Company car summary")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_AND_FUEL_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Car & fuel benefits")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_DETAILS_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Company car details") 'ek add new company car management report
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_CLASS1A_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Class 1A NIC")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_1_SUMMARY_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "P46 Quarter 1 summary")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_2_SUMMARY_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "P46 Quarter 2 summary")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_3_SUMMARY_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "P46 Quarter 3 summary")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_4_SUMMARY_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "P46 Quarter 4 summary")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_FUEL_REP_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Car fuel details")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_FUEL_VAT_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Car fuel VAT")
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_MISSING_MAKE_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Company cars missing make + list price")
'AM To be removed  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_MILES, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Company cars miles")
  'RH
  Call AddParentedReport(RPT_MANAGEMENT_COMPANY_CAR_NO_EMMISIONS_F, m_RepSettings(RPT_MANAGEMENT_COMPANY_CAR).Name, BC_COMPANY_CARS_F, "Company cars with no CO2 figure")
  'end compny cars
  Call AddParentedReport(RPT_MANAGEMENT_VANS_G, sParentName, BC_SHAREDVAN_G, "Vans")
  
  'km - home loans only apply to v2001 and below
'  If p11d32.AppYear <= 2001 Then Call AddParentedReport(RPT_MANAGEMENT_LOANS_HOME_H, sParentName, BC_LOAN_HOME_H)
    
  Call AddParentedReport(RPT_MANAGEMENT_LOANS_OTHER_H, sParentName, BC_LOAN_OTHER_H)
  
  Call AddParentedReport(RPT_MANAGEMENT_PRIVATE_MEDICAL_I, sParentName, BC_PRIVATE_MEDICAL_I)
  Call AddParentedReport(RPT_MANAGEMENT_QUALIFYING_RELOCATION_J, sParentName, BC_QUALIFYING_RELOCATION_J)
  Call AddParentedReport(RPT_MANAGEMENT_SERVICES_PROVIDED_K, sParentName, BC_SERVICES_PROVIDED_K)
  Call AddParentedReport(RPT_MANAGEMENT_ASSETS_AT_DISPOSAL_L, sParentName, BC_ASSETSATDISPOSAL_L)
  Call AddParentedReport(RPT_MANAGEMENT_NURSERY_OTHER_M, sParentName, BC_NON_CLASS_1A_M)
  Call AddParentedReport(RPT_MANAGEMENT_SUBSCRIPTION_AND_PROFESSIONAL_M, sParentName, BC_CLASS_1A_M)
  Call AddParentedReport(RPT_MANAGEMENT_INCOME_TAX_PAID_NOT_DEDUCTED_M, sParentName, BC_INCOME_TAX_PAID_NOT_DEDUCTED_M)
  
  sParentName2 = p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_HMIT_SECTION_STRING) & S_O_EXPENSES_CAPTION
  Call AddParentedReport(RPT_MANAGEMENT_O, sParentName, BC_OOTHER_N, sParentName2, True, True)
  Call AddParentedReport(RPT_MANAGEMENT_TRAVEL_AND_SUBSISTENCE_N, sParentName2, BC_TRAVEL_AND_SUBSISTENCE_N)
  Call AddParentedReport(RPT_MANAGEMENT_ENTERTAINMENT_N, sParentName2, BC_ENTERTAINMENT_N)

  'Call AddParentedReport(RPT_MANAGEMENT_GENERAL_EXPENSES_BUSINESS_N, sParentName2, BC_GENERAL_EXPENSES_BUSINESS_N)
  Call AddParentedReport(RPT_MANAGEMENT_PHONE_HOME_N, sParentName2, BC_PHONE_HOME_N)
  Call AddParentedReport(RPT_MANAGEMENT_NON_QUALIFYING_RELOCATION_N, sParentName2, BC_NON_QUALIFYING_RELOCATION_N)
  Call AddParentedReport(RPT_MANAGEMENT_CHAUFFEUR_N, sParentName2, BC_CHAUFFEUR_OTHERO_N)
  Call AddParentedReport(RPT_MANAGEMENT_OOTHER_N, sParentName2, BC_OOTHER_N)
    
  'FC - Abacus Export
  sParentName3 = "Abacus Export"
  m_RepSettings(RPT_ABACUSUDM).Name = sParentName3
  m_RepSettings(RPT_ABACUSUDM).GroupHeader = True

  m_RepSettings(RPT_ABACUSUDM_ASSETS_TRANSFERRED_A).Name = "Asset transferred"
  m_RepSettings(RPT_ABACUSUDM_ASSETS_TRANSFERRED_A).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_PAYMENTS_ON_BEHALF_B).Name = "Payments on behalf"
  m_RepSettings(RPT_ABACUSUDM_PAYMENTS_ON_BEHALF_B).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_TAX_ON_NOTIONAL_PAYMENTS_B).Name = "Tax on notional payments"
  m_RepSettings(RPT_ABACUSUDM_TAX_ON_NOTIONAL_PAYMENTS_B).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_VOUCHERS_AND_CREDIT_CARDS_C).Name = "Vouchers or credit cards"
  m_RepSettings(RPT_ABACUSUDM_VOUCHERS_AND_CREDIT_CARDS_C).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_LIVING_ACCOMODATION_D).Name = "Living accommodation"
  m_RepSettings(RPT_ABACUSUDM_LIVING_ACCOMODATION_D).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_EMPLOYEE_CARS_E).Name = "Employee cars"
  m_RepSettings(RPT_ABACUSUDM_EMPLOYEE_CARS_E).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_COMPANY_CAR_F).Name = "Company cars"
  m_RepSettings(RPT_ABACUSUDM_COMPANY_CAR_F).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_FUEL_F).Name = "Fuel"
  m_RepSettings(RPT_ABACUSUDM_FUEL_F).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_VANS_G).Name = "Vans"
  m_RepSettings(RPT_ABACUSUDM_VANS_G).ParentName = sParentName3
  'AM To be removed
'  m_RepSettings(RPT_ABACUSUDM_LOANS_HOME_H).Name = "Home loans"
'  m_RepSettings(RPT_ABACUSUDM_LOANS_HOME_H).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_LOANS_OTHER_H).Name = "Beneficial loans"
  m_RepSettings(RPT_ABACUSUDM_LOANS_OTHER_H).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_PRIVATE_MEDICAL_I).Name = "Private medical"
  m_RepSettings(RPT_ABACUSUDM_PRIVATE_MEDICAL_I).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_QUALIFYING_RELOCATION_J).Name = "Qualifying relocation"
  m_RepSettings(RPT_ABACUSUDM_QUALIFYING_RELOCATION_J).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_SERVICES_PROVIDED_K).Name = "Services provided"
  m_RepSettings(RPT_ABACUSUDM_SERVICES_PROVIDED_K).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_ASSETS_AT_DISPOSAL_L).Name = "Asset at disposal"
  m_RepSettings(RPT_ABACUSUDM_ASSETS_AT_DISPOSAL_L).ParentName = sParentName3
  
  m_RepSettings(RPT_ABACUSUDM_NURSERY_OTHER_M).Name = "Other Items - Non Class 1A"
  m_RepSettings(RPT_ABACUSUDM_SUBSCRIPTIONS_AND_PROFESSIONAL_M).Name = "Other Items - Class 1A"
  m_RepSettings(RPT_ABACUSUDM_NURSERY_OTHER_M).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_SUBSCRIPTIONS_AND_PROFESSIONAL_M).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_INCOME_TAX_PAID_NOT_DEDUCTED_M).Name = "Income tax paid but not deducted"
  m_RepSettings(RPT_ABACUSUDM_INCOME_TAX_PAID_NOT_DEDUCTED_M).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_TRAVEL_AND_SUBSISTENCE_N).Name = "Travel and subsistence"
  m_RepSettings(RPT_ABACUSUDM_TRAVEL_AND_SUBSISTENCE_N).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_ENTERTAINMENT_N).Name = "Entertainment"
  m_RepSettings(RPT_ABACUSUDM_ENTERTAINMENT_N).ParentName = sParentName3
  'm_RepSettings(RPT_ABACUSUDM_GENERAL_EXPENSES_BUSINESS_N).Name = "General business expenses"
  'm_RepSettings(RPT_ABACUSUDM_GENERAL_EXPENSES_BUSINESS_N).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_PHONE_HOME_N).Name = "Home phones"
  m_RepSettings(RPT_ABACUSUDM_PHONE_HOME_N).ParentName = sParentName3
  m_RepSettings(RPT_ABACUSUDM_NON_QUALIFYING_RELOCATION_N).Name = "Non-qualifying relocation"
  m_RepSettings(RPT_ABACUSUDM_NON_QUALIFYING_RELOCATION_N).ParentName = sParentName3
  'm_RepSettings(RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_O).Name = "Chauffeur Expenses"         'rh
  'm_RepSettings(RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_O).ParentName = sParentName3           'rh
  'm_RepSettings(RPT_ABACUSUDM_NOTHER_O).Name = "N Other"                     'rh
  'm_RepSettings(RPT_ABACUSUDM_NOTHER_O).ParentName = sParentName3                       'rh
  m_RepSettings(RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_N_OTHER).Name = p11d32.Rates.BenClassTo(BC_CHAUFFEUR_OTHERO_N, BCT_FORM_CAPTION)
  m_RepSettings(RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_N_OTHER).ParentName = sParentName3
  
  m_RepSettings(RPT_ABACUSUDM_OOTHER_N).Name = "Other"
  m_RepSettings(RPT_ABACUSUDM_OOTHER_N).ParentName = sParentName3

  'Header
  m_RepSettings(RPT_USER).Name = "User reports"
  m_RepSettings(RPT_USER).GroupHeader = True
  'reports
  
  PRNLoopLength = 1000
  
Initialise_ReportPrint_END:
  Call xReturn("Initialise_ReportPrint")
  Exit Sub
Initialise_ReportPrint_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Initialise_ReportPrint", "Initialise ReportPrint", "Error initialising the report print class.")
  Resume Initialise_ReportPrint_END
End Sub

Private Sub AddParentedReport(rpt As RPT_TYPE, sParentName As String, ByVal bc As BEN_CLASS, Optional sCaption As String = "", Optional bCaptionOnly As Boolean = False, Optional bGroupHeader As Boolean = False)
  m_RepSettings(rpt).Name = ManagementReportCaption(bc, sCaption, bCaptionOnly)
  m_RepSettings(rpt).ParentName = sParentName
  m_RepSettings(rpt).GroupHeader = bGroupHeader
End Sub
Public Property Get Name(ByVal ReportIndex As P11D_REPORTS) As String
    Name = m_RepSettings(ReportIndex).Name
  If Me.ExportNoSpaces Then
    Name = Replace(Name, " ", "")
  End If
End Property
Public Property Get GroupHeader(ByVal ReportIndex As P11D_REPORTS) As Boolean
  GroupHeader = m_RepSettings(ReportIndex).GroupHeader
End Property
Public Property Get ParentName(ByVal ReportIndex As P11D_REPORTS) As String
  ParentName = m_RepSettings(ReportIndex).ParentName
End Property

Public Property Get Orientation(ByVal ReportIndex As P11D_REPORTS) As REPORT_ORIENTATION
  Orientation = m_RepSettings(ReportIndex).Orientation
End Property
Public Property Let Orientation(ByVal ReportIndex As P11D_REPORTS, ByVal NewValue As REPORT_ORIENTATION)
  m_RepSettings(ReportIndex).Orientation = NewValue
End Property
Public Property Get TemplatePathAndFile()
  CheckMReportsTemplatePathAndFile
  TemplatePathAndFile = FullPath(ReportPathUser) & "Template" & S_REPORT_FILE_EXTENSION
End Property

Public Sub StartReportWizard(ByVal sTemplateFile As String, ByVal RWT As REPORTW_TARGET, Optional ByVal sExportPathAndFile As String)
  Dim RW As ReportWizard
  Dim benEmployer As IBenefitClass
  Dim sSourceDirectory As String
  Dim sExportPath As String, sExportFile As String, sExportExtension As String
  
  
  On Error GoTo StartReportWizard_ERR
  
  Call xSet("StartReportWizard")
  Set RW = ReportWizardNew()
  RW.DoReportOnCancel = False
  Set RW.ReportInterface = p11d32.udm
  Call p11d32.udm.Initialise
  
  If Not FileExists(sTemplateFile) Then Call ErrorMessage(ERR_FILE_NOT_EXIST, Err, "StartReportWizard", "Start Report Wizard", "The report template file " & TemplatePathAndFile & " does not exist.")
  
  'RK Set appropriate directories for Report Wizard
  Call SplitPath(sTemplateFile, sSourceDirectory)
  RW.ReportFileName = sTemplateFile
  RW.ReportFilesPath = sSourceDirectory
      
  Set benEmployer = p11d32.CurrentEmployer
  If RWT = REPORTW_PREPARE_EXPORT And (Len(sExportPathAndFile) > 0) Then
    RW.ExportType = EXPORT_CSV
    RW.NotifyLineMeetsCriteria = True 'this is no strictly universal to all callers
    Call SplitPath(sExportPathAndFile, sExportPath, sExportFile, sExportExtension)
    RW.ExportPath = sExportPath
    RW.ExportFileName = sExportFile & sExportExtension
  End If
  'CAD need to think for future
  'Call RW.AddControlCode("{REPORT_BANNER}", HMITBannerString("Enter you report title here", False), "Report Banner")
  Call RW.AddControlCode(EmployeeLetterCode(ELC_COMPANY_NAME, ELCT_LETTER_FILE_CODES, False), benEmployer.Name, "Employer Name")
  Call RW.AddControlCode(EmployeeLetterCode(ELC_PAYE_REF, ELCT_LETTER_FILE_CODES, False), benEmployer.value(employer_Payeref_db), "PAYE Ref")
  Call RW.AddControlCode(EmployeeLetterCode(ELC_DATE_TAXYEAR, ELCT_LETTER_FILE_CODES, False), p11d32.Rates.value(TaxFormYear), "Tax Year")
  
  
  RW.AutoPreviewTarget = RWT
  Call RW.StartReportWizard

StartReportWizard_END:
  Call xReturn("StartReportWizard")
  Exit Sub
  
StartReportWizard_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "StartReportWizard", "Start Report Wizard", "Error starting the report wizard.")
  Resume StartReportWizard_END
  Resume
End Sub

Public Sub DoExportAbacusReportWizard()
  Dim sExportFile As String, i As Long
  Dim tmpFile As String, fr As TCSFileread
  Dim sFileBuf As String, sFilesBuf As String, sFIle  As String
  Dim RW As ReportWizard, es As EMPLOYEE_SELECTION
  Dim benEmployer As IBenefitClass
  Dim FirstFileLine As Boolean
  
  On Error GoTo ExportAbacusReportWizard_ERR
  
  Call xSet("ExportAbacusReportWizard")
  Call SetCursor
  es = Me.EmployeeSelection
  Me.EmployeeSelection = ES_ALL
  Set benEmployer = p11d32.CurrentEmployer
  sFIle = benEmployer.Name & ".csv"
  sExportFile = BrowseForFolderEx(MDIMain.hwnd, CurDir$, "Please choose a directory for the output file [" & sFIle & "]")
  If Len(sExportFile) = 0 Then GoTo ExportAbacusReportWizard_END
  Call PrgStartCaption((RPT_LAST_ABACUSUDM - RPT_FIRST_ABACUSUDM) + 1, "Starting export")
  sFIle = FullPath(sExportFile) & sFIle
  For i = RPT_FIRST_ABACUSUDM To RPT_LAST_ABACUSUDM
'AM To be removed    If Not (p11d32.AppYear > 2000 And (i = RPT_ABACUSUDM_EDUCATION_OTHER_N)) Then  'AM
      Call PrgStepCaption("Starting export of report " & Name(i))
      tmpFile = GetTempFilename(CoreClass.GetTempDirectory(), "P11D")
      Set RW = ReportWizardNew()
      RW.UseCancel = True
      Set RW.ReportInterface = p11d32.udm
      Call p11d32.udm.Initialise
      RW.ReportFileName = ManagementReportPathAndFile(i)
      Set benEmployer = p11d32.CurrentEmployer
      RW.AutoPreviewTarget = REPORTW_PREPARE_EXPORT
      RW.ExportFileName = tmpFile
      RW.ExportType = EXPORT_CSV
      Call RW.StartReportWizard
      Set RW = Nothing
          
      Set fr = New TCSFileread
      If Not fr.OpenFile(tmpFile) Then Err.Raise ERR_FILE_OPEN, "ExportAbacusReportWizard", "File [" & tmpFile & "] does not exist"
      FirstFileLine = True
      Do While fr.GetLine(sFileBuf)
        If Len(sFileBuf) > 0 Then
          If Not FirstFileLine Or (Len(sFilesBuf) = 0) Then
            sFilesBuf = sFilesBuf & sFileBuf & vbCrLf
          End If
          FirstFileLine = False
        End If
      Loop
'    End If
  Next i
 
  Call PrgStopCaption
  If Len(sFilesBuf) > 0 Then
    Call TextFileSave(sFIle, sFilesBuf)
    
  Else
    Err.Raise ERR_REPORT_INVALID, "ExportAbacusReportWizard", "No information found for this employer"
  End If
  
  Call DisplayMessage(MDIMain, "Abacus export report finished correctly" & vbCrLf & "Save in file [" & sFIle & "]", "Abacus export", "Ok", "")
  
ExportAbacusReportWizard_END:
  Call PrgStopCaption
  Me.EmployeeSelection = es
  Call ClearCursor
  Call xReturn("ExportAbacusReportWizard")
  Exit Sub
  
ExportAbacusReportWizard_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ExportAbacusReportWizard", "Execute Batch Report Wizard", "Abacus export report failed to execute correctly")
  Resume ExportAbacusReportWizard_END
  Resume
End Sub

Public Function InitPrintDialog() As Boolean
  Dim dev As String, port As String
  Dim i As Long
  On Error GoTo InitPrintDialog_Err
  
  Call xSet("InitPrintDialog")
  
  If Not p11d32.MoveMenuUpdateEmployeeCheckEmployer Then GoTo InitPrintDialog_End
  
  'which screen are we on
  m_PrinterAvailable = GetDefaultPrinter(dev, port)
  If m_PrinterAvailable Then
    F_Print.lblName = dev & " on " & port
  Else
    F_Print.lblName = "(No current printer)"
  End If
  F_Print.lblType = dev
  F_Print.lblWhere = port
  
' F_Print.Show vbModal
  Call p11d32.Help.ShowForm(F_Print, vbModal)
  
InitPrintDialog_End:
  Call xReturn("InitPrintDialog")
  Exit Function

InitPrintDialog_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitPrintDialog", "Initialise Printer Dialogue", "Unable to initialise printer dialogue.")
  Resume InitPrintDialog_End
  Resume
End Function
Private Sub EndPrintDialog()
  On Error Resume Next
  m_PrinterAvailable = False
  Unload F_Print
  Set F_Print = Nothing
End Sub
Public Function ReportType(ByVal rpt As P11D_REPORTS, Optional bErrors As Boolean = False) As RPT_TYPE
  If rpt >= RPT_FIRST_STANDARD And rpt <= RPT_LAST_STANDARD Then
    ReportType = RPTT_STANDARD 'need collection of employees
  ElseIf rpt >= [RPT_FIRST_MANAGEMENT] And rpt <= [RPT_LAST_MANAGEMENT] Then
    ReportType = RPTT_MANAGEMENT 'fire off report wizard
  ElseIf rpt >= [RPT_FIRST_OTHER_REPORT] And rpt <= [RPT_LAST_OTHER_REPORT] Then
    ReportType = RPTT_OTHER 'not dependant on employees
  'FC - AE - 160801
  ElseIf rpt >= RPT_FIRST_ABACUSUDM And rpt <= RPT_LAST_ABACUSUDM Then
    ReportType = RPTT_ABACUSUDM
  ElseIf rpt = -1 Then
    ReportType = RPTT_USER
  Else
    If bErrors Then Call ECASE("Invalid RPT passed to ReportType.")
    ReportType = RPTT_UNKNOWN
  End If
End Function
Public Property Get ReportPathManagement() As String
  ReportPathManagement = FullPath(FullPath(AppPath) & S_SYSTEMDIR_MREPORTS)
End Property
Public Property Get ReportPathUser() As String
  ReportPathUser = m_ReportPathUser
End Property
Public Property Let ReportPathUser(ByVal NewValue As String)
  m_ReportPathUser = NewValue
End Property
Public Property Get UserLettersPath() As String
  UserLettersPath = m_UsersLettersPath
End Property
Public Property Let UserLettersPath(ByVal NewValue As String)
  m_UsersLettersPath = NewValue
End Property
Public Property Get SystemLettersPath()
  SystemLettersPath = FullPath(AppPath) & S_SYSTEMDIR_LETTERS
End Property

Public Sub DoWizardReport(ByVal ReportType As P11D_REPORTS, ByVal ReportDest As REPORT_TARGET)
  Dim sFIle As String
  Dim sPathAndFile As String
  
  On Error GoTo DoWizardReport_ERR
  
  Call xSet("DoWizardReport")
  
  sPathAndFile = ManagementReportPathAndFile(ReportType)
    
  Select Case ReportDest
    Case REPORT_TARGET.PREPARE_REPORT
      Call StartReportWizard(sPathAndFile, REPORTW_PREVIEW)
    Case REPORT_TARGET.PRINT_REPORT
      Call StartReportWizard(sPathAndFile, REPORTW_PRINT)
    Case Else
      ECASE ("Invalid report target.")
  End Select
  
DoWizardReport_END:
  
  Call xReturn("DoWizardReport")
  Exit Sub
DoWizardReport_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DoWizardReport", "Do Wizard Report", "Error selecting the wizard report " & ReportType & ".")
  Resume DoWizardReport_END
  Resume
End Sub
Public Function ManagementReportPathAndFile(ByVal ReportType As P11D_REPORTS, Optional ByVal sUserReportNameLessFileExtensionOverride As String = "") As String
  Dim sFIle As String
  
  On Error GoTo ManagementReportPathAndFile_ERR
  
  Select Case ReportType
    Case RPT_MANAGEMENT_NIC_CLASS1A_EMPLOYER
      sFIle = "1ASum"
    Case RPT_MANAGEMENT_NIC_CLASS1A_EMPLOYEE
      sFIle = "E1ASum"
    Case RPT_MANAGEMENT_BENEFIT_SUMMARY
      sFIle = "Ben_Su"
    Case RPT_MANAGEMENT_EMPLOYEE_LIST
      sFIle = "Emp_Li"
    Case RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_1_SUMMARY_F
      sFIle = "F_P461"
    Case RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_2_SUMMARY_F
      sFIle = "F_P462"
    Case RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_3_SUMMARY_F
      sFIle = "F_P463"
    Case RPT_MANAGEMENT_COMPANY_CAR_P46_QUARTER_4_SUMMARY_F
      sFIle = "F_P464"
    Case RPT_MANAGEMENT_ASSETS_TRANSFERRED_A
      sFIle = "A_Tran"
    Case RPT_MANAGEMENT_PAYMENTS_ON_BEHALF_B
      sFIle = "B_Paym"
    Case RPT_MANAGEMENT_TAX_ON_NOTIONAL_PAYMENTS_B
      sFIle = "B_NTax"
    Case RPT_MANAGEMENT_VOUCHERS_AND_CREDIT_CARDS_C
      sFIle = "C_Vouc"
    Case RPT_MANAGEMENT_LIVING_ACCOMODATION_D
      sFIle = "D_Acco"
    Case RPT_MANAGEMENT_EMPLOYEE_CARS_E
      sFIle = "E_EECa"
    'company cars
    Case RPT_MANAGEMENT_COMPANY_CAR_SUMMARY_F
      sFIle = "F_CarS"
    Case RPT_MANAGEMENT_COMPANY_CAR_AND_FUEL_F
      sFIle = "F_CoCa"
    Case RPT_MANAGEMENT_COMPANY_CAR_DETAILS_F
      sFIle = "F_CDet"
    Case RPT_MANAGEMENT_COMPANY_CAR_CLASS1A_F
      sFIle = "F_Cl1A"
    Case RPT_MANAGEMENT_COMPANY_CAR_FUEL_REP_F
      sFIle = "F_CFUE"
    Case RPT_MANAGEMENT_COMPANY_CAR_FUEL_VAT_F
      sFIle = "F_VAT"
    Case RPT_MANAGEMENT_COMPANY_CAR_MISSING_MAKE_F
      sFIle = "F_CarM"
'AM To be removed
'    Case RPT_MANAGEMENT_COMPANY_CAR_MILES
'      sFile = "F_CoMi"
    'RH
    Case RPT_MANAGEMENT_COMPANY_CAR_NO_EMMISIONS_F
      sFIle = "F_CarE"
    'end cars
    Case RPT_MANAGEMENT_VANS_G
      sFIle = "G_Vans"
'AM To be removed
'    Case RPT_MANAGEMENT_LOANS_HOME_H
'      sFile = "H_LoaH"
    Case RPT_MANAGEMENT_LOANS_OTHER_H
      sFIle = "H_LoaO"
    Case RPT_MANAGEMENT_PRIVATE_MEDICAL_I
      sFIle = "I_Medi"
    Case RPT_MANAGEMENT_QUALIFYING_RELOCATION_J
      sFIle = "J_Relo"
    Case RPT_MANAGEMENT_SERVICES_PROVIDED_K
      sFIle = "K_Serv"
    Case RPT_MANAGEMENT_ASSETS_AT_DISPOSAL_L
      sFIle = "L_Asse"
    Case RPT_MANAGEMENT_NURSERY_OTHER_M
      sFIle = "M_Nurs"
    Case RPT_MANAGEMENT_SUBSCRIPTION_AND_PROFESSIONAL_M
      sFIle = "M_Subs"
    Case RPT_MANAGEMENT_INCOME_TAX_PAID_NOT_DEDUCTED_M
      sFIle = "M_Tax"
    Case RPT_MANAGEMENT_TRAVEL_AND_SUBSISTENCE_N
      sFIle = "N_Trav"
    Case RPT_MANAGEMENT_ENTERTAINMENT_N
      sFIle = "N_Ente"
    'Case RPT_MANAGEMENT_GENERAL_EXPENSES_BUSINESS_N
    '  sFIle = "N_Gnrl"
    Case RPT_MANAGEMENT_PHONE_HOME_N
      sFIle = "N_Phon"
    Case RPT_MANAGEMENT_NON_QUALIFYING_RELOCATION_N
      sFIle = "N_NQRe"
    Case RPT_MANAGEMENT_CHAUFFEUR_N
      sFIle = "N_Chau"
    Case RPT_MANAGEMENT_OOTHER_N
      sFIle = "N_Othe"
    'Abacus Export
    Case RPT_ABACUSUDM_ASSETS_TRANSFERRED_A
      sFIle = "AE_AsT"
    Case RPT_ABACUSUDM_PAYMENTS_ON_BEHALF_B
      sFIle = "AE_PBE"
    Case RPT_ABACUSUDM_TAX_ON_NOTIONAL_PAYMENTS_B
      sFIle = "AE_TNP"
    Case RPT_ABACUSUDM_VOUCHERS_AND_CREDIT_CARDS_C
     sFIle = "AE_VCC"
    Case RPT_ABACUSUDM_LIVING_ACCOMODATION_D
      sFIle = "AE_LIV"
    Case RPT_ABACUSUDM_EMPLOYEE_CARS_E
      sFIle = "AE_EeC"
    Case RPT_ABACUSUDM_COMPANY_CAR_F
      sFIle = "AE_CoC"
    Case RPT_ABACUSUDM_FUEL_F
      sFIle = "AE_CFu"
    Case RPT_ABACUSUDM_VANS_G
      sFIle = "AE_VAN"
    'AM To be removed
'    Case RPT_ABACUSUDM_LOANS_HOME_H
'      sFile = "AE_LHO"
    Case RPT_ABACUSUDM_LOANS_OTHER_H
      sFIle = "AE_LOT"
    Case RPT_ABACUSUDM_PRIVATE_MEDICAL_I
      sFIle = "AE_MED"
    Case RPT_ABACUSUDM_QUALIFYING_RELOCATION_J
      sFIle = "AE_QRE"
    Case RPT_ABACUSUDM_SERVICES_PROVIDED_K
      sFIle = "AE_SPR"
    Case RPT_ABACUSUDM_ASSETS_AT_DISPOSAL_L
      sFIle = "AE_AsD"
'AM To be removed
'    Case RPT_ABACUSUDM_EDUCATION_OTHER_N
'      sFile = "AE_EDU"  'AM
    Case RPT_ABACUSUDM_NURSERY_OTHER_M
      sFIle = "AE_NUR"
    Case RPT_ABACUSUDM_SUBSCRIPTIONS_AND_PROFESSIONAL_M
      sFIle = "AE_SUB"
    Case RPT_ABACUSUDM_INCOME_TAX_PAID_NOT_DEDUCTED_M
      sFIle = "AE_ITP"
    Case RPT_ABACUSUDM_TRAVEL_AND_SUBSISTENCE_N
      sFIle = "AE_TRV"
    Case RPT_ABACUSUDM_ENTERTAINMENT_N
      sFIle = "AE_ENT"
    'Case RPT_ABACUSUDM_GENERAL_EXPENSES_BUSINESS_N
    '  sFIle = "AE_BUS"
    Case RPT_ABACUSUDM_PHONE_HOME_N
      sFIle = "AE_PHO"
    Case RPT_ABACUSUDM_NON_QUALIFYING_RELOCATION_N
      sFIle = "AE_NQR"
    Case RPT_ABACUSUDM_CHAUFFEUR_EXPENSES_N_OTHER
      sFIle = "AE_NTh"
    Case RPT_ABACUSUDM_OOTHER_N
      sFIle = "AE_OTH"
    Case -1
      'User reports
      If Len(sUserReportNameLessFileExtensionOverride) = 0 Then
        sUserReportNameLessFileExtensionOverride = UserReportFileLessExtension
      End If
      If Len(sUserReportNameLessFileExtensionOverride) = 0 Then Call Err.Raise(ERR_USER_REPORT, "DoWizardReport", "User report selected but file name is zero length string.")
      ManagementReportPathAndFile = FullPath(p11d32.ReportPrint.ReportPathUser) & sUserReportNameLessFileExtensionOverride & S_REPORT_FILE_EXTENSION
      GoTo ManagementReportPathAndFile_END
    Case Else
      ECASE ("Report requested " & ReportType & " is not valid in DoWizardReport.")
  End Select
  
  If ReportType <> -1 Then sFIle = sFIle
  
  ManagementReportPathAndFile = ReportPathManagement & sFIle & S_REPORT_FILE_EXTENSION
  
ManagementReportPathAndFile_END:
  Exit Function
ManagementReportPathAndFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ManagementReportPathAndFile", "Management Report File Name", "Error getting a management report file name.")
  Resume ManagementReportPathAndFile_END
End Function
Public Function DoStandardReport(ByVal pr As P11D_REPORTS, ByVal ReportDest As REPORT_TARGET, Optional rep As Reporter = Nothing, Optional Employee As Employee = Nothing) As Boolean
  Dim sReportName As String
  Dim s As String
  Dim rd As REPORT_DESTINATION
  On Error GoTo DoStandardReport_Err
  
  Call xSet("DoStandardReport")
  
  rd = Me.Destination
  
  If SelectedEmployees.Count = 0 Then
    If Not p11d32.CurrentEmployeeIsNothing Then
      Call SelectedEmployees.Add(p11d32.CurrentEmployer.CurrentEmployee)
    Else
      Call Err.Raise(ERR_DOREPORT, "DoStandardReport", "Unable to excute report." & vbCrLf & "No employees selected for printing.")
    End If
  End If
  
  sReportName = Name(pr)
  If IsEmail(pr) Then
    If Not FileExists(p11d32.TempPath, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "DoReportStandard", "The temp directory does not exist = " & p11d32.TempPath & ", unable to email.")
    Me.Destination = REPD_PRINTER_PREVIEW
  End If
  
  If (Me.Destination = REPD_FILE_HTML) Or (Me.Destination = REPD_FILE_PRN) Then
    If Not FileExists(p11d32.ReportPrint.ExportDirectory, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "DoReportStandard", "The export directory does not exist = " & p11d32.ReportPrint.ExportDirectory & ", unable to export to file.")
    If Me.ExportAllowUserNameHTML Then
      Do Until Len(s) > 0
        s = InputBox("Please enter report name for export....", "Report Name", sReportName)
      Loop
      sReportName = s
    End If
  End If
  
  DoStandardReport = F_PrintCancel.DoReportStandardSub(rep, pr, ReportDest, Orientation(pr), sReportName)
  
  
  
  If IsEmail(pr) Then
    Call DestroyATCMAIL
  End If
  
DoStandardReport_End:
  Me.Destination = rd
  Call xReturn("DoStandardReport")
  Exit Function
DoStandardReport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DoStandardReport", "Execute Report", "Unable to execute report, report = " & pr)
  Resume DoStandardReport_End
  Resume
End Function

Public Function DoOtherReport(ByVal rep As Reporter, ByVal ReportType As P11D_REPORTS, ByVal ReportDest As REPORT_TARGET, Optional Employee As Employee = Nothing) As Boolean
  Dim InError As Boolean
  Dim s As String
  Dim sReportName As String
  
  
  On Error GoTo DoOtherReport_Err
  
  Call xSet("DoOtherReport")
  
  If rep Is Nothing Then Set rep = ReporterNew()
  If Not rep.InitReport("P11D report:" & Name(ReportType), ReportDest, Orientation(ReportType), True) Then Call Err.Raise(ERR_DOREPORT, "DoOtherReport", "Unable to initialise Reporter." & vbCrLf & "Unable to initialise report engine.")
  
  Select Case ReportType
    Case RPT_P11Db
      Call Report_P11db(rep, p11d32.CurrentEmployer)
    Case RPT_MM_SUBMISSION
      Call Report_MM_Submission(rep, p11d32.CurrentEmployer)
    Case Else
      ECASE ("Invalid other report")
  End Select
  DoOtherReport = True
  
DoOtherReport_End:
  If Not rep Is Nothing Then
    If InError Then rep.AbortReport
    
    If (Not InError) And Me.Destination = REPD_FILE_HTML Then
      Call ExportReport(rep, Nothing, sReportName = m_RepSettings(ReportType).Name)
    End If
    Call rep.EndReport(InError)
    
    If (Not InError) And (ReportDest = PREPARE_REPORT) Then
      rep.PreviewReport
    End If
      
    Set rep = Nothing
  End If
  Call xReturn("DoOtherReport")
  Exit Function
  
DoOtherReport_Err:
  InError = True
  Call ErrorMessage(ERR_ERROR, Err, "DoOtherReport", "Do Other Report", "Unable to execute other report, report index = " & ReportType & ".")
  Resume DoOtherReport_End
End Function
Public Property Get NoOfEmployeesSelected(Employer As Employer) As Long
  Dim ben As IBenefitClass
  Dim l As Long
  Dim i As Long
  
  On Error GoTo NoOfEmployeesSelected_ERR
  
  Call xSet("NoOfEmployeesSelected")
  
  If Employer Is Nothing Then Call Err.Raise(ERR_EMPLOYER_INVALID, "NoOfEmployeesSelected", "Employer is nothing")
  
  Set ben = Employer
  
  Select Case EmployeeSelection
    Case ES_ALL
      l = Employer.NoOfEmployees
    Case ES_INVERSE_SELECTED
      l = Employer.NoOfEmployees
      l = l - ben.value(employer_NoOfSelectedEmployees)
    Case ES_SELECTED
      l = ben.value(employer_NoOfSelectedEmployees)
    Case ES_CURRENT
      If (Not p11d32.CurrentEmployer.CurrentEmployee Is Nothing) Then
        l = 1
      End If
    Case Else
      Call ECASE("Invalid employee selection:" & EmployeeSelection)
  End Select
  
  NoOfEmployeesSelected = l
  
NoOfEmployeesSelected_END:
  Call xReturn("NoOfEmployeesSelected")
  Exit Function
NoOfEmployeesSelected_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NoOfEmployeesSelected", "No Of Employees Selected", "Error finding out how many employees are selected for print.")
  Resume NoOfEmployeesSelected_END
  Resume
End Property

Public Function NonZeroTest(ee As Employee) As Boolean
  Dim b As Boolean
  
  On Error GoTo NonZeroTest_ERR
  
  Call xSet("NonZeroTest")
  
  b = ee.PrintPrepare
  Select Case P11DTotalValue
    Case PTV_NONZERO_AND_ZERO
      NonZeroTest = True
    Case PTV_NONZERO
      NonZeroTest = b
    Case PTV_ZERO
      NonZeroTest = Not b
    Case Else
      Call ECASE("Invalid P11D_TOTAL_VALUE in NonZeroTest")
  End Select
  
NonZeroTest_END:
  Call xReturn("NonZeroTest")
  Exit Function
NonZeroTest_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NonZeroTest", "Non Zero Test", "Error determining whether to print depending on the print options 'total benefit' status.")
  Resume NonZeroTest_END
End Function
Private Function LetterFileSystem(ByVal sFileNoPathNoExtension As String) As String
  Dim fr As TCSFileread
  Dim s As String
  
  On Error GoTo LetterFileSystem_ERR
  
  s = FullPath(p11d32.ReportPrint.SystemLettersPath) & "system\" & sFileNoPathNoExtension & S_EMPLOYEE_LETTER_FILE_EXTENSION
  
  If Not FileExists(s, False) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "SystemLetterFile", "The file " & s & " does not exist.")
  LetterFileSystem = GetFileText(s)
  
  Exit Function
LetterFileSystem_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "LetterFileSystem"), "Error in LetterFileSystem " & Err.Description)
End Function

Public Function StandardReportLetterFileText(ByVal pr As P11D_REPORTS) As String
  Dim qs As QString
  Dim s As String
  
  On Error GoTo StandardReportLetterFileText_ERR
  
  Set qs = New QString
  
  If pr <> RPT_EMPLOYEE_LETTER_EMAIL Then
    qs.Append (EmployeeLetterCode(ELC_NORMAL, ELCT_LETTER_FILE_CODES, False))
    'place normal at the front
  End If
  
  Select Case pr
      Case RPT_P46CAR
        s = LetterFileSystem("P46CAR")
      Case RPT_EMPLOYEE_LETTER
        s = GetFileText(p11d32.ReportPrint.EmployeeLetterPathAndFile)
      Case RPT_EMPLOYEE_LETTER_EMAIL
        s = GetFileText(p11d32.ReportPrint.EmployeeLetterPathAndFile)
      Case RPT_EMPLOYEE_LETTER_WORKING_PAPERS
        s = GetFileText(p11d32.ReportPrint.EmployeeLetterPathAndFile)
      Case RPT_HMIT
        s = LetterFileSystem("P11D")
      Case RPT_HMIT_PLUS_WORKING_PAPERS
        s = LetterFileSystem("P11DWP")
      Case RPT_WORKINGPAPERS
        s = LetterFileSystem("WP")
      Case RPT_PRINTEDEMPLOYEES
      Case Else
        Call Err.Raise(ERR_REPORT_TEXT_TO_REP, "DoReport", "Unable to excute report: " & CStr(pr) & " no such report.")
  End Select
  
  
  Call qs.Append(s)
  
  
  If pr = RPT_EMPLOYEE_LETTER_WORKING_PAPERS Then
    qs.Append (EmployeeLetterCode(ELC_NEWPAGE, ELCT_LETTER_FILE_CODES, False))
    qs.Append (EmployeeLetterCode(ELC_WORKING_PAPERS, ELCT_LETTER_FILE_CODES, False))
  End If
  
  StandardReportLetterFileText = qs
  
StandardReportLetterFileText_END:
  Exit Function
StandardReportLetterFileText_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "StandardReportLetterFileText"), Err.Description)
End Function
Private Sub SetMargin(rep As Reporter, ByVal bOn As Boolean, bIsEmail As Boolean)
  Dim s As String
  
  If bIsEmail Then Exit Sub
  If bOn Then
    s = CStr(p11d32.ReportPrint.EmployeeLetterMargin)
  Else
    s = "0"
  End If
  Call rep.Out("{LEFTMARGIN=" & s & "}")
  
End Sub
Private Function EmployeeLetterCodeIsSubReport(ByVal ELC As EMPLOYEE_LETTER_CODE)
  EmployeeLetterCodeIsSubReport = ((ELC >= [ELC_FIRST_SUB_REPORT] And ELC <= [ELC_LAST_SUB_REPORT]))
End Function

Private Function EmployeeLetterCodeIsSubReportAttachment(ByVal ELC As EMPLOYEE_LETTER_CODE)
  EmployeeLetterCodeIsSubReportAttachment = (EmployeeLetterCodeIsSubReport(ELC) And (ELC <> ELC_TABLE))
End Function
Public Function IsEmail(ByVal pr As P11D_REPORTS)
  IsEmail = (pr = RPT_EMPLOYEE_LETTER_EMAIL)
End Function
Private Function IsFontLetterCode(ByVal ELC As EMPLOYEE_LETTER_CODE)
  IsFontLetterCode = ((ELC = ELC_BOLD) Or (ELC = ELC_NORMAL))
End Function
Private Sub WaitForPRN(sReportName As String, ByVal ee As Employee, sPRNPathAndFilePerDriver As String, ByVal rep As Reporter)
  Dim prg As Object
  Dim sPRNPathAndFileExport As String
  Dim bFileCopy As Boolean
  Dim bExit As Boolean
  Dim lValue As Long
  Dim lMax As Long
  Dim lIndicator As TCSPROG.Indicator
  Dim lAlignment As TCSPROG.TextAlignment
  Dim sCaption As String
  
'Wait for a length of time until prn file is generated
  On Error GoTo WaitForPRN_ERR
  
  'sPRNPathAndFilePerDriver must be deleted before this function and before 'EndReport
  
  Call xKill(sPRNPathAndFileExport)
  
  sPRNPathAndFileExport = Me.ExportFileName(sReportName, ee, rep)
  Set prg = MDIMain.sts.prg
  lValue = prg.value
  lMax = prg.Max
  lIndicator = prg.Indicator
  sCaption = prg.Caption
  lAlignment = prg.TextAlignment
  
  prg.Max = PRNLoopLength
  prg.Min = 0
  prg.Caption = "Waiting for printer to create: " & "'" & sPRNPathAndFilePerDriver & "'"
  prg.Indicator = None
  prg.TextAlignment = ALIGN_LEFT
  
'  Do Until bFileCopy RK Removed 04/06/03
'    If prg.value = prg.Max Then prg.value = prg.Min
'    Call prg.Step
'    'copy the file accross
'    If Not FileExists(sPRNPathAndFilePerDriver) Then  'AM
'      Call Err.Raise(ERR_WAIT_FOR_PRN, "WaitForPRN", "Specified PRN source file does not exist: " & sPRNPathAndFilePerDriver)
'    End If
'    bFileCopy = FileCopyEx(sPRNPathAndFilePerDriver, sPRNPathAndFileExport)
'  Loop
  
  ' Check for prn file for a length of time
  Do Until prg.value = prg.Max Or FileExists(sPRNPathAndFilePerDriver)
    Call prg.Step
    Sleep 10
  Loop
    
  If Not FileExists(sPRNPathAndFilePerDriver) Then  'RK amendments to Am
    Call Err.Raise(ERR_WAIT_FOR_PRN, "WaitForPRN", "Specified PRN source file does not exist: " & sPRNPathAndFilePerDriver)
  Else
    Call FileCopyEx(sPRNPathAndFilePerDriver, sPRNPathAndFileExport)
  End If

  
  
WaitForPRN_END:
  If Not prg Is Nothing Then
    prg.Max = lMax
    prg.value = lValue
    prg.Caption = sCaption
    prg.Indicator = lIndicator
    prg.TextAlignment = lAlignment
  End If
  Exit Sub
WaitForPRN_ERR:
  Call Err.Raise(Err.Number, "WaitForPRN", Err.Description)
End Sub
Public Sub ReportTextToRep(ByVal rep As Reporter, ByVal ee As Employee, ByVal pr As P11D_REPORTS, sEmployeeLetterText As String, ByVal RO As REPORT_ORIENTATION, ByVal bIsEmail As Boolean, sReportName As String) 'RK Email 19/03/03
'Public Sub ReportTextToRep(ByVal rep As Reporter, ByVal ee As employee, ByVal pr As P11D_REPORTS, sEmployeeLetterText As String, ByVal RO As REPORT_ORIENTATION, ByVal bIsEmail As Boolean, ByVal atms As MailAttachments, sReportName As String) 'RK Email 19/03/03
  Dim p0 As Long, p1 As Long
  Dim sCode As String
  Dim qs As QString
  Dim bSubReport As Boolean, sExportFile As String
  Dim i As Long
  Dim bFont As Boolean
  Dim sFont As String
  'Dim atm As MailAttachment 'RK Email 19/03/03
  Dim sPRNPathAndFilePerDriver As String
  
  'we now have the text for the letter file so replace elelments
  
  On Error GoTo ReportTextToRep_ERR
  
  Call xSet("EmployeeLetter")
  
  'If bIsEmail And atms Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ReportTextToRep", "Trying to email and atatchment collection is nothing.") 'RK Email 19/03/03
  If ee Is Nothing Then Call Err.Raise(ERR_NO_EMPLOYEE, "ReportTextToRep", "The employee is nothing.")
  If rep Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ReportTextToRep", "The reporter is nothing.")
   
  If Not bIsEmail And Not (p11d32.ReportPrint.ExportOption = EXPORT_HTML_INTEXP5) Then rep.PageFooter = TimeStampReport  'km
  Call SetMargin(rep, True, bIsEmail)
  Set qs = New QString
  
  p1 = 1
  p0 = InStr(p1, sEmployeeLetterText, "{", vbBinaryCompare)
  
  If (Not p11d32.ReportPrint.AllReports) Then
    For i = EMPLOYEE_LETTER_CODE.ELC_FIRST_ITEM To EMPLOYEE_LETTER_CODE.ELC_LAST_ITEM
      If (EmployeeLetterCodeIsSubReport(i) And (i <> ELC_P46CAR)) Then
        sCode = EmployeeLetterCode(i, ELCT_LETTER_FILE_CODES, False, Nothing)
        If InStr(1, sEmployeeLetterText, sCode, vbTextCompare) > 0 Then
          Call Err.Raise(ERR_REPORT_INVALID, "ReportTextToRep", "Employee letters can only contain P46 car information with the version of the software." & vbCrLf & "Please remove the tag " & sCode)
        End If
      End If
    Next
  End If
  
  Do While p0 > 0
    'ok so which code do we have
    For i = EMPLOYEE_LETTER_CODE.ELC_FIRST_ITEM To EMPLOYEE_LETTER_CODE.ELC_LAST_ITEM
      sCode = EmployeeLetterCode(i, ELCT_LETTER_FILE_CODES, False, Nothing)
      If InStr(p0, sEmployeeLetterText, sCode, vbTextCompare) = p0 Then
        Exit For
      End If
    Next
    If i = EMPLOYEE_LETTER_CODE.ELC_LAST_ITEM + 1 Then Call Err.Raise(ERR_INVALID_EMPLOYEE_LETTER_CODE, "Report_EmployeeLetter", "Invalid report code found at character " & p0)
    
    If bIsEmail Then
      Call qs.AppendFrom(sEmployeeLetterText, p1, p0 - p1)
    Else
      
      rep.Out (Mid$(sEmployeeLetterText, p1, p0 - p1))
    End If
    
    p1 = p0 + Len(sCode) - 1
    
    bSubReport = EmployeeLetterCodeIsSubReportAttachment(i)
    
    If bSubReport Then
      Call SetMargin(rep, False, bIsEmail)
    End If
    
    If bSubReport And bIsEmail Then
      Call rep.InitReport("", PREPARE_REPORT, RO)
       
      rep.Out (EmployeeLetterCode(i, ELCT_REPORT_CODE, bIsEmail, ee, rep))
      Call rep.EndReport(True)
      If rep.ReportPages > 0 Then
        sExportFile = p11d32.TempPath & EmployeeLetterCode(i, ELCT_FILE_EXPORT, bIsEmail) & rep.ExportTypeExt(Me.ExportOption)
        Call rep.ExportReport(sExportFile, Me.ExportOption, True)
        m_ATCMail.AddAttachment sExportFile 'RK Email 19/03/03
        'Set atm = atms.Add(atms.Count, sExportFile, MAIL_ATT_DATA) 'RK Email 19/03/03
        'atm.DisplayName = EmployeeLetterCode(i, ELCT_CAPTION, bIsEmail, ee, rep) & ".htm" 'RK Email 19/03/03
      End If
    ElseIf bIsEmail Then
      qs.Append (EmployeeLetterCode(i, ELCT_REPORT_CODE, bIsEmail, ee, rep))
    Else
      If IsFontLetterCode(i) Then
        sFont = EmployeeLetterCode(i, ELCT_REPORT_CODE, bIsEmail, ee, rep)
      Else

        Call rep.Out(EmployeeLetterCode(i, ELCT_REPORT_CODE, bIsEmail, ee, rep))
      End If
      
      Call rep.Out(sFont)
    End If
    
    If bSubReport Then Call SetMargin(rep, True, bIsEmail)
    
    p1 = p1 + 1
    If p1 > Len(sEmployeeLetterText) Then Exit Do
    p0 = InStr(p1, sEmployeeLetterText, "{", vbBinaryCompare)
    If p0 = 0 Then
      If p1 < Len(sEmployeeLetterText) Then
        If bIsEmail Then
          Call qs.AppendFrom(sEmployeeLetterText, p1)
        Else
          
          Call rep.Out(Mid$(sEmployeeLetterText, p1))
        End If
      End If
    End If
  Loop
  
  Call SetMargin(rep, False, bIsEmail)
  
  If bIsEmail Then
    Call email(qs.bstr, ee) 'RK Email 19/03/03
    'Call Email(qs.bstr, ee, atms) 'RK Email 19/03/03
  Else
    If Me.Destination = REPD_FILE_PRN Then
      sPRNPathAndFilePerDriver = Me.ExportDirectory & Me.PRNFileName & Me.PRNFileExtension
      Call xKill(sPRNPathAndFilePerDriver)
    End If
    Call rep.EndReport
    If Me.Destination = REPD_FILE_PRN Then
      Call WaitForPRN(sReportName, ee, sPRNPathAndFilePerDriver, rep)
    End If
  End If
  
  
  
ReportTextToRep_END:
  Exit Sub
ReportTextToRep_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportTextToRep", "Report Text To Rep", "Error in ReportTextToRep.")
  Resume ReportTextToRep_END
  Resume
End Sub
'Public Function UseNotes(Optional ByVal bRaiseErrors As Boolean = False) As Boolean 'RK Email 19/03/03
'  Dim s As String
'
'  On Error GoTo UseNotes_ERR
'
'  If UseNotesForEmail Then
'    If Not FileExists(NotesIniFile) Then Call Err.Raise(ERR_FILE_EXISTS, ErrorSource(Err, "UseNotes"), "Unable to find the notes.ini file - " & NotesIniFile)
'  End If
'
'  Mail.UseNotesAPI = UseNotesForEmail
'
'  UseNotes = UseNotesForEmail
'
'UseNotes_END:
'  Exit Function
'UseNotes_ERR:
'  s = "Please check your 'notes' settings on the Print/Options 'Other' tab, "
'  If bRaiseErrors Then
'    Call Err.Raise(Err.Number, ErrorSource(Err, "UseNotes"), s & Err.Description)
'  Else
'    Call ErrorMessage(ERR_ERROR, Err, "UseNotes", "UseNotes", s)
'  End If
'End Function
Private Function ParseForNetscapeClient(sParsedEmployeeLetterText As String) As String
  Dim qs As QString
  
  On Error GoTo ParseForNetscapeClient_ERR
  
  Set qs = New QString
  
  qs.Append ("<HTML>")
  qs.Append (HTMLEncode(sParsedEmployeeLetterText))
  qs.Append ("</HTML>")
  ParseForNetscapeClient = qs
ParseForNetscapeClient_END:
  Exit Function
ParseForNetscapeClient_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ParseForNetscapeClient", "Parse For Netscape Client", "Error in ParseForNetscapeClient")
  Resume ParseForNetscapeClient_END
End Function

Public Function ExportFileName(sReportName As String, ByVal ee As Employee, ByVal rep As Reporter, Optional ByVal bIncludeExportPath As Boolean = True) As String
  Dim sPathAndFile As String
  Dim sFileExt As String
  Dim sExportPath As String
  Dim sFileNameAndExtension As String
  On Error GoTo ExportFileName_ERR
  
  'If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ExportFileName", "The employee is nothing.")
  
  If Me.Destination = REPD_FILE_PRN Then
    sFileExt = Me.PRNFileExtension
  Else
    sFileExt = rep.ExportTypeExt(Me.ExportOption)
  End If
  
  If bIncludeExportPath Then sExportPath = Me.ExportDirectory
  
  If Not ee Is Nothing Then
    sPathAndFile = ExportFileNameEmployeeReference(ee)
  End If
  
  sFileNameAndExtension = sReportName & sPathAndFile & sFileExt
  sFileNameAndExtension = lows.FileNameSafe(sFileNameAndExtension)
  
  sPathAndFile = sExportPath & sFileNameAndExtension
  
  
ExportFileName_END:
  ExportFileName = sPathAndFile
  Exit Function
ExportFileName_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ExportFileName", "Export File Name", "Error getting the export file name")
  Resume ExportFileName_END
  Resume
End Function
Private Function ExportFileNameEmployeeReference(ByVal ee As Employee) As String
  Dim s As String
  Dim sSurname As String
  Dim sFirstName As String
  Dim ben As IBenefitClass
  Dim parts As StringList
  Dim part As String
  Dim i As Long
  
  Select Case p11d32.ReportPrint.ExportEmployeeFileNameFields
    Case REPORT_EXPORT_EMPLOYEE_FILE_NAME_FIELDS.REEFNF_PERSONNEL_NUMBER
      s = ee.PersonnelNumber
    Case REPORT_EXPORT_EMPLOYEE_FILE_NAME_FIELDS.REEFNF_SURNAME_FIRTNAME_PERSONNEL_NUMBER
      Set ben = ee
      Set parts = New StringList
      s = ""
      parts.Add (ben.value(ee_Surname_db))
      parts.Add (ben.value(ee_Firstname_db))
      parts.Add (ee.PersonnelNumber)
      For i = 1 To parts.Count
        part = parts.Item(i)
        part = Trim$(part)
        If (Len(part) > 0) Then
          If (Len(s) > 0) Then
            s = s & "-"
          End If
          s = s & part
        End If
      Next
    Case Else
       Call Err.Raise(ERR_BEN_IS_NOTHING, "ExportFileNameEmployeeReference", "Invalid value for ExportFileNameEmployeeReference.")
  End Select
  
  If (Me.ExportNoSpaces) Then
    s = Replace(s, " ", "")
  End If
  
  s = "-" & s
  
  ExportFileNameEmployeeReference = s
End Function
Public Sub ExportReport(ByVal rep As Reporter, ByVal ee As Employee, sReportName As String)
  Dim sPathAndFile As String
  
  On Error GoTo ExportReport_ERR
  
  If rep.ReportPages > 0 Then
    sPathAndFile = Me.ExportFileName(sReportName, ee, rep)
  
    Call rep.ExportReport(sPathAndFile, Me.ExportOption, True)
  End If
  
ExportReport_END:
  Exit Sub
ExportReport_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ExportReport", "Export Report", "Error exporting report.")
  Resume ExportReport_END
  Resume
End Sub
Public Property Get ParseForNetscape() As Boolean
  ParseForNetscape = (EmailTextAsHTML) And (Not UseNotesForEmail)
End Property
Private Sub email(sParsedEmployeeLetterText As String, ByVal benEE As IBenefitClass) 'RK Email 19/03/03
'Private Sub Email(sParsedEmployeeLetterText As String, ByVal benEE As IBenefitClass, atms As MailAttachments) 'RK Email 19/03/03
  Dim vTolist(0 To 0)
  
  On Error GoTo Email_ERR
  
  vTolist(0) = benEE.value(ee_Email_db)
  If ParseForNetscape Then sParsedEmployeeLetterText = ParseForNetscapeClient(sParsedEmployeeLetterText)
  
  Call m_ATCMail.MailSend(Me.EmailSubject, sParsedEmployeeLetterText, vTolist, Nothing, Nothing, Nothing) 'RK Email 19/03/03 Attachments already added
  
  'Call Mail.MAILSendEx(Me.EmailSubject, sParsedEmployeeLetterText, vToList, Nothing, atms) 'RK Email 19/03/03
  Call m_ATCMail.ClearAllMessageItems
Email_END:
  Exit Sub
Email_ERR:
  If Not benEE Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "Email", "Email", "Error in email for employee " & benEE.value(ee_PersonnelNumber_db))
  Else
    Call ErrorMessage(ERR_ERROR, Err, "Email", "Email", "Error in email")
  End If
  Resume Email_END
  Resume
End Sub
Public Sub ToolBarPreview(CurrentForm As Form)
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  
  On Error GoTo ToolBarPreview_ERR
  Call xSet("ToolBarPreview")

  If IsBenefitForm(CurrentForm) And AllReports Then
    Set ibf = CurrentForm
    If ibf.benclass = BC_NONSHAREDVANS_G Then
      If ibf.benefit Is Nothing Then
        If p11d32.CurrentEmployer.CurrentEmployee.AnyVanBenefit Then
          Set ben = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
        End If
      Else
        Set ben = ibf.benefit
      End If
    Else
      Set ben = ibf.benefit
    End If
    If Not p11d32.ReportPrint.PreviewBenefitWK(ben) Then Call p11d32.ReportPrint.InitPrintDialog
  Else
    Call p11d32.ReportPrint.InitPrintDialog
  End If
  
ToolBarPreview_END:
  Call xReturn("ToolBarPreview")
  Exit Sub
ToolBarPreview_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ToolBarPreview", "Tool Bar Preview", "Error in calling the preview.")
  Resume ToolBarPreview_END
  Resume
End Sub

Public Function PreviewBenefitWK(ben As IBenefitClass) As Long
  Dim rep As Reporter
  
  On Error GoTo PreviewBenefitWK_ERR
  Call xSet("PreviewBenefitWK")
  
  If ben Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "PreviewBenefitWK", "The benefit is nothing.")
  
  Set rep = ReporterNew()
  Call rep.InitReport("P11D report:" & p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_FORM_CAPTION), PREPARE_REPORT, p11d32.ReportPrint.Orientation(RPT_HMIT))
  Call ben.Calculate
  
  Call ben.PrintWk(rep)
  
  Call rep.EndReport
  Call rep.PreviewReport
  
  
  PreviewBenefitWK = True
  
PreviewBenefitWK_END:
  Call xReturn("PreviewBenefitWK")
  Exit Function
PreviewBenefitWK_ERR:
  PreviewBenefitWK = False
  Call ErrorMessage(ERR_ERROR, Err, "PreviewBenefitWK", "Preview Benefit WK", "Error previewing a benefit's working paper.")
  Resume PreviewBenefitWK_END
  Resume
End Function

Private Sub Class_Initialize()
  EnableEmailReports = True
  Call Initialise
End Sub

Private Function CheckMReportsTemplatePathAndFile() As Boolean
  'Checks location and presence of UReports, MReports  and Template.rep (Report Wizard template)
  On Error GoTo CheckMReportsTemplatePathAndFile_Err
  Call xSet("CheckMReportsTemplatePathAndFile")
  Dim TemplatePathAndFile_MReports As String
  Dim TemplatePathAndFile_UReports As String
  Dim lComparison As Long
    
  TemplatePathAndFile_MReports = FullPath(ReportPathManagement) & "Template" & S_REPORT_FILE_EXTENSION
  TemplatePathAndFile_UReports = FullPath(ReportPathUser) & "Template" & S_REPORT_FILE_EXTENSION
  If Not FileExists(ReportPathUser, True) Then
    Call CreateDefaultDirectory(ReportPathUser, "User reports", False)
    Call FileSystem.FileCopy(TemplatePathAndFile_MReports, TemplatePathAndFile_UReports)
  ElseIf Not FileExists(TemplatePathAndFile_UReports) Then
    Call FileSystem.FileCopy(TemplatePathAndFile_MReports, TemplatePathAndFile_UReports)
  Else
    'RK test for Length differences not timestamp as machine/server settings may be different
    If StrComp(FileSystem.FileLen(TemplatePathAndFile_MReports), FileSystem.FileLen(TemplatePathAndFile_UReports), vbTextCompare) Then
      If MsgBox("Your template file differs from the one provided with the application." & vbCrLf & _
                "Do you wish to revert to the system template?", vbYesNo, "Check Template Version") = vbYes Then
         Call FileSystem.FileCopy(TemplatePathAndFile_MReports, TemplatePathAndFile_UReports)
      End If
    End If
  End If
  
  CheckMReportsTemplatePathAndFile = True

CheckMReportsTemplatePathAndFile_End:
  Call xReturn("CheckMReportsTemplatePathAndFile")
  Exit Function

CheckMReportsTemplatePathAndFile_Err:
  CheckMReportsTemplatePathAndFile = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckMReportsTemplatePathAndFile", "Error in CheckMReportsTemplatePathAndFile", "Undefined error.")
  Resume CheckMReportsTemplatePathAndFile_End
  Resume
End Function

Public Property Get ExportDirectory() As String
  ExportDirectory = m_ExportDirectory
End Property

Public Property Let ExportDirectory(sNewValue As String)
  m_ExportDirectory = sNewValue
End Property

Private Sub Class_Terminate()
  Set m_ATCMail = Nothing
End Sub
Public Sub InitATCMAIL()
  If m_ATCMail Is Nothing Then
    Set m_ATCMail = New Mail
    Set m_ATCMail.OwnerForm = MDIMain
    Let m_ATCMail.AutoDetect = True
  End If
End Sub
Public Sub DestroyATCMAIL()
  If Not m_ATCMail Is Nothing Then
    'If (Not m_ATCMail.RedeptionUsing) Then
      Set m_ATCMail = Nothing
    'End If
  End If
End Sub

Public Property Get ExportOption() As REPORT_EXPORTS
  ExportOption = m_ExportOption
End Property

Public Property Let ExportOption(NewVal As REPORT_EXPORTS)
  'RK additional default handling aa this property gets 'lost'
  m_ExportOption = NewVal
End Property


Public Property Get EmployeeLetterFontSize() As Long
  EmployeeLetterFontSize = m_EmployeeLetterFontSize
End Property

Public Property Let EmployeeLetterFontSize(NewVal As Long)
  'RK additional default handling aa this property gets 'lost'
  If Len(NewVal) = 0 Then
    m_EmployeeLetterFontSize = 11
  Else
    m_EmployeeLetterFontSize = NewVal
  End If
End Property

Public Property Get BenClassLastUDM() As Long
  If (p11d32.ReportPrint.AbacusUDM) Then
    BenClassLastUDM = BC_UDM_ABACUS_BENEFITS_LAST_ITEM
  Else
    BenClassLastUDM = BC_UDM_BENEFITS_LAST_ITEM
  End If
  
End Property

Public Function GetLetterFiles(sLetterFiles() As String) As Long
  GetLetterFiles = GetFiles(sLetterFiles, p11d32.ReportPrint.SystemLettersPath, "*" & S_EMPLOYEE_LETTER_FILE_EXTENSION)
End Function

Public Function GetUserLetterFiles(sLetterFiles() As String) As Long
  GetUserLetterFiles = GetFiles(sLetterFiles, p11d32.ReportPrint.UserLettersPath, "*" & S_EMPLOYEE_LETTER_FILE_EXTENSION)   'EK separation in user and system dirs
End Function

Public Function EmployeeLetterPathAndFile() As String
  EmployeeLetterPathAndFile = EmployeeLetterPath & EmployeeLetterFile
End Function
Public Function IsMasterLetterFile(sPathAndFile As String) As Boolean
  On Error GoTo IsMasterLetterFile_ERR
  
  Call xSet("IsMasterLetterFile")
  
  If StrComp(sPathAndFile, p11d32.WorkingDirectory & S_SYSTEMDIR_LETTERS & p11d32.LetterFile & S_EMPLOYEE_LETTER_FILE_EXTENSION, vbTextCompare) = 0 Then
    IsMasterLetterFile = True
  End If
  
IsMasterLetterFile_END:
  Call xReturn("IsMasterLetterFile")
  Exit Function
IsMasterLetterFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IsMasterLetterFile", "Is Master File", "Error determining if the file " & sPathAndFile & " is a master letter file.")
  Resume IsMasterLetterFile_END
End Function

Public Property Get DraftReports() As Boolean
  Dim bDraftReports As Boolean
  
  bDraftReports = p11d32.JuneRelease
  If Not IsRunningInIDE Then
    bDraftReports = bDraftReports Or DraftReportsp46
  End If
  DraftReports = bDraftReports
End Property
Public Property Get DraftReportsp46() As Boolean
   If Not IsRunningInIDE Then
    DraftReportsp46 = (p11d32.LicenceType = LT_DEMO) Or (p11d32.LicenceType = LT_UNLICENSED)
  End If
End Property

'used to put uer reports to screen
Private Sub IEnumFiles_File(ByVal vData As Variant, ByVal sPathAndFile As Variant, ByVal sFIle As String)
  Dim sParent As String
  Dim sName As String
  Dim i As Long
  Dim tvwReports As TreeView
  
  On Error GoTo IEnumFiles_File_ERR
  
  Set tvwReports = vData
  
  If StrComp(p11d32.ReportPrint.TemplatePathAndFile, sPathAndFile, vbTextCompare) = 0 Then GoTo IEnumFiles_File_END
  i = Len(S_REPORT_FILE_EXTENSION)
  sParent = p11d32.ReportPrint.ParentName(RPT_USER)
  sName = Left$(sFIle, Len(sFIle) - i)
  Call ReportToTree(tvwReports, L_REPORT_USER_TAG, sName)
  
IEnumFiles_File_END:
  Exit Sub
IEnumFiles_File_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IEnumFiles", "I Enum Files", "Error processing a user report file, file = " & sPathAndFile & ".")
  Resume IEnumFiles_File_END
End Sub
Public Sub EmailSettingsShow()

 On Error GoTo err_err
  
  Call InitATCMAIL
  Call ATCMail.ShowOptions
  Call DestroyATCMAIL

err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "EmailSettingsShow", "EmailSettingsShow", "Failed to show the email settings")
  Resume err_end


End Sub
