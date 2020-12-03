Attribute VB_Name = "Constants"
Option Explicit


Public Const S_FIELD_LOAN_CHEAP_TAXABLE As String = "Taxable Cheap Loan"
Public Const S_FIELD_CAR_FUEL_WITHDRAWN_DATE As String = "Fuel withdrawn date"
Public Const S_FIELD_AVAILABLE_FROM As String = "Available from"
Public Const S_FIELD_PERSONEL_NUMBER As String = "P_NUM"
Public Const S_FIELD_CAR_REGISTRATION As String = "Reg"

Public Const S_STRING_PROPERTY_OPEN As String = "{{("
Public Const S_STRING_PROPERTY_CLOSE As String = "(}}"

Public Enum NI_VALID
  INVALID = 0
  STANDARD = 1
  TWO_NUMBER = 2
End Enum
Public Const L_NI_NUMBER_TEXT_BOX_INDEX As Long = 2

Public Const S_UNKNOWN As String = "UNKNOWN"

Public Const S_GROUP_CODE1 As String = "Group 1"
Public Const S_GROUP_CODE2 As String = "Group 2"
Public Const S_GROUP_CODE3 As String = "Group 3"

Public Const S_VERSION_CHECK_UNKNOWN  As String = "UNKNOWN"

'****************** HELP ***************
Public Const S_DEFAULT_HELP_PAGE As String = "P11D_Introduction.htm"
Public Const HH_DISPLAY_TOPIC As Long = &H0
Public Const HH_DISPLAY_TEXT_POPUP As Long = &HE
Public Const HH_HELP_CONTEXT As Long = &HF
Public Const HH_CLOSE_ALL As Long = &H12
Public Const HH_SYNC As Long = &H9

Public Const S_URL_ABACUS_WEB_SITE As String = "http://www.abacus.thomsonreuters.com/"
Public Const S_URL_DOWNLOADS As String = S_URL_ABACUS_WEB_SITE & "downloads.htm"
Public Const S_URL_AUTOMATIC_UPDATES As String = S_URL_ABACUS_WEB_SITE & "AutomaticUpdates/default.asp"

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
    ByVal pszFile As String, _
    ByVal uCommand As Long, _
    ByVal dwData As Long) As Long

Public Declare Function HtmlHelpString Lib "hhctrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
    ByVal MergeFile As String, _
    ByVal uCommand As Long, _
    ByVal SubFileAndPage As String) As Long


Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'*********** END HELP *******************
Public Const VK_CONTROL As Long = &H11
Public Const VK_SHIFT  As Long = &H10

Public Const S_FILE_EXTENSION_BAK As String = ".bak"

Public Const S_EMPLOYER_DIRECTORY_LABEL As String = "Employer files in - "

Public Const L_MAX_LONG As Long = 2147483647
Public Const L_EMPLOYEES_OPTION_BASE As Long = 2
Public Const WM_KEYDOWN As Long = &H100



Public Const WM_COMMAND As Long = &H111

Public Const L_CH_FORWARD_SLASH As Long = 47
Public Const L_CH_SPACE As Long = 32
Public Const L_CH_HYPHON As Long = 45

Public Const STRYR2000 = "dd/mm/yyyy" ' cd/apf remove check

Public Const S_TELEPHONE As String = "ndurrant@deloitte.co.uk "
'Public Const S_OLD_CONTACT As String = "For help, please contact abatec on (020) 7438 3669" 'RK Hack to convert telephone number written out by Core
'Public Const S_NEW_CONTACT As String = "For help, please contact abatec on (020) 7303 8122"

Public Const S_MAKE_KEY As String = "KEY_"

Public Const L_MIRAS_DEFAULT As Long = 30000

Public Const S_PASSWORD_OVERIDE As String = "UKCENTRAL"
Public Const S_CDB_EMPLOYEE_NUMBER_PREFIX As String = "CDB_"
Public Const S_CDB_EMPLOYEE_NUMBER_PREFIX_LIKE As String = "CDB_*"


Public Const S_DID_LOAN_COMMENCE_ON_FIRST_DAY_OF_TAX_YEAR_IMPORT_HEADER As String = "Commences on first day of tax year?"

Public Const S_NO_COMPANYDEFINED_CATEGORY As String = "No category"

Public Const S_EMPLOYEE_LETTER_FILE_EXTENSION As String = ".elm"


Public Const S_EMPLOYEE_LETTER_BACKUP_FILE_EXTENSION As String = ".elb"
Public Const S_REPORT_FILE_EXTENSION As String = ".rep"
Public Const S_DB_FILE_EXTENSION As String = ".mdb"

'EK added list of constants to list the output dirs available
Public Const S_USERDIR_EXPORT As String = "Export\"
Public Const S_USERDIR_MMEDIA As String = "MMedia\"
Public Const S_USERDIR_UREPORTS As String = "UReports\"
Public Const S_USERDIR_ULETTERS As String = "ULetters\" 'EK NEW FOR 2003
Public Const S_USERDIR_IMPORT_TRACKING As String = "ImpTrack\" 'EK NEW FOR 2004
Public Const S_USERDIR_PAYEONLINE As String = "PAYEOnline\" 'RK New for 2004
Public Const S_USERDIR_INTRANET As String = "Intranet\"
Public Const S_SYSTEMDIR_HELP As String = "Help\"
Public Const S_SYSTEMDIR_LETTERS As String = "Letters\"
Public Const S_SYSTEMDIR_MREPORTS As String = "MReports\"

Public Const S_ERROR As String = "Error"
Public Const S_FPCS_DEFAULT As String = "Scheme 1"
Public Const S_IRFPCS As String = "Inland Revenue"
Public Const L_IRFPCS As Long = 1&
Public Const L_REPORT_USER_TAG As Long = -1

Public Const S_P1  As String = "Panel1"
Public Const S_P2  As String = "Panel2"

Public Const S_O_EXPENSES_CAPTION = " - Expenses"

Public Const L_NO_LAST_LBINDEX As Long = -1
Public Const L_VANS_BENINDEX As Long = 1
Public Const L_NAME_COL As Long = 28
Public Const L_REFERENCE_COL As Long = 17
Public Const L_NINUMBER_COL As Long = 15
Public Const L_STATUS_COL As Long = 10
Public Const L_GROUP1_COL As Long = 10
Public Const L_GROUP2_COL As Long = 10
Public Const L_GROUP3_COL As Long = 10

Public Const L_CAR_ULTRA_LOW_CO2 As Long = 75

Public Enum LV_EE_ITEMS
  LV_EE_NAME = 0
  LV_EE_PERSONNEL_NUMBER
  LV_EE_NI_NUMBER
  LV_EE_STATUS
  LV_EE_GROUP1
  LV_EE_GROUP2
  LV_EE_GROUP3
End Enum


Public Const L_LV_COL_INDEX_EMPLOYEE_REFERENCE As Long = LV_EE_PERSONNEL_NUMBER + 1
Public Const L_LV_COL_INDEX_EMPLOYEE_GROUP1 As Long = LV_EE_GROUP1 + 1
Public Const L_LV_COL_INDEX_EMPLOYEE_GROUP2 As Long = LV_EE_GROUP2 + 1
Public Const L_LV_COL_INDEX_EMPLOYEE_GROUP3 As Long = LV_EE_GROUP3 + 1


Public Const L_RELOCEXEMPT  As Long = 8000
Public Const L_LOANDEMINIMUS As Long = 10000


Public Const S_COMPANY_CAR_CHECKER_WARNING As String = "WARNING: The following process will change your company car data." & vbCrLf & _
                                                        vbCrLf & _
                                                        "It will attempt to correctly set the replacement and replaced flags " & vbCrLf & _
                                                        "for all cars and appropriately amend available to dates." & vbCrLf & _
                                                        "It will also ensure registration dates are not inconsitent" & vbCrLf & _
                                                        "with the available from date."

'Public Const S_CHK_GRID_CAR_REG As String = "Reg Number"
'Public Const S_CHK_GRID_CAR_AVAIL_FROM As String = "Available From"
'Public Const S_CHK_GRID_CAR_AVAIL_TO As String = "Available To"
'Public Const S_CHK_GRID_CAR_REG_DATE As String = "Reg Date"
Public Const S_DATA_CHECKER_WIZARD_NAME As String = "Data Checker Wizard"
'Public Const S_CHK_GRID_EE_JOINED As String = "Joined"
'Public Const S_CHK_GRID_EE_LEFT As String = "Left"
'Public Const S_CHK_GRID_EE_NI As String = "NI Number"
'Public Const S_CHK_GRID_EE_PNUM As String = "Personnel Number"
'Public Const S_CHK_GRID_EE_FNAME As String = "Firstname"
'Public Const S_CHK_GRID_EE_SNAME As String = "Surname"



Public Const S_LOANALONG As String = "A - For trade, profession"
Public Const S_LOANBLONG As String = "B - Home loan"
Public Const S_LOANCLONG As String = "C - Eligible for interest"
Public Const S_LOANDLONG As String = "D - Not within A,B or C"
Public Const S_LOANELONG As String = "E - Part of relocation"
Public Const S_LOANFLONG As String = "F - Do not know"

Public Const S_LOANASHORT As String = "A"
Public Const S_LOANBSHORT As String = "B"
Public Const S_LOANCSHORT As String = "C"
Public Const S_LOANDSHORT As String = "D"
Public Const S_LOANESHORT As String = "E"
Public Const S_LOANFSHORT As String = "F"
Public Const S_LOANSTERLING As String = "Sterling"
Public Const S_LOANFRANC As String = "Swiss Franc"
Public Const S_LOANYEN As String = "Yen"

Public Const S_STAFF As String = "Staff"
Public Const S_DIRECTOR As String = "Director"

Public Const S_GENDER_NA As String = "N/A"
Public Const S_GENDER_MALE As String = "M"
Public Const S_GENDER_FEMALE As String = "F"

Public Const S_UNDER2499 As String = "0 to 2499"
Public Const S_2500TO17999 As String = "2500 to 17999"
Public Const S_18000PLUS As String = "18000 plus"

Public Const S_ACTUAL As String = "Actual" 'so
Public Const S_ANNUALLY As String = "Annually"
Public Const S_QUARTERLY As String = "Quarterly"
Public Const S_MONTHLY As String = "Monthly"
Public Const S_WEEKLY As String = "Weekly"

Public Const S_COMPANY_CAR_DAYS_UNAVAILABLE_FUEL_DESCRIPTION As String = "Do the car's days unavailable also relate to fuel?"

Public Const S_COMPANY_CAR_CHECKS As String = "Company Car Checks"
Public Const S_EMPLOYEE_CHECKS As String = "Employee Checks"
Public Const S_ONLINE_CHECKS As String = "Online Submission Checks"







Public Const S_VOUCHERS_AND_CREDITCARDS_C As String = "Credit cards and vouchers"
Public Const S_SUBSCRIPTION_N As String = "Subscriptions"
Public Const S_Other_M_Class1a As String = "Other items - Class 1A"
Public Const S_Other_M_NonClass1a As String = "Other items - Non Class 1A"

Public Const S_HOME_LOANS_TYPE As String = "[HOME LOANS]"
Public Const S_BENEFICIAL_LOANS_TYPE As String = "[NORMAL LOAN]"

Public Const S_VAN_FUEL_AVAILABLE_DESCRIPTION = "Was fuel available?"

Public Const S_PRIVATE_MEDICAL_I As String = "Medical"
Public Const S_EDUCATION_N_OTHER As String = "Education"
Public Const S_NURSERY_N_OTHER As String = "Nursery"

Public Const S_PAYMENTS_ON_BEFALF_B As String = "Payments on behalf"
Public Const S_TAX_NOTIONAL_PAYMENTS_B As String = "Notional payments"
Public Const S_SHARES_M As String = "Shares"
Public Const S_INCOME_TAX_PAID_NOT_DEDUCTED_N As String = "Tax paid"
Public Const S_CHAUFFEUR_N_OTHER As String = "Chauffeur"

Public Const S_ENTERTAINMENT_N As String = "Entertainment"
Public Const S_TRAVEL_N As String = "Travel and subsistence"
Public Const S_GENERAL_EXPENSES_N As String = "General"
Public Const S_N As String = "N Other Expense"
Public Const S_HOME As String = "HOME"
Public Const S_ASSETSTRANSFERRED_UDBCODE As String = "A"
Public Const S_ASSETSATDISPOSAL_UDBCODE As String = "L"
Public Const S_SERVICESPROVIDED_UDBCODE As String = "K"
Public Const S_NOSAVE As String = "Amendments to this benefit will not be saved."
Public Const S_NOEMPSAVE As String = "Amendments to this employee will not be saved."

Public Const S_EDIT_CDB_FIELD_db As String = "EditingCompanyDefinedBenefits"

Public Const S_SHARES_MRECIEVED As String = "Received"
Public Const S_SHAREDVAN_KEY As String = "[SHARED_VAN]"

Public Const S_CCCC_MESSAGE_PREFIX_CHECK As String = "This process will check the consistency of company car data for: "
Public Const S_CCCC_MESSAGE_PREFIX_CHANGE As String = "This process will check the consistency of company car data and make changes to it for: "
Public Const S_CCCC_MESSAGE_SUFFIX As String = "Please use the preview or print button after running the check to view the results."

Public Const NO_PRINTERS_AVAILABLE As String = "No printers available"
Public Const NO_PRINTERS_AVAILABLE_MSG As String = "There are no printers available for printing.  Please check the printer settings on your system."

Public Const L_FIX_SPLIT_NAMES As Long = 48

'f12 menu constants
Public Const S_MNU_F12_DEBUG_SQL As String = "Debug SQL"
Public Const S_MNU_F12_UPDATE_FIX As String = "Update fix level"
Public Const S_MNU_F12_MAGNETIC_MEDIA_ERROR_LOGGING As String = "Magnetic Media Error logging"
Public Const S_MNU_F12_MAGNETIC_MEDIA_USER_DATA_SIZE As String = "Magnetic Media User data size"
Public Const S_MNU_F12_SPLIT_NAMES As String = "Split names"
Public Const S_MNU_F12_DELETE_ALL_CDBS As String = "Delete All CDBS"
Public Const S_MNU_F12_UPDATE_LIST_ITEM As String = "Update list item"
Public Const S_MNU_F12_SHOWEMPLOYERS_FIX_LEVEL As String = "Show fix levels"
'MP RV TTP#320 Public Const S_MNU_F12_SET_ACTUAL_MILES As String = "Set Actual Miles"
'Public Const S_MNU_F12_ENABLE_EMAIL_REPORTS As String = "Enable Email"
Public Const S_MNU_F12_EMAIL_DEBUG As String = "Debug Email"
Public Const S_MNU_F12_DISPLAY_INVALIDFIELDS As String = "Display invalid field information"
Public Const S_MNU_F12_KILL_BENEFITS As String = "Kill Benefits"
Public Const S_MNU_F12_ENTER_SERIAL_NUMBER As String = "Enter Serial Number"

Public Const S_MNU_F12_PAYE_ONLINE_SHOW_EXTRA_SUBMISSION_PROPERTIES_MENU As String = "PAYE Online - Show extra submission properties menu"
Public Const S_MNU_F12_PAYE_REFERENCE_ANY_FORMAT As String = "Enable any employer PAYE reference format"
Public Const S_MNU_F12_DATA_TYPE_LIST_VIEW_SORTING As String = "Datatype ListView Sorting"
Public Const S_MNU_F12_F12_SETTINGS_FORM As String = "F12 settings form"
Public Const S_MNU_F12_SORT_OTHER_ALPHABETICALLY As String = "Sort other type benefits alphabetically on working papers etc"
Public Const S_MNU_F12_REREAD_CONTEXT_SENSITIVE_HELP_LINKS As String = "Reread Context sensitive help links"
Public Const S_MNU_F12_VIEW_PROCEED_BUTTON_IF_ERRORS As String = "PAYE online 'view proceed button if errors'"


Public Const S_MNU_F12_QA_MANAGEMENT_REPORTS As String = "QA - Management Reports"
Public Const S_MNU_F12_VERSION_CHECK As String = "Disable online version check"
'end f12

'UDM benefit field constants
Public Const S_UDM_FROM As String = "From"
Public Const S_UDM_To As String = "To"

Public Const L_MM_FIELDSIZE_EMPLOYEE_NUMBER As Long = 17

'File Handling Constants (JN)
Public Const S_UNTITLED As String = "Untitled" 'JN
Public Const S_DATABASE_FILE_MASK As String = ".mdb" 'JN
Public Const L_FOLDER_INVALID As Long = -1

Public Const S_FIRSTNAME As String = "First Name"
Public Const S_SURNAME As String = "Surname"
Public Const S_TITLE As String = "Title"
Public Const S_PNUM As String = "Personnel Number"
Public Const S_IMP_PNUM_ALIAS As String = "[" & S_PNUM & "]"
Public Const S_IMP_PARENT As String = "Parent" 'POST PROCESSING of inconst of fields parent in all import queries ie cars parent = employee, miles parent = car
Public Const S_IMP_FIELD_CATEGORY_DESCRIPTION As String = "Category Description" 'IR description


'UDM Standard Descriptions
Public Const S_UDM_DESCRIPTION As String = "Description"
Public Const S_UDM_VALUE As String = "Value"
Public Const S_UDM_MADE_GOOD_NET As String = "Made Good"
Public Const S_UDM_MADE_GOOD_GROSS As String = "Made Good Gross"
Public Const S_UDM_BENEFIT As String = "Benefit"
Public Const S_UDM_BENEFIT_TITLE As String = "Benefit title"
Public Const S_UDM_BOX_NUMBER As String = "Box Number"
Public Const S_UDM_NIC_CLASS1A_BENEFIT As String = "Class 1A NIC"
Public Const S_UDM_BENEFITS_POTENTIALLY_SUBJECT_TO_CLASS1A As String = "Benefits potentially subject to Class 1A"
Public Const S_UDM_NIC_BENEFIT_IS_SUBJECT_TO_CLASS1A As String = "Benefit is subject to Class1A NIC"
Public Const S_UDM_NIC_CLASS1A_ABLE As String = "Subject to Class 1A"
Public Const S_UDM_NIC_CLASS1A_ADJUSTMENT As String = "Class 1A Adjustment"
Public Const S_UDM_NIC_AMOUNT_MADEGOOD_TAXDEDUCTED As String = "Is amount tax deducted"
Public Const S_UDM_IR_DESCRIPTION As String = "Category"
Public Const S_UDM_IR_BENEFIT_SUBJECT_TO_CLASS1A = "Class1A benefit"
Public Const S_UDM_OPRA_AMOUNT_FOREGONE = "OpRA amount foregone"
Public Const S_UDM_OPRA_AMOUNT_FOREGONE_HELP = "The OpRA amount foregone is not apportioned for benefit calculations. Please see https://www.gov.uk/government/publications/optional-remuneration-arrangements for further information."
Public Const S_UDM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE = "OpRA amount foregone used as value"
Public Const S_UDM_VALUE_NON_OPRA = "Value non OpRA"
Public Const S_PAYEONLINE_MDB As String = "PayeOnline.mdb"
Public Const S_CURRENCY As String = "£"

'Serial Number constants
Public Const S_SERIAL_NUMBER_STANDARD = "P11DD847B378T4373419" 'Full functionality, all reports, emails etc.  Used for year-end (main) release.
Public Const S_SERIAL_NUMBER_INTRANET = "4687263T4S6P84662L19" 'As above with intranet publishing
Public Const S_SERIAL_NUMBER_SHORT = "P11DD8S4678V37746619" 'Reduced functionality (no printing (except P46, no employye letter etc.).  Used for b/f (minor) release.
Public Const S_SERIAL_NUMBER_DEMO = "P11DD366O659V3774619" 'Full functionality available for both releases

'XML
Public Const D_XMLBENEFITMAX As Double = 9999999.99

Public Enum LICENCE_TYPE
  LT_UNLICENSED = 0
  LT_STANDARD ''Full functionality, all reports, emails etc. Used for year-end (main) release.
  LT_INTRANET 'As above with intranet publishing
  LT_SHORT 'Reduced functionality (no printing (except P46, no employye letter etc.).  Used for b/f (minor) release.
  LT_DEMO 'Full functionality available for both releases
End Enum

Public Enum OUTPUT_TYPE
  OT_SCREEN
  OT_MAGENTIC_MEDIA
  OT_PAYE_ONLINE
End Enum

Public Enum INTRANET_OUTPUT_TYPE
  IOT_P11D
  IOT_P11D_WORKING_PAPERS
  IOT_EMPLOYEE_LETTER
  IOT_WORKING_PAPERS
End Enum

Public Enum INTRANET_LOGIN_USERNAME_SOURCE
  ILUS_USERNAME
  [ILUS_FIRSTITEM] = ILUS_USERNAME
  ILUS_PERSONNEL_NUMBER
  ILUS_EMAIL
  ILUS_FULLNAME
  [ILUS_LASTITEM] = ILUS_FULLNAME
End Enum

Public Enum INTRANET_AUTHENTICATION_TYPE
  IAT_FULL
  IAT_WINDOWS
  IAT_OTHER
End Enum

Public Enum BENEFITS_ENUM_TYPE
  BET_NEED_TO_CALCULATE
End Enum

Public Enum BEN_CLASS
  
  BC_ASSETSTRANSFERRED_A = 1
  [BC_FIRST_ITEM] = BC_ASSETSTRANSFERRED_A '1
  
  BC_PAYMENTS_ON_BEFALF_B '2
  BC_TAX_NOTIONAL_PAYMENTS_B '3
  BC_VOUCHERS_AND_CREDITCARDS_C '4
  BC_LIVING_ACCOMMODATION_D '5
  BC_EMPLOYEE_CAR_E '6
  BC_COMPANY_CARS_F '7
  BC_FUEL_F '8
  BC_NONSHAREDVANS_G '9
  BC_LOAN_OTHER_H '10
  BC_PRIVATE_MEDICAL_I '11
  BC_QUALIFYING_RELOCATION_J '12
  BC_SERVICES_PROVIDED_K '13
  BC_ASSETSATDISPOSAL_L '14
  'BC_SHARES_M '15 'leave to enable roll forward of files
  BC_CLASS_1A_M '16
  BC_NON_CLASS_1A_M '17
  BC_INCOME_TAX_PAID_NOT_DEDUCTED_M '18
  BC_TRAVEL_AND_SUBSISTENCE_N '19
  BC_ENTERTAINMENT_N '20
  BC_GENERAL_EXPENSES_BUSINESS_N '21
  BC_PHONE_HOME_N '22
  BC_NON_QUALIFYING_RELOCATION_N '23
  BC_CHAUFFEUR_OTHERO_N '24
  BC_OOTHER_N '25
  
  [BC_UDM_BENEFITS_LAST_ITEM] = BC_OOTHER_N
  [BC_UDM_ABACUS_BENEFITS_LAST_ITEM] = [BC_UDM_BENEFITS_LAST_ITEM] 'RETAIN JUST INCASE WE HAVE COCKED UP, CAD 01/04/2004
  BC_nonSHAREDVAN_G '26
  BC_SHAREDVAN_G '27
  
  BC_LOANS_H '28
  
  BC_SHAREDVANs_G '9
  
  BC_EMPLOYEE 'last item udm
  BC_EMPLOYER
  
  BC_CDB
  [BC_REAL_BENEFITS_LAST_ITEM] = BC_CDB
  
  
  BC_ALL
  BC_NONSHAREDVANS_FUEL_G
  
  [BC_LAST_ITEM] = BC_CDB
  
  
End Enum

Public Enum REGULAR_PAYMENTS_METHOD
  RPM_MONTHLY
  RPM_INTERVAL
End Enum

Public Enum BRING_FORWARD_TYPE
  BFT_OVERWRITE = 0
  BFT_UPDATE
End Enum

Public Enum NAME_ORDERS
  [NO_FIRST_ITEM]
  NO_FN_INITIALS_SURNAME = [NO_FIRST_ITEM]
  NO_FN_SURNAME
  NO_INITIALS_SURNAME
  NO_SURNAME_TITLE_FN_INITIALS
  NO_SURNAME_INITIALS
  NO_SURNAME_FN_INITIALS
  NO_SURNAME_FN
  NO_TITLE_FN_INITIALS_SURNAME
  NO_TITLE_INITIALS_SURNAME
  [NO_LAST_ITEM] = NO_TITLE_INITIALS_SURNAME
End Enum

Public Enum TREE_IMAGES
  IMG_UNSELECTED = 1
  IMG_SELECTED
  IMG_FOLDER_CLOSED
  IMG_FOLDER_OPEN
  IMG_LETTER_CLOSED
  IMG_LETTER_OPEN
  IMG_SELECTED_STATUS
  IMG_INFO
  IMG_EXCLAMATION
  IMG_OK
  IMG_REPORT
End Enum

Public Enum CHECK_MESSAGE_TYPE
  CMT_LIST_ITEM
  CMT_ALERT_MESSAGE_CHECK
  CMT_ALERT_MESSAGE_CHANGE
  CMT_TREEVIEW_NODE_TITLE
  CMT_ALERT_MESSAGE_DESCRIPTION
  
End Enum

Public Enum COMPANY_CAR_CHECKER_MESSAGE_TYPE
  cccmt_list_item
  CCCMT_ALERT_MESSAGE_CHECK
  CCCMT_ALERT_MESSAGE_CHANGE
  CCCMT_TREEVIEW_NODE_TITLE
End Enum

Public Enum EMPLOYEE_CHECKS
  [_EC_FIRST_ITEM]
  EC_NI_NUMBER = [_EC_FIRST_ITEM]
  EC_FIRSTNAME
  EC_SURNAME
  [_EC_LAST_ITEM] = EC_SURNAME
End Enum

Public Enum CHECK_BEFORE_PRINT
  YES_THIS_TIME_ONLY = 0
  NO_THIS_TIME_ONLY
  YES_ALWAYS
  NEVER
End Enum

Public Enum CHECKORDERBY
  [_ORDER_FIRST_ITEM]
  ORDER_PNUM = [_ORDER_FIRST_ITEM]
  ORDER_SURNAME
  ORDER_FULLNAME
  ORDER_NI
  [_ORDER_LAST_ITEM] = ORDER_NI
End Enum

Public Enum CHECKS
  CK_ALL_CHECKS = -1
  [_CK_FIRST_ITEM] = 0
  CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE = [_CK_FIRST_ITEM]
  CK_CC_OVERLAPS
  CK_CC_NOCARS
  CK_CC_AVAILDATES
  CK_CC_REGDATES
  CK_CC_EE_AVAILDATES
  CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED
  [_CK_EC_START]
  CK_EC_NI = [_CK_EC_START]
  [_CK_EC_END] = CK_EC_NI
  [_CK_LAST_ITEM] = [_CK_EC_END]
End Enum

Public Enum EE_CHECKS
  EE_CK_ALL = -1
  EE_CK_NI = 1
  EE_CK_P_NUM = 2
  EE_CK_NAME = 4
End Enum

Public Enum COMPANY_CAR_CHECKER_CHECKS
  [_CCCC_FIRST_ITEM]
  CCCC_DATES = [_CCCC_FIRST_ITEM]
  CCCC_OVERLAPS
  CCCC_NOCARS
  CCCC_AvailDates
  CCCC_RegDates
  CCCC_EeAvailDates
  [_CCCC_LAST_ITEM] = CCCC_EeAvailDates
End Enum

Public Enum ApplicationErrors
  ERR_Application = TCSCLIENT_ERROR
  ERR_DAYS_INCONSISTENT
  ERR_EMPLOYER_DB
  ERR_INVALID_REPORT_STANDARD_DATA
  ERR_BEN_CLASS_NOT_EQUAL
  ERR_COPY_FAILED
  ERR_MOVE_MENU_UPDATE_EMPLOYEE
  ERR_NODE_IS_NOTHING
  ERR_RATES_TABLE
  ERR_REP_IS_NOTHING
  ERR_NO_RECORDS
  ERR_NOT_BENEFIT_FORM
  ERR_RS_IS_NOTHING
  ERR_TOTALS_NOT_EQUAL_PARAMARRAY
  ERR_RATES_INITIALISE
  ERR_FILE_FAIL_OPEN
  ERR_RUNFIXES
  ERR_INVALID_EMPLOYEE_LETTER_CODE
  ERR_INVALID_QUERY
  ERR_FIX_LEVEL_LOW
  ERR_NO_EMPLOYEES_SELECTED
  ERR_NO_SPLIT_CHOICES
  ERR_CAN_NOT_OPEN_FILE
  ERR_NO_SPLIT_NAMES
  ERR_NO_ID_RECORD
  ERR_NO_FIXLEVEL
  ERR_NOT_RUNFIXES
  ERR_CURENTFORM_NOT_EMPLOYERS
  ERR_SYNCDB
  ERR_FALSE
  ERR_A4
  ERR_BENCLASS_INVALID
  ERR_EMPLOYEE_IS_NOTHING
  ERR_EMPLOYEERELEASE
  ERR_OPENEMPLOYER
  ERR_INIT
  ERR_TEXTSTREAM_IS_NOTHING
  ERR_INVALIDBENCLASS
  ERR_NORECORDS
  ERR_ONLY_ONE_EMPLOYER
  ERR_DETAIL_TO_LISTITEM
  ERR_POST_PROCESS
  ERR_FPCSBAND
  ERR_NOCARSCHEME
  ERR_IRFPCS
  ERR_P11DINIT
  ERR_DOREPORT
  ERR_DIRECTORY_NOT_EXIST
  ERR_REPWIZARD_IS_NOTHING
  ERR_INITIMPORT
  ERR_INVALID_COMPANY_CAR_CHECKER_CHECK
  ERR_ADDIMPORTQUERIES
  ERR_CREATEEMPLOYER
  ERR_NO_EDIT_CDB
  ERR_LV_IS_NOTHING
  ERR_INVALIDCOL_INDEX
  
  ERR_ENUM__VALUE_INVALID
  ERR_FILE_NOT_EXIST
  ERR_COPY_FAIL
  ERR_FILE_OPEN
  ERR_FILE_INVALID
  ERR_BACKUP_FILE
  
  ERR_NO_EMPLOYEE
  ERR_BEN_HAS_NO_UDM_FIELDS
  ERR_NO_FIND_EMPLOYEE
  ERR_FILE_OPEN_EXCLUSIVE
  ERR_ADDSTATICS
  ERR_BEN_IS_NOTHING
  ERR_FIELD_LEN_0
  ERR_EMPLOYER_INVALID
  ERR_NO_FLOPPY
  ERR_INVALID_FIELDS
  ERR_DB_IS_NOTHING
  ERR_WORKSPACE_IS_NOTHING
  ERR_NO_FREE_SPACE
  ERR_PARENT_IS_NOTHING
  ERR_REPORTER_IS_NOTHING
  ERR_INVALID_DATA_FORMAT_SPACE
  
  ERR_TYPE0RECORD
  ERR_TYPE1RECORD
  
  ERR_TYPE20RECORD
  ERR_TYPE2BENRECORD
  ERR_TYPE3RECORD
  ERR_TYPE4RECORD
  ERR_NI_INVALID_BUT_ADDED
  
  ERR_PERSONEL_NUMBER_CHANGED
  ERR_BEN_INCORRECT
  ERR_BEN_NOT_REPORTABLE
  ERR_INVALID_NI_RATIO
  ERR_IS_NOTHING
  ERR_MMFIELD_REQUIRED_NOT_FILLED
  ERR_DB_OPEN_ERROR
  ERR_NO_EMPLOYER
  ERR_APP_NAME
  ERR_FILE_EXISTS
  ERR_STRING_TOO_SHORT
  ERR_NOT_NUMERIC
  ERR_NOT_UPTODATE
  ERR_COMMAND_PARAMS
  ERR_IS_LINK_BEN
  ERR_IS_NOT_CDB
  ERR_INVALID_BENEFIT_INDEX
  ERR_PASSWORD
  ERR_NOT_NOTHING
  ERR_REPAIR_COMPACT
  ERR_DATES
  ERR_NOT_ARRAY
  ERR_USER_REPORT
  ERR_INVALID_FORM
  ERR_CALCULATE_FUEL_BENEFITS
  ERR_INVALID_UDM_RECORD
  ERR_DUPLICATE_NI_NUMBERS
  ERR_REPORT_TEXT_TO_REP
  ERR_APP_YEAR_INVALID
  ERR_FILE_DELETE
  ERR_DIRECTORY_EXISTS
  ERR_ZERO_LENGTH_STRING
  ERR_REPORT_INVALID
  ERR_WAIT_FOR_PRN
  ERR_STRING_TOO_LONG
  ERR_WRITING_DATA
  ERR_DIRECTORY_CREATE
  ERR_SPACE_IN_PAYE_NUM
  ERR_ELEC_BEFORE_1998
  ERR_CAR_LIST_PRICE
  ERR_MARORS
  ERR_ROTARY
  ERR_MULTIPLEFUELWITHDRAWN
  ERR_NEEDENGINESIZE
  ERR_NEEDCO2ORENGINESIZE
  ERR_XML_DOC_INVALID
  ERR_XMLNUMBERTOOBIG  ' EK XML ERRORS
  
  'EK Error for people with accomodation expenses
  ERR_ACCOMODATION_EXPENSES
  
  'IK Paye online errors
  ERR_PAYEFIELD_REQUIRED_NOT_FILLED
  ERR_PAYEFIELD_OPTIONAL_NOT_FILLED
  ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS
  ERR_PAYE_DATE_OUTOFSYNC
  ERR_PAYE_TEXT_FIELD_INVALID
  ERR_PAYE_ONLINE_CANCEL
  ERR_CAR_AMOUNT
  ERR_CAR_MAKE_AND_MODEL
  ERR_CAR_ENGINE_SIZE
  ERR_CAR_CO2
  ERR_PAYE_CHANGED_FIELD
  ERR_PAYE_INVALID_FIELD
  ERR_DISPLAY
  ERR_WORKING_DIRECTORY
  ERR_INVALID
  ERR_BENEFITS
  ERR_PREVIEW
  ERR_PRINT_CANCEL
  ERR_CAR_DATES_INVALID
  ERR_CAR_FUEL_AVAILABLE_TO_INVALID
  ERR_DIESEL_REGISTERED_AFTER_1_1_2006
  ERR_CAR_ELEC_PRE_98
      
  TCSCLIENT_ERROR_END = ERR_CAR_ELEC_PRE_98
End Enum


Public Enum INI_READ_WRITE
  Ini_read
  Ini_Write
End Enum

Public Enum MM_STANDARD_OUT
  MM_SO_VAN_TYPE
  MM_SO_VAN_TYPE_WITH_DESCRIPTION
  MM_SO_ASSETSTRANSFERRED_TYPE_WITH_DESCRIPTION
  MM_SO_ASSETSTRANSFERRED_TYPE
  MM_SO_VALUE_MADEGOOD_BENEFIT
  MM_SO_RECORD_HEADER
  MM_SO_ASSETSTRANSFERRED_TYPE_WITH_DESCRIPTION_NO_RECORD_HEADER
  MM_SO_NONE
  MM_SO_SECTION_B
  MM_SO_IR_DESC_TYPE_BENEFIT_AND_VALUE
  MM_SO_IR_DESC_TYPE_BENEFIT
  MM_SO_SUMMARY_TYPE
End Enum

Public Enum EMPLOYEE_SELECTION
  [_ES_FIRST_ITEM] = 0
  ES_SELECTED = [_ES_FIRST_ITEM]
  ES_INVERSE_SELECTED
  ES_ALL
  [_ES_LAST_ITEM]
  ES_CURRENT = [_ES_LAST_ITEM]
End Enum

Public Enum EDIT_CDBS
  ECDB_Value
  ECDB_SETVALUE
End Enum

Public Enum PW_MODE
  PWM_CHECK_CURRENT
  PWM_SET
End Enum


Public Enum MM_RECORD_TYPE
  MM_REC_FILE_OPEN = 1
  MM_REC_EMPLOYER_OPEN
  MM_REC_EMPLOYEE
  MM_REC_BENEFIT
  MM_REC_BENFIT_SECTION_F
  MM_REC_BENEFIT_SECTION_U
  MM_REC_BENEFIT_SECTION_A  'AM
  MM_REC_BENEFIT_SECTION_B  'AM
  MM_REC_BENEFIT_SECTION_L  'AM
  MM_REC_BENEFIT_SECTION_N  'AM
  MM_REC_EMPLOYER_CLOSE
  MM_REC_FILE_CLOSE
End Enum

Public Enum ITERATION_TYPE
  AnyDirty
  AnyInvalid
  WriteDirtyAndAreAllWritten
  AnyReportable
End Enum

Public Enum EMPLOYEE_LETTER_CODE
  ELC_ADDRESS
  [ELC_FIRST_ITEM] = ELC_ADDRESS
  
  ELC_BOLD
  
  ELC_COMPANY_NAME
  ELC_CONTACT_NAME
  ELC_CONTACT_NUMBER
  
  ELC_DATE_NOW
  ELC_DATE_TAXYEAR
  ELC_DATE_NEXT_TAXYEAR
  
  'ELC_DATE_PT_SUBMISSION_DEADLINE
  'ELC_DATE_PT_REVENUE_CALC_DEADLINE
  ELC_DATE_KEEP_DETAILS_UNTILL
  ELC_DATE_RESPONSE
  
  ELC_DATE_TAX_YEAR_START
  ELC_DATE_TAX_YEAR_END
  
  ELC_EMPLOYEE_NAME_INITIALS
  ELC_EMPLOYEE_NAME_FIRST
  ELC_EMPLOYEE_NAME_FULL
  ELC_EMPLOYEE_NAME_SALUTATION
  ELC_EMPLOYEE_NAME_SURNAME
  ELC_EMPLOYEE_NAME_TITLE
  
  
  ELC_GROUP1
  ELC_GROUP2
  ELC_GROUP3
  
  ELC_NEWPAGE
  ELC_NI_NUMBER
  ELC_NORMAL
  
  ELC_PAYE_REF
  ELC_PERSONNEL_NUMBER
  
  
  ELC_SIGNATORY
  
  [ELC_FIRST_SUB_REPORT]
  ELC_HMIT = [ELC_FIRST_SUB_REPORT]
  ELC_P46CAR
  ELC_TABLE
  ELC_WORKING_PAPERS
  [ELC_LAST_SUB_REPORT] = ELC_WORKING_PAPERS
  [ELC_LAST_ITEM] = ELC_WORKING_PAPERS
  
  
  
End Enum

Public Enum EMPLOYEE_LETTER_CODE_TYPE
  ELCT_LETTER_FILE_CODES
  ELCT_REPORT_CODE
  ELCT_MENU_CAPTION
  ELCT_CAPTION
  ELCT_MENU_PARENT
  ELCT_FILE_EXPORT
End Enum

Public Enum CONTROL_PROPERTY
  CP_ENABLED
  CP_VISIBLE
End Enum

Public Enum EMPLOYER_LV_COLUMNS
  ELVC_EMPLOYER_NAME = 1
  ELVC_PAYE
  ELVC_FILE_NAME
  ELVC_NO_EMPLOYEES
  ELVC_FIX_LEVEL
End Enum

Public Enum COMPANY_CAR_FUEL_TYPE
  CCFT_PETROL = 0
  CCFT_DIESEL
  CCFT_HYBRID
  CCFT_ELECTRIC
  CCFT_BIFUEL_WITH_CO2_FOR_GAS
  CCFT_BIFUELOLD  'redundant, convert in read db to CCFT_BIFUEL_WITH_CO2_FOR_GAS !!!! 'CAD 2003/2004, 'don't delete as will screw existing data
  CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B
  CCFT_EUROIVDIESEL
  CCFT_GAS_ONLY 'new in 20032004 to get rid of the existing crap and confirm to ir calcs ...CAD
  CCFT_E85_BIO_ENTHANOL_AND_PETROL
  CCFT_RDE2_DIESEL
  
End Enum


Public Enum COMPANY_CAR_PAYMENT_FREQUENCY
  CCPF_ANNUALLY = 0
  CCPF_QUARTERLY
  CCPF_MONTHLY
  CCPF_WEEKLY
  CCPF_ACTUAL
End Enum

Public Enum FILE_TYPES
  FIT_SYSTEM_DEFINED = 0
  FIT_USER_DEFINED
  [FIT_LAST_ITEM] = FIT_USER_DEFINED
End Enum

'windows constants
Public Const WM_LBUTTONDOWN As Long = &H201

'RC - Have put version number in const
'Public Const S_VERSION_NUM = "2.3.0"

'EK for IR dropdowns section A.
Public Enum IR_DESC_A
  IRDA_PLEASESELECT = 0
  IRDA_CARS
  IRDA_PROPERTY
  IRDA_PRECIOUSMETALS
  IRDA_OTHER
  ' IRDA_MULTIPLE
End Enum



'EK for IR dropdowns section L.
Public Enum IR_DESC_L
  
  IRDL_PLEASESELECT = 0
  IRDL_HOLIDAYACCOM
  IRDL_TIMESHAREACCOM
  IRDL_AIRCRAFT
  IRDL_BOAT
  IRDL_CORPORATEHOSP
  ' IRDL_MULTIPLE
  IRDL_OTHER
  
End Enum


Public Const S_IR_DESC_OTHER = "Other"
Public Const S_IR_DESC_PLEASE_SELECT = "Please select an item"
' ek 1/04 Items for Asset at Disposal Dropdown

Public Const S_IR_DESC_L_HOLIDAYACCOM As String = "Holiday accommodation"
Public Const S_IR_DESC_L_TIMESHAREACCOM As String = "Timeshare accommodation"
Public Const S_IR_DESC_L_AIRCRAFT As String = "Aircraft"
Public Const S_IR_DESC_L_BOAT As String = "Boat"
Public Const S_IR_DESC_L_CORPORATEHOSP As String = "Corporate hospitality"

Public Const S_PAYE_ONLINE_NAMESPACE_EFILER_GOVTALK As String = "df"
Public Const S_PAYE_ONLINE_NAMESPACE_EFILER_ERROR As String = "err"

'ek 1/04 Items for Asset Transferred Dropdown
Public Const S_IR_DESC_A_CARS As String = "Cars"
Public Const S_IR_DESC_A_PROPERTY As String = "Property"
Public Const S_IR_DESC_A_PRECIOUSMETALS As String = "Precious metals"

Public Const S_IR_DESC_B_DOMESTIC_BILLS As String = "Domestic bills"
Public Const S_IR_DESC_B_ACCOUNTANCY_FEES As String = "Accountancy fees"
Public Const S_IR_DESC_B_PRIVATE_ED As String = "Private education"
Public Const S_IR_DESC_B_PRIVATE_CAR_EX As String = "Private car expenses"
Public Const S_IR_DESC_B_SEASON_TICKET As String = "Season tickets"

'SectionM
  'section m 1a
  Public Const S_IR_DESC_M_C1A_SUBS_AND_FEES As String = "Subscriptions & fees"
  Public Const S_IR_DESC_M_C1A_ED_ASS As String = "Educational assistance CL1A"
  Public Const S_IR_DESC_M_C1A_NON_QUAL_RELOC As String = "Non-qualifying relocation ben"
  Public Const S_IR_DESC_M_C1A_STOP_LOSS_CHARGES As String = "Stop loss charges"
  
  'section m PAYE, output
  Public Const S_IR_DESC_M_C1A_PAYE_ONLINE_SUBS_AND_FEES As String = "subscriptions and fees"
  Public Const S_IR_DESC_M_C1A_PAYE_ONLINE_ED_ASS As String = "educational assistance CL1A"
  
  'section m non 1a
  Public Const S_IR_DESC_M_NC1A_SUBS_AND_FEES As String = "Subs and professional fees"
  Public Const S_IR_DESC_M_NC1A_NURSERY As String = "Nursery places"
  Public Const S_IR_DESC_M_NC1A_ED_ASS As String = "Educational assistance"
  Public Const S_IR_DESC_M_NC1A_LOANS_WRIT_WAIV As String = "Loans written or waived"
  
  'section m non 1a PAYE, output
  Public Const S_IR_DESC_M_NC1A_PAYE_ONLINE_SUBS_AND_FEES As String = "subs and professional fees"

'section n
Public Const S_IR_DESC_N_PERSONAL_INC_EXP As String = "Personal Incidental Expenses"
Public Const S_IR_DESC_N_WORK_HOME As String = "Work Done at Home"

Public Const S_INVALID_FILE_CHARS As String = ":\/?*<>|"

Public Const D_FUEL_TYPE_L_NO_LONGER_HAS_DISCOUNT  As Date = #1/1/2006#

Public Enum IR_DESC_B
  
  IRDB_PLEASESELECT = 0
  IRDB_DOMESTIC_BILLS
  IRDB_ACCOUNTANCY_FEES
  IRDB_PRIVATE_ED
  IRDB_PRIVATE_CAR_EX
  IRDB_SEASON_TICKET
  IRDB_OTHER
  
End Enum

  
Public Enum IR_DESC_N
  
  IRDN_PLEASESELECT = 0
  IRDN_SUBS_AND_FEES
  IRDN_ED_ASS_CL1A
  IRDN_NON_QUAL_RELOC
  IRDN_STOP_LOSS_CHARGES
  IRDN_OTHER
  
End Enum


Public Enum IR_DESC_NON
  
  IRDNON_PLEASESELECT = 0
  IRDNON_SUBS_AND_FEES
  IRDNON_NURSERY
  IRDNON_ED_ASSIS
  IRDNON_LOANS_WRIT_WAIV
  IRDNON_OTHER
  
End Enum


Public Enum PayeOnlneNameSpace
  GoTalk = 0
  Errors
End Enum


Public Enum IR_DESC_O
  
  IRDO_PLEASESELECT = 0
  IRDO_PERSONAL_INC_EXP
  IRDO_WORK_HOME
  irdo_OTHER
  
End Enum

Public Type P46_FUEL_TYPE_STRINGS
  Letter As String
  Description As String
End Type

Public Const S_QUOT As String = """"
Public Const S_QUOTQUOT As String = """"""

Public Const CDATA_START As String = "<![CDATA["
Public Const CDATA_END As String = "]]>"

' Note: Both ReplaceXMLMetacharacters & XMLText are dependent on this value
Public Const S_INVALID_XML_CHARS As String = S_QUOT & "<>&'"

Public Const PAYEONLINE_STATUS_NOT_SUBMITTED As String = "Not submitted"
Public Const PAYEONLINE_STATUS_WITH_HMRC As String = "With HMRC not confirmed"
Public Const PAYEONLINE_STATUS_SUBMITTED As String = "Submitted"

Public Const S_DB_FIELD_OPRA_AMOUNT_FOREGONE As String = "OPRAAmountForegone"

Public Const S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT As String = "CashEquivOrRelevantAmt"
Public Const S_PAYE_ONLINE_COST_OR_AMOUNT_FORGONE As String = "CostOrAmtForgone"
Public Const S_PAYE_ONLINE_GROSS_OR_AMOUNT_FORGONE As String = "GrossOrAmtForgone"
Public Const S_PAYE_ONLINE_TAXABLE_PAYMENT_OR_RELEVANT_AMOUNT As String = "TaxablePmtOrRelevantAmt"

Public Function P46FuelTypeStrings(ccft As COMPANY_CAR_FUEL_TYPE) As P46_FUEL_TYPE_STRINGS
  Dim p46s As P46_FUEL_TYPE_STRINGS
  
    Select Case ccft
      Case CCFT_PETROL
        p46s.Letter = "A"
        p46s.Description = "Petrol"  'RH
      Case CCFT_DIESEL
        p46s.Letter = "D"
        p46s.Description = "Diesel (not RDE2)"  'RH
      Case CCFT_HYBRID
        p46s.Letter = "A"
        p46s.Description = "Hybrid electric" 'RH
      Case CCFT_ELECTRIC
        p46s.Letter = "A"
        p46s.Description = "Zero emmissions (inc electric)" 'RH
      Case CCFT_BIFUEL_WITH_CO2_FOR_GAS
        p46s.Letter = "A"
        p46s.Description = "Bi-fuel" 'RH 'THIS IS BIFUEL
      Case CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B
        p46s.Letter = "A"
        p46s.Description = "Bi-fuel conv/other bi-fuel"
      Case CCFT_EUROIVDIESEL
        p46s.Letter = "D"
        p46s.Description = "Euro IV Diesel" 'AM
      Case CCFT_GAS_ONLY
        p46s.Letter = "A"
        p46s.Description = "Gas Only"
      Case CCFT_E85_BIO_ENTHANOL_AND_PETROL
        p46s.Letter = "A"
        p46s.Description = "E85"
      Case CCFT_RDE2_DIESEL:
        p46s.Letter = "F"
        p46s.Description = "RDE2 (Euro 6d)"
      Case Else
        Call ECASE("Invalid fuel type to convert to fuel type letter")
    End Select
  
    P46FuelTypeStrings = p46s
End Function

