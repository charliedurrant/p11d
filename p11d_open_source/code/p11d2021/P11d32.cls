VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "p11d32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IKeyboardHook
Implements IEnumEmployers


Public P11DWS As Workspace

Public ReportPrint As ReportPrint

Public FindFilesSearchSubDirs As Boolean
Public FindFilesDirList As String
Public FindFilesSavedFilesList As String
Public PayeReferenceAnyFormat As Boolean 'default to false
Public Employers As ObjectList
Public Importing As Importing
Public CurrentEmployer As Employer
Public Rates As Rates
Public MagneticMedia As MagneticMedia
Public PAYEonline As PAYEonline
Public Help As Help

Public Intranet As Intranet
Public BringForward As BringForward
Public EmployeeSortOrder As ListSortOrderConstants
Public EmployeeSortOrderColumn As LV_EE_ITEMS

Public RegularPaymentsMethod As REGULAR_PAYMENTS_METHOD
Public RegularPaymentsForceEndOfMonth As Boolean

Public DisplayInvalidFields As Boolean

Public udm As udm
Public TemplateAppversion As String
Public TemplateFixlevel As Long
Public Username As String

' Public UserEmployeeLetterFile As String ' EK separation of user and application directories
' Public UserEmployeeLetterPath  As String ' EK separation of user and application directories

Public MoveToNextEmployeeWithBenefit As Boolean
Public FixLevelsShow As Boolean
Public KillBenefits As Boolean
Public NameOrder As NAME_ORDERS
Public CompanyCarCheckerCheck As Long

'Inisection constants
Private Const S_INI_SECTION_REPORTS As String = "Reports"
Private Const S_INI_SECTION_EMPLOYEE_LETTER As String = "Employee Letter"
Private Const S_INI_SECTION_MAGNETIC_MEDIA As String = "MagneticMedia"
Private Const S_INI_SECTION_PAYEONLINE As String = "PayeOnline"
Private Const S_INI_SECTION_BRING_FORWARD As String = "BringForward"
Private Const S_INI_SECTION_GENERAL As String = "General"
Private Const S_INI_SECTION_INTRANET As String = "Intranet"

'Private P11D32.IniPathAndFile As String
Private m_AppYear As Long
Private m_AppYearShort As String

Private m_BenData As BEN_DATA
Private m_UDMTableOffsets() As BEN_CLASS
Private m_Contact As String
Private m_LicenceType As LICENCE_TYPE
Private m_SerialNumber As String
Private m_GlobalCopySerialNumber As String
Private m_WorkingDirectory As String

Private m_IsGlobal As Boolean
Private m_IniPathAndFile As String
Private m_IniPathAndFileLastYear As String
Private m_IniPathAndFileLastYearExists As Boolean
Private m_IniPathAndFileGlobal As String
Private m_UserDataDirectory As String
Private m_UserDataDirectoryYearSpecific As String
Public Enum VIEW_ERRORS_TYPE
  VET_IMPORTING = 1
  VET_MAGNETIC_MEDIA
  VET_PAYEONLINE_VALIDATION
  VET_PAYEONLINE_SUBMISSION
End Enum
Public DataTypeListViewSorting As Boolean
Public SortEmployeeReferenceAsNumber As Boolean
Public ValidateNINumberOnEmployeeScreen As Boolean
Public VersionCheckEnabled As Boolean

Private m_versionCheck As VersionCheck

Private m_versionCheckResult As String
Private m_versionCheckFinished As Boolean
Implements IDebugMenu

Private Function IDebugMenu_DatabaseMenuItem(ByVal MenuName As String) As Boolean
  IDebugMenu_DatabaseMenuItem = False
End Function

Private Function IDebugMenu_RepairCompactAvailable() As Boolean
 IDebugMenu_RepairCompactAvailable = True
End Function

Private Function IDebugMenu_SQLAvailable() As Boolean
  IDebugMenu_SQLAvailable = False
End Function
Public Sub RepairAndCompactGeneral()
  Dim sPathAndFile As String
  
On Error GoTo err_end
  
  sPathAndFile = FileOpenDlg("Please select a database file to repair and compact", "All Files (*.*)|*.*", p11d32.WorkingDirectory)
  If Len(sPathAndFile) = 0 Then
    GoTo err_end
  End If
  If Not RepairCompactDB(sPathAndFile, MODE_REPAIR + MODE_COMPACT) Then Call Err.Raise(ERR_REPAIR_COMPACT, "RepairAndCompact", "Failed to repair and compact " & sPathAndFile)
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "RepairAndCompactGeneral", "Repair And Compact", Err.Description)
  Resume err_end:
End Sub


Public Property Get UserDataDirectory() As String
  If (Len(m_UserDataDirectory) = 0) Then
    m_UserDataDirectory = FullPath(GetSpecialFolderEx(L_USER_APP_DATA)) & "Applications\" & app.companyName & "\P11D\"  '"\\uklon0558\London Abatec\Projects\Dump\" '
  End If
  UserDataDirectory = m_UserDataDirectory
End Property
Public Property Get UserDataDirectoryYearSpecific() As String
  If (Len(m_UserDataDirectoryYearSpecific) = 0) Then
    m_UserDataDirectoryYearSpecific = FullPath(GetSpecialFolderEx(L_USER_APP_DATA)) & "Applications\" & app.companyName & "\P11D" & p11d32.AppYear & "\"   '"\\uklon0558\London Abatec\Projects\Dump\" '
  End If
  UserDataDirectoryYearSpecific = m_UserDataDirectoryYearSpecific
End Property
Public Property Get LoanHomeToBeneficialBringFwd() As Boolean
  LoanHomeToBeneficialBringFwd = (Me.AppYear = 2003) And p11d32.BringForward.Yes
End Property
Public Property Get SystemPath()
  SystemPath = FullPath(FullPath(AppPath) & "System")
End Property
Public Property Get SystemPAYEOnlinePath()
  SystemPAYEOnlinePath = FullPath(FullPath(FullPath(AppPath) & "System") & "PAYEOnline")
End Property
Public Property Get SystemPAYEOnlineEXBXSDPath()
  SystemPAYEOnlineEXBXSDPath = SystemPAYEOnlinePath & "EXB.xsd"
End Property
Public Property Get SystemPAYEOnlineEXBPriorXSDPath() 'only used after march release for p46 before the year end!
  SystemPAYEOnlineEXBPriorXSDPath = SystemPAYEOnlinePath & "EXB_PRIOR_YEAR.xsd"
End Property
Public Property Get SystemIntranetPath()
  SystemIntranetPath = FullPath(FullPath(FullPath(AppPath) & "System") & "Intranet")
End Property

'Public Property Get Contact()
'  Dim retval As String
'  'RK change to reset telephone number (core library restriction)
'  If StrComp(GetStatic("Contact"), S_OLD_CONTACT, vbTextCompare) = 0 Then
'    Call AddStatic("Contact", S_NEW_CONTACT, S_NEW_CONTACT, False)
'  End If
'  Contact = GetStatic("Contact")
'End Property
Public Property Get TempPath()
  Dim sBuffer As String
  Dim retval As String
  
  sBuffer = String$(TCSBUFSIZ, 0)
  retval = GetTempPath(TCSBUFSIZ, sBuffer)
  If retval = 0 Then Exit Property
  TempPath = FullPath(Left$(sBuffer, retval))
End Property
Public Function MoveMenuUpdateEmployeeCheckEmployer() As Boolean
  If Not CurrentEmployer Is Nothing Then
    MoveMenuUpdateEmployeeCheckEmployer = p11d32.CurrentEmployer.MoveMenuUpdateEmployee
  End If
End Function

Public Sub BenDataLinksCreate(ey As Employer)
  Dim benEE As IBenefitClass, ee As Employee
  Static bDoneOnce As Boolean
  
  On Error GoTo BenDataLinksCreate_ERR
  
  Call xSet("BenDataLinksCreate")
  
  If ey Is Nothing Then Call Err.Raise(ERR_EMPLOYER_INVALID, "BenDataLinksCreate", "Employer is nothing.")
  
  If Not bDoneOnce Then
    Set benEE = New Employee
    Set benEE.Parent = ey
    Set ee = benEE
    Call ee.LoadBenefits(TBL_ALLBENEFITS, False)
    bDoneOnce = True
  End If
  
BenDataLinksCreate_END:
  Call xReturn("BenDataLinksCreate")
  Exit Sub
BenDataLinksCreate_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BenDataLinksCreate", "Ben Data Links Create", "Error creating ben data links.")
  Resume BenDataLinksCreate_END
End Sub

Public Function PDDB(ws As Workspace) As Database
  Set PDDB = OpenDB(ws, PDDBPath, False)
End Function
Public Property Get PDDBPath()
  PDDBPath = FullPath(AppPath) & "P11D.PD"
End Property

Public Function ViewErrors(VET As VIEW_ERRORS_TYPE) As Boolean
  Dim rs As Recordset
  Dim emp As Employer
  Dim ben As IBenefitClass
  
  Dim sqli As New SQLQUERIES_IMP
  
  On Error GoTo ViewErrors_Err
  Call xSet("ViewErrors")
  
  'CAD mwthods need consolidation importing needs to be like Magnetic media
  'both with maybe IErrorReport
  Select Case VET
    Case VET_IMPORTING
      If Not F_Employers.LB.SelectedItem Is Nothing Then
        Set emp = p11d32.Employers(F_Employers.LB.SelectedItem.Tag)
        Set ben = emp
        Call emp.OpenEmployer(False)
        Set p11d32.Importing.Employer = emp
        Call StartAutoSTD(sqli.Queries(IMP_SELECT_ERRORS), emp.db, p11d32.Importing.ReportHeader(emp), PREPARE_REPORT)
        
        Set p11d32.Importing.Employer = Nothing
      End If
    Case VET_MAGNETIC_MEDIA
      Call p11d32.MagneticMedia.Errors(PREPARE_REPORT)
    Case VET_PAYEONLINE_VALIDATION
      Call p11d32.PAYEonline.Errors(PREPARE_REPORT, VET_PAYEONLINE_VALIDATION)
    Case Else
      Call ECASE("Invalid View Errors Type.")
  End Select
  
ViewErrors_End:
  Set rs = Nothing
  If Not ben Is Nothing Then ben.Kill
  Call xReturn("ViewErrors")
  Exit Function

ViewErrors_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ViewErrors", "View Errors", "Unable to view errors.")
  Resume ViewErrors_End
End Function

Private Function InitBenDataLinks()
  Dim i As Long
  
  On Error GoTo InitBenDataLinks_ERR

  Call xSet("InitBenDataLinks")
  
  For i = [BC_FIRST_ITEM] To [BC_REAL_BENEFITS_LAST_ITEM]
    ReDim m_BenData.DataLinks(i).StaticData(1 To Rates.BenClassTo(i, BCT_BENITEMS_LAST_ITEM))
    m_BenData.DataLinks(i).Initialised = False
  Next
  
InitBenDataLinks_END:
  Call xReturn("InitBenDataLinks")
  Exit Function
InitBenDataLinks_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "InitBenDataLinks", "Init Ben Data Links", "Error initialising the benefit data links.")
  Resume InitBenDataLinks_END
  Resume
End Function
Public Property Get DataLinkInitialised(bc As BEN_CLASS) As Boolean
  DataLinkInitialised = m_BenData.DataLinks(bc).Initialised
End Property
Public Property Let DataLinkInitialised(bc As BEN_CLASS, NewValue As Boolean)
  m_BenData.DataLinks(bc).Initialised = NewValue
End Property
Public Function BenClassIsReal(ByVal bc As BEN_CLASS) As Boolean
  BenClassIsReal = bc >= [BC_FIRST_ITEM] And bc <= [BC_REAL_BENEFITS_LAST_ITEM]
  If Not BenClassIsReal Then
    ECASE ("benclass is not real in BenClassIsReal") 'CAD err
  End If
End Function

Private Function BenClassIsUDM(ByVal bc As BEN_CLASS)
  BenClassIsUDM = bc >= [BC_FIRST_ITEM] And bc <= p11d32.ReportPrint.BenClassLastUDM Or (bc = BC_EMPLOYER Or bc = BC_EMPLOYEE)
  If Not BenClassIsUDM Then
    ECASE ("benclass is not UDM in BenClassIsReal") 'CAD err
  End If
End Function
Public Sub BenDataLinkAddUDMTableOffset(ByVal bc As BEN_CLASS)
  Dim i As Long
  Dim b As Boolean
  'check if added aready
  For i = 1 To m_BenData.UDMTableOffsetUBound
    If m_BenData.UDMTableOffsets(i) = bc Then
      Exit Sub
    End If
  Next
  m_BenData.UDMTableOffsetUBound = m_BenData.UDMTableOffsetUBound + 1
  ReDim Preserve m_BenData.UDMTableOffsets(1 To m_BenData.UDMTableOffsetUBound)
  m_BenData.UDMTableOffsets(m_BenData.UDMTableOffsetUBound) = bc
End Sub
Public Property Get BenDataLinkUDMRequiresCalculate(ByVal bc As BEN_CLASS, ByVal Item As Long) As Boolean
  If BenClassIsReal(bc) Then BenDataLinkUDMRequiresCalculate = m_BenData.DataLinks(bc).StaticData(Item).RequiresCalculate
End Property
Public Property Let BenDataLinkUDMRequiresCalculate(ByVal bc As BEN_CLASS, ByVal Item As Long, NewValue As Boolean)
  If BenClassIsReal(bc) Then m_BenData.DataLinks(bc).StaticData(Item).RequiresCalculate = NewValue
End Property
Public Property Get BenDataLinkBenClassFromTableOffet(ByVal lTableOffset As Long) As BEN_CLASS
    BenDataLinkBenClassFromTableOffet = m_BenData.UDMTableOffsets(lTableOffset)
End Property
Public Property Get BenDataLinkBenfitTable(ByVal bc As BEN_CLASS) As BENEFIT_TABLES
  If BenClassIsReal(bc) Then BenDataLinkBenfitTable = m_BenData.DataLinks(bc).BenefitTable
End Property
Public Property Let BenDataLinkBenfitTable(ByVal bc As BEN_CLASS, ByVal NewValue As BENEFIT_TABLES)
  If BenClassIsReal(bc) Then m_BenData.DataLinks(bc).BenefitTable = NewValue
End Property

Public Property Get BenDataLinkDataType(ByVal bc As BEN_CLASS, ByVal Item As Long) As DATABASE_FIELD_TYPES
  If BenClassIsReal(bc) Then BenDataLinkDataType = m_BenData.DataLinks(bc).StaticData(Item).DataType
End Property
Public Property Let BenDataLinkDataType(ByVal bc As BEN_CLASS, ByVal Item As Long, ByVal NewValue As DATABASE_FIELD_TYPES)
  If BenClassIsReal(bc) Then m_BenData.DataLinks(bc).StaticData(Item).DataType = NewValue
End Property
Public Property Get BenDataLinkMMRequired(ByVal bc As BEN_CLASS, ByVal Item As Long) As Boolean
  If BenClassIsReal(bc) Then BenDataLinkMMRequired = m_BenData.DataLinks(bc).StaticData(Item).MMRequired
End Property
Public Property Let BenDataLinkMMRequired(ByVal bc As BEN_CLASS, ByVal Item As Long, ByVal NewValue As Boolean)
  If BenClassIsReal(bc) Then m_BenData.DataLinks(bc).StaticData(Item).MMRequired = NewValue
End Property
Public Property Get BenDataLinkMMFieldSize(ByVal bc As BEN_CLASS, ByVal Item As Long) As Long
  If BenClassIsReal(bc) Then BenDataLinkMMFieldSize = m_BenData.DataLinks(bc).StaticData(Item).MMFieldSize
End Property
Public Property Let BenDataLinkMMFieldSize(ByVal bc As BEN_CLASS, ByVal Item As Long, NewValue As Long)
  If BenClassIsReal(bc) Then m_BenData.DataLinks(bc).StaticData(Item).MMFieldSize = NewValue
End Property
Public Property Get BenDataLinkUDMDescription(ByVal bc As BEN_CLASS, ByVal Item As Long) As String
  Dim i As Long
  Dim IFD As FieldDetails
  i = BenDataLinkFieldIndex(bc, Item)
  Set IFD = m_BenData.DataLinks(bc).UDMDescriptions(i)
  BenDataLinkUDMDescription = IFD.Name
End Property
'this maps to description which is either the description but if blank the name
Public Property Get BenDataLinkUDMDisplayName(ByVal bc As BEN_CLASS, ByVal Item As Long) As String
  Dim i As Long
  Dim IFD As FieldDetails
  i = BenDataLinkFieldIndex(bc, Item)
  Set IFD = m_BenData.DataLinks(bc).UDMDescriptions(i)
  If (Len(IFD.Description)) = 0 Then
    BenDataLinkUDMDisplayName = IFD.Name
  Else
    BenDataLinkUDMDisplayName = IFD.Description
  End If
End Property
Public Property Let BenDataLinkUDMDisplayName(ByVal bc As BEN_CLASS, ByVal Item As Long, NewValue As String)
  Dim i As Long
  Dim IFD As FieldDetails
  i = BenDataLinkFieldIndex(bc, Item)
   
  IFD = m_BenData.DataLinks(bc).UDMDescriptions(i)
  IFD.Description = NewValue
   
End Property

'convert the ITEM as in ITEM_BENEFIT to the udm index
Private Function BenDataLinkFieldIndex(ByVal bc As BEN_CLASS, ByVal Item As Long) As Long
  Dim i As Long
    
  For i = 1 To m_BenData.DataLinks(bc).UDMUBound
    If m_BenData.DataLinks(bc).UDMFieldIDs(i) = Item Then
      BenDataLinkFieldIndex = i
      Exit Function
    End If
  Next
   
  Call Err.Raise(ERR_INVALID_UDM_RECORD, "BenDataLinkFieldIndex", "Could not map benefit item to UDM index")
  
End Function

'this is really FieldDetails.Name
Public Property Let BenDataLinkUDMDescription(ByVal bc As BEN_CLASS, ByVal Item As Long, NewValue As String)
  Dim IFD As FieldDetails
  
  If BenClassIsUDM(bc) Then
    Set IFD = New FieldDetails
    m_BenData.DataLinks(bc).UDMUBound = m_BenData.DataLinks(bc).UDMUBound + 1
    ReDim Preserve m_BenData.DataLinks(bc).UDMDescriptions(1 To m_BenData.DataLinks(bc).UDMUBound)
    IFD.Name = NewValue
    If m_BenData.DataLinks(bc).StaticData(Item).DataType = TYPE_NOSUM Then ECASE ("Data type not set when setting UDM description, benclass = " & bc & ", item = " & Item & ".")
    
    IFD.DataType = m_BenData.DataLinks(bc).StaticData(Item).DataType
    Set m_BenData.DataLinks(bc).UDMDescriptions(m_BenData.DataLinks(bc).UDMUBound) = IFD
    
    ReDim Preserve m_BenData.DataLinks(bc).UDMFieldIDs(1 To m_BenData.DataLinks(bc).UDMUBound)
    m_BenData.DataLinks(bc).UDMFieldIDs(m_BenData.DataLinks(bc).UDMUBound) = Item
    
    ReDim Preserve m_BenData.DataLinks(bc).UDMFields(1 To m_BenData.DataLinks(bc).UDMUBound)
    m_BenData.DataLinks(bc).UDMFields(m_BenData.DataLinks(bc).UDMUBound) = NewValue
  End If
  
  
End Property
'return and let error occur, this can only be used afer BenDataLinkUDMDescription is called
'as that inits the item
Public Property Get BenDataLinkFieldDetails(ByVal bc As BEN_CLASS, ByVal Item As Long) As FieldDetails
  Dim i As Long
  i = BenDataLinkFieldIndex(bc, Item)
  Set BenDataLinkFieldDetails = m_BenData.DataLinks(bc).UDMDescriptions(i)
End Property
Public Property Get BenDataLinkUDMFieldID(bc As BEN_CLASS, lIndex As Long)
  BenDataLinkUDMFieldID = m_BenData.DataLinks(bc).UDMFieldIDs(lIndex)
End Property
Public Property Get BenDataLinkUDMUBound(ByVal bc As BEN_CLASS) As Long
  If BenClassIsReal(bc) Then BenDataLinkUDMUBound = m_BenData.DataLinks(bc).UDMUBound
End Property
Public Property Get BenDataLinkUDMFields(ByVal bc As BEN_CLASS) As Variant
  If BenClassIsUDM(bc) Then BenDataLinkUDMFields = m_BenData.DataLinks(bc).UDMFields()
End Property

Public Property Get BenDataLinkUDMFieldDetails(ByVal bc As BEN_CLASS) As Variant
  If BenClassIsUDM(bc) Then BenDataLinkUDMFieldDetails = m_BenData.DataLinks(bc).UDMDescriptions()
End Property
Public Property Get BenDataLinkMMFieldRequiredNotFilled(ByVal bc As BEN_CLASS, ByVal Item As Long, ByVal value As Variant) As Boolean
  On Error GoTo BenDataLinkMMFieldRequiredNotFilled_ERR
  
  Call xSet("BenDataLinkMMFieldRequiredNotFilled")
  
  If BenClassIsReal(bc) Then
    If m_BenData.DataLinks(bc).StaticData(Item).MMRequired Then
      Select Case m_BenData.DataLinks(bc).StaticData(Item).DataType
        Case TYPE_STR
          value = CStr(value)
          If Len(value) > 0 Then BenDataLinkMMFieldRequiredNotFilled = True
        Case TYPE_DATE
          If IsDate(value) Then BenDataLinkMMFieldRequiredNotFilled = True
        Case Else
          If IsNumeric(value) Then BenDataLinkMMFieldRequiredNotFilled = True
      End Select
    Else
      BenDataLinkMMFieldRequiredNotFilled = True
    End If
  End If
  
BenDataLinkMMFieldRequiredNotFilled_END:
  Call xReturn("BenDataLinkMMFieldRequiredNotFilled")
  Exit Function
BenDataLinkMMFieldRequiredNotFilled_ERR:
  Resume BenDataLinkMMFieldRequiredNotFilled_END
End Property

Private Function GetIniEntryBool(section As String, Key As String, DEFAULT As Boolean) As Boolean
  Dim s As String
  s = GetIniEntry(section, Key, DEFAULT, p11d32.IniPathAndFile)
  If (StrComp(s, "True", vbTextCompare) = 0) Then
    GetIniEntryBool = True
  End If
End Function
Private Function DateTaxYearFix(ByVal bFrom, ByVal d As Date) As Date
  If bFrom Then
    If (d < p11d32.Rates.value(TaxYearStart) Or d > p11d32.Rates.value(TaxYearEnd)) Then
       d = p11d32.Rates.value(TaxYearStart)
    End If
  Else
    If (d > p11d32.Rates.value(TaxYearStart) Or d < p11d32.Rates.value(TaxYearEnd)) Then
       d = p11d32.Rates.value(TaxYearEnd)
    End If
  End If
  DateTaxYearFix = d
  
End Function
Private Function IniSettingLastYear(ByVal section As String, ByVal Key As String, ByVal DEFAULT As String, ByVal bForceLastYear As Boolean) As String
  IniSettingLastYear = DEFAULT
  If IniPathAndFileLastYearExists And ((p11d32.AppYear >= 2005) Or bForceLastYear) Then
    IniSettingLastYear = GetIniEntry(section, Key, DEFAULT, IniPathAndFileLastYear)
  End If
End Function
Public Sub JuneReleaseWarning()
  Dim bWarning As Boolean
  Dim sWarning As String
  
  If (Not p11d32.JuneRelease) Then
    Exit Sub
  End If
  
  If (IniJuneReleaseWarning(Ini_read)) Then
    sWarning = "This version of the software is only intended for:" & vbCrLf & vbCrLf & _
               "1. The importing of data into the system" & vbCrLf & _
               "2. The production of P46(Car) returns via paper and online submissions" & vbCrLf & vbCrLf & _
               "If you require further information please call our support line at " & S_TELEPHONE

    Call MsgBox(sWarning, vbInformation, "Interim release")
    IniJuneReleaseWarning (Ini_Write)
  End If
  
End Sub
Private Function IniJuneReleaseWarning(IRW As INI_READ_WRITE) As Boolean
  Dim sKey As String
  
  sKey = "JuneReleaseWarning-" & Trim(GetVersionString(False))
  If IRW = Ini_Write Then
    Call WriteIniEntry(S_INI_SECTION_GENERAL, sKey, False, p11d32.IniPathAndFile) 'EK for installsheild change to write working dir to ini file
  Else
    IniJuneReleaseWarning = GetIniEntryBool(S_INI_SECTION_GENERAL, sKey, True)
  End If
End Function
Public Function IniSettings(IRW As INI_READ_WRITE) As Boolean
  Dim i As Long
  Dim dQuarterEnd As Date, dQuarterStart As Date
  Dim s As String
  On Error GoTo IniSettings_ERR
  
  Call xSet("IniSettings")
  
  If IRW = Ini_Write Then
    If Not p11d32.NoSerial Then
      Call WriteIniEntry(S_INI_SECTION_GENERAL, "SerialNumber", p11d32.SerialNumber, p11d32.IniPathAndFile)
      Call WriteIniEntry(S_INI_SECTION_GENERAL, "GlobalCopySerialNumber", p11d32.GlobalCopySerialNumber, p11d32.IniPathAndFile)
    End If
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "WorkingDirectory", p11d32.WorkingDirectory, p11d32.IniPathAndFile) 'EK for installsheild change to write working dir to ini file
    
    Call WriteIniEntry(S_INI_SECTION_EMPLOYEE_LETTER, "EmployeeLetterFile", p11d32.ReportPrint.EmployeeLetterFile, p11d32.IniPathAndFile) 'this is the last file selecte
    Call WriteIniEntry(S_INI_SECTION_EMPLOYEE_LETTER, "EmployeeLetterPath", p11d32.ReportPrint.EmployeeLetterPath, p11d32.IniPathAndFile) ' this is the last path selected
    Call WriteIniEntry(S_INI_SECTION_EMPLOYEE_LETTER, "UserLettersDirectory", p11d32.ReportPrint.UserLettersPath, p11d32.IniPathAndFile)
    
    For i = [RPT_FIRST_INI] To [RPT_LAST_INI]
      Call WriteIniEntry(S_INI_SECTION_REPORTS, ReportPrint.Name(i), ReportPrint.Orientation(i), p11d32.IniPathAndFile)
    Next
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "DefaultReportIndex", ReportPrint.DefaultReportIndex, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "DefaultSelectEmployeeReportIndex", ReportPrint.DefaultSelectEmployeeReportIndex, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "UserReportFileLessExtension", ReportPrint.UserReportFileLessExtension, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "UserReportSelectEmployeeFileLessExtension", ReportPrint.UserReportSelectEmployeeFileLessExtension, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "EmployeeSelection", ReportPrint.EmployeeSelection, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "HMITSelectionChoice", ReportPrint.HMITSelectionChoice, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "HMITSections", ReportPrint.HMITSections, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P46DateFrom", ReportPrint.P46DateFrom, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P46DateTo", DateValReadToScreen(ReportPrint.P46DateTo), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P46Range", DateValReadToScreen(ReportPrint.P46Range), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P46PrintReplacedP46s", p11d32.ReportPrint.P46PrintReplacedP46s, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P46UserDateFrom", DateStringEx(ReportPrint.P46UserDateFrom, p11d32.Rates.value(TaxYearStart)), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P46UserDateTo", DateStringEx(ReportPrint.P46UserDateTo, p11d32.Rates.value(TaxYearEnd)), p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "TimeStamp", p11d32.ReportPrint.TimeStamp, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "PrintedEmployees", p11d32.ReportPrint.PrintedEmployees, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "EmployeeLetterMargin", p11d32.ReportPrint.EmployeeLetterMargin, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "EmployeeLetterFontName", p11d32.ReportPrint.EmployeeLetterFontName, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "EmployeeLetterFontSize", p11d32.ReportPrint.EmployeeLetterFontSize, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "P11DTotalValue", p11d32.ReportPrint.P11DTotalValue, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "CheckOptions", p11d32.ReportPrint.CheckOptions, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "CheckOrderBy", p11d32.ReportPrint.CHECKORDERBY, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "ChecksAutoRefresh", p11d32.ReportPrint.ChecksAutoRefresh, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "A4ForcePrint", p11d32.ReportPrint.A4ForcePrint, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "ChecksSelected", p11d32.ReportPrint.ChecksSelected, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "Destination", p11d32.ReportPrint.Destination, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "ExportDirectory", p11d32.ReportPrint.ExportDirectory, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "PRNFileName", p11d32.ReportPrint.PRNFileName, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "PRNFileExtension", p11d32.ReportPrint.PRNFileExtension, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "EmailTextAsHTML", p11d32.ReportPrint.EmailTextAsHTML, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "DatesOnWorkingPaper", p11d32.ReportPrint.DatesOnWorkingPaper, p11d32.IniPathAndFile)

    Call WriteIniEntry(S_INI_SECTION_REPORTS, "ExportOption", p11d32.ReportPrint.ExportOption, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "ExportAllowUserNameHTML", p11d32.ReportPrint.ExportAllowUserNameHTML, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "ExportEmployeeFileNameFields", p11d32.ReportPrint.ExportEmployeeFileNameFields, p11d32.IniPathAndFile)
        
        
        
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "EmailSubject", p11d32.ReportPrint.EmailSubject, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "SeparatePrintJobs", p11d32.ReportPrint.SeparatePrintJobs, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "NotesIniFile", p11d32.ReportPrint.NotesIniFile, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "UseNotesForEmail", p11d32.ReportPrint.UseNotesForEmail, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "HMITFieldTrim", p11d32.ReportPrint.HMITFieldTrim, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "PrnLoopLength", p11d32.ReportPrint.PRNLoopLength, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_REPORTS, "RemeberEmployeeSelection", p11d32.ReportPrint.RemeberEmployeeSelection, p11d32.IniPathAndFile)
   
    Call WriteIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "OutPutDirectory", FullPath(p11d32.MagneticMedia.OutputDirectory), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "DataFormat", p11d32.MagneticMedia.DataFormat, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "LastPathAndFileCreated", p11d32.MagneticMedia.LastPathAndFileCreated, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "RecordViewID", p11d32.MagneticMedia.RecordViewID, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "LastPathAndFileCreated", p11d32.PAYEonline.LastPathAndFileCreated, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "OutPutDirectory", FullPath(p11d32.PAYEonline.OutputDirectory), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "P46DateFrom", DateValReadToScreen(p11d32.PAYEonline.P46DateFrom), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "P46DateTo", DateValReadToScreen(p11d32.PAYEonline.P46DateTo), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "P46Range", p11d32.PAYEonline.P46Range, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "P46UserDateFrom", DateStringEx(p11d32.PAYEonline.P46UserDateFrom, p11d32.Rates.value(TaxYearStart)), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "P46UserDateTo", DateStringEx(p11d32.PAYEonline.P46UserDateTo, p11d32.Rates.value(TaxYearEnd)), p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "Efiler_test_submission", p11d32.PAYEonline.Efiler_Test_Submission, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "Efiler_mock_submission", p11d32.PAYEonline.Efiler_Mock_Submission, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "Efiler_proceed_submission", p11d32.PAYEonline.Efiler_Proceed_Submission, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "Type", p11d32.PAYEonline.Type_Selected, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "ViewProceedButtonIfErrors", p11d32.PAYEonline.ViewProceedButtonIfErrors, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "ByPassLocalSchemaCheck", p11d32.PAYEonline.ByPassLocalSchemaCheck, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_PAYEONLINE, "DisableCheckSubmissionWithAbatec", p11d32.PAYEonline.DisableCheckSubmissionWithAbatec, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_BRING_FORWARD, "BringForwardType", p11d32.BringForward.BringForwardType, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_BRING_FORWARD, "HMITSChosen", p11d32.BringForward.HMITSChosen, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_BRING_FORWARD, "CDCs", p11d32.BringForward.CDCs, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_BRING_FORWARD, "CDBs", p11d32.BringForward.CDBs, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_BRING_FORWARD, "comments", p11d32.BringForward.Comments, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_BRING_FORWARD, "SharedVans", p11d32.BringForward.SharedVans, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "NameOrder", p11d32.NameOrder, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "EmployeeSortOrder", p11d32.EmployeeSortOrder, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "EmployeeSortOrderColumn", p11d32.EmployeeSortOrderColumn, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "FindFilesSearchSubDirs", p11d32.FindFilesSearchSubDirs, p11d32.IniPathAndFile) 'JN
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "FindFilesDirList", p11d32.FindFilesDirList, p11d32.IniPathAndFile)  'SO
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "FindFilesSavedFilesList", p11d32.FindFilesSavedFilesList, p11d32.IniPathAndFile) 'SO

    Call WriteIniEntry(S_INI_SECTION_GENERAL, "RegularPaymentsMethod", p11d32.RegularPaymentsMethod, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "RegularPaymentsForceEndOfMonth", p11d32.RegularPaymentsForceEndOfMonth, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "KillBenefits", p11d32.KillBenefits, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "CompanyCarCheckerCheck", p11d32.CompanyCarCheckerCheck, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "PayeReferenceAnyFormat", p11d32.PayeReferenceAnyFormat, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "DataTypeListViewSorting", p11d32.DataTypeListViewSorting, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "SortEmployeeReferenceAsNumber", p11d32.DataTypeListViewSorting, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "MoveToNextEmployeeWithBenefit", p11d32.MoveToNextEmployeeWithBenefit, p11d32.IniPathAndFile)
    
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "ImportTracking", p11d32.Importing.Tracking, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "ImportTrackingDirectory", FullPath(p11d32.Importing.TrackingPath), p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "ValidateNINumberOnEmployeeScreen", p11d32.ValidateNINumberOnEmployeeScreen, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_GENERAL, "VersionCheck", p11d32.VersionCheckEnabled, p11d32.IniPathAndFile)
    
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "OutputDirectory", FullPath(p11d32.Intranet.OutputDirectory), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "LoginUserNameSource", p11d32.Intranet.LoginUserNameSource, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "AuthenticationType", p11d32.Intranet.AuthenticationType, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "OutputType", p11d32.Intranet.OutputType, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "BannerTitle", Replace(p11d32.Intranet.BannerTitle, "=", ""), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "UserInfoHTML", Replace(p11d32.Intranet.UserInfoHTML, "=", ""), p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "CaseSensitiveOnFullAuthentication", p11d32.Intranet.CaseSensitiveOnFullAuthentication, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "BannerBackColor", p11d32.Intranet.BannerBackColor, p11d32.IniPathAndFile)
    Call WriteIniEntry(S_INI_SECTION_INTRANET, "BannerForeColor", p11d32.Intranet.BannerForeColor, p11d32.IniPathAndFile)

    Call WriteIniEntry(S_INI_SECTION_REPORTS, "SortOtherTypeBenefitsAlphabetically", p11d32.ReportPrint.SortOtherTypeBenefitsAlphabetically, p11d32.IniPathAndFile)
 Else
    p11d32.GlobalCopySerialNumber = GetIniEntry(S_INI_SECTION_GENERAL, "GlobalCopySerialNumber", "", p11d32.IniPathAndFile)
    p11d32.SerialNumber = GetIniEntry(S_INI_SECTION_GENERAL, "SerialNumber", "", p11d32.IniPathAndFile)
    If (Len(p11d32.SerialNumber) = 0) Then
      p11d32.SerialNumber = GetIniEntry(S_INI_SECTION_GENERAL, "SerialNumber", "", p11d32.IniPathAndFileGlobal)
      p11d32.GlobalCopySerialNumber = p11d32.SerialNumber
    Else
      If StrComp(p11d32.GlobalCopySerialNumber, GetIniEntry(S_INI_SECTION_GENERAL, "SerialNumber", "", p11d32.IniPathAndFileGlobal), vbTextCompare) <> 0 Then
        p11d32.SerialNumber = GetIniEntry(S_INI_SECTION_GENERAL, "SerialNumber", "", p11d32.IniPathAndFileGlobal)
        p11d32.GlobalCopySerialNumber = GetIniEntry(S_INI_SECTION_GENERAL, "SerialNumber", "", p11d32.IniPathAndFileGlobal)
      End If
    End If
    
    If (Len(p11d32.SerialNumber) = 0) Then Call Err.Raise(ERR_WORKING_DIRECTORY, "IniSettings", "Failed to enter a valid serial number, exiting the program")
    
    s = GetIniEntry(S_INI_SECTION_GENERAL, "WorkingDirectory", IniSettingLastYear(S_INI_SECTION_GENERAL, "WorkingDirectory", "", True), p11d32.IniPathAndFile)
    Call CreateAndSetWorkingDirectory(Nothing, s)
    Call CreateUserDirectoriesAndData
    
    
    'this is the last file selected
    p11d32.ReportPrint.EmployeeLetterFile = GetIniEntry(S_INI_SECTION_EMPLOYEE_LETTER, "EmployeeLetterFile", LetterFile & S_EMPLOYEE_LETTER_FILE_EXTENSION, p11d32.IniPathAndFile)
    p11d32.ReportPrint.EmployeeLetterPath = GetIniEntry(S_INI_SECTION_EMPLOYEE_LETTER, "EmployeeLetterPath", p11d32.ReportPrint.SystemLettersPath, p11d32.IniPathAndFile)
    
    ' FullPath(CurDir()) = p11d32.WorkingDirectory
    ReportPrint.DefaultReportIndex = CLng(GetIniEntry(S_INI_SECTION_REPORTS, "DefaultReportIndex", RPT_FIRST_INI, p11d32.IniPathAndFile))
    ReportPrint.DefaultSelectEmployeeReportIndex = CLng(GetIniEntry(S_INI_SECTION_REPORTS, "DefaultSelectEmployeeReportIndex", RPT_FIRST_MANAGEMENT, p11d32.IniPathAndFile))
    
    For i = [RPT_FIRST_INI] To [RPT_LAST_INI]
      ReportPrint.Orientation(i) = CLng(GetIniEntry(S_INI_SECTION_REPORTS, ReportPrint.Name(i), REPORT_ORIENTATION.PORTRAIT, p11d32.IniPathAndFile))
    Next
    
    ReportPrint.UserReportFileLessExtension = GetIniEntry(S_INI_SECTION_REPORTS, "UserReportFileLessExtension", "", p11d32.IniPathAndFile)
    ReportPrint.UserReportSelectEmployeeFileLessExtension = GetIniEntry(S_INI_SECTION_REPORTS, "UserReportSelectEmployeeFileLessExtension", "", p11d32.IniPathAndFile)
    
    ReportPrint.EmployeeSelection = CLng(GetIniEntry(S_INI_SECTION_REPORTS, "EmployeeSelection", ReportPrint.EmployeeSelection, p11d32.IniPathAndFile))
    ReportPrint.HMITSelectionChoice = CLng(GetIniEntry(S_INI_SECTION_REPORTS, "HMITSelectionChoice", ReportPrint.HMITSelectionChoice, p11d32.IniPathAndFile))
    ReportPrint.HMITSections = CLng(GetIniEntry(S_INI_SECTION_REPORTS, "HMITSections", 0, p11d32.IniPathAndFile))
    ReportPrint.P46Range = GetIniEntry(S_INI_SECTION_REPORTS, "P46Range", ReportPrint.P46Range, p11d32.IniPathAndFile)
    
    ReportPrint.P46PrintReplacedP46s = GetIniEntry(S_INI_SECTION_REPORTS, "P46PrintReplacedP46s", False, p11d32.IniPathAndFile)
    
    Call p11d32.Rates.GetP46QuarterStartEnd(dQuarterStart, dQuarterEnd, (ReportPrint.P46Range + 1))
    
    
    ReportPrint.P46DateFrom = TryConvertDateDMY(GetIniEntry(S_INI_SECTION_REPORTS, "P46DateFrom", DateStringEx(dQuarterStart, p11d32.Rates.value(TaxYearStart)), p11d32.IniPathAndFile), p11d32.Rates.value(TaxYearStart))
    ReportPrint.P46DateFrom = DateTaxYearFix(True, ReportPrint.P46DateFrom)
    ReportPrint.P46DateTo = TryConvertDateDMY(GetIniEntry(S_INI_SECTION_REPORTS, "P46DateTo", DateStringEx(dQuarterEnd, p11d32.Rates.value(TaxYearEnd)), p11d32.IniPathAndFile), p11d32.Rates.value(TaxYearEnd))
    ReportPrint.P46DateTo = DateTaxYearFix(False, ReportPrint.P46DateTo)

    ReportPrint.P46UserDateFrom = GetIniEntry(S_INI_SECTION_REPORTS, "P46UserDateFrom", DateStringEx(dQuarterStart, p11d32.Rates.value(TaxYearStart)), p11d32.IniPathAndFile)
    ReportPrint.P46UserDateTo = GetIniEntry(S_INI_SECTION_REPORTS, "P46UserDateTo", DateStringEx(dQuarterEnd, p11d32.Rates.value(TaxYearEnd)), p11d32.IniPathAndFile)
    
    'p11d32.ReportPrint.RunChecks = GetIniEntry(S_INI_SECTION_REPORTS, "RunChecks", vbChecked, p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.CheckOptions = GetIniEntry(S_INI_SECTION_REPORTS, "CheckOptions", CStr(YES_THIS_TIME_ONLY), p11d32.IniPathAndFile)
    p11d32.ReportPrint.CHECKORDERBY = GetIniEntry(S_INI_SECTION_REPORTS, "CheckOrderBy", CStr(ORDER_PNUM), p11d32.IniPathAndFile)
    p11d32.ReportPrint.ChecksAutoRefresh = GetIniEntry(S_INI_SECTION_REPORTS, "ChecksAutoRefresh", False, p11d32.IniPathAndFile)
    p11d32.ReportPrint.A4ForcePrint = GetIniEntry(S_INI_SECTION_REPORTS, "A4ForcePrint", True, p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.ChecksSelected = GetIniEntry(S_INI_SECTION_REPORTS, "ChecksSelected", CStr(0), p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.TimeStamp = GetIniEntry(S_INI_SECTION_REPORTS, "TimeStamp", vbChecked, p11d32.IniPathAndFile)
    p11d32.ReportPrint.PrintedEmployees = GetIniEntry(S_INI_SECTION_REPORTS, "PrintedEmployees", False, p11d32.IniPathAndFile) 'JN
    p11d32.ReportPrint.EmployeeLetterMargin = GetIniEntry(S_INI_SECTION_REPORTS, "EmployeeLetterMargin", 10, p11d32.IniPathAndFile)
    p11d32.ReportPrint.EmployeeLetterFontName = GetIniEntry(S_INI_SECTION_REPORTS, "EmployeeLetterFontName", "Arial", p11d32.IniPathAndFile)
    p11d32.ReportPrint.EmployeeLetterFontSize = GetIniEntry(S_INI_SECTION_REPORTS, "EmployeeLetterFontSize", "11", p11d32.IniPathAndFile)
    If (p11d32.ReportPrint.EmployeeLetterFontSize < 6) Then
      p11d32.ReportPrint.EmployeeLetterFontSize = 11
    End If
    
    p11d32.ReportPrint.P11DTotalValue = GetIniEntry(S_INI_SECTION_REPORTS, "P11DTotalValue", P11D_TOTAL_VALUE.PTV_NONZERO_AND_ZERO, p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.Destination = GetIniEntry(S_INI_SECTION_REPORTS, "Destination", REPORT_DESTINATION.REPD_PRINTER_PREVIEW, p11d32.IniPathAndFile)
    
    'p11d32.ReportPrint.ExportDirectory = GetIniEntry(S_INI_SECTION_REPORTS, "ExportDirectory", p11d32.ReportPrint.ExportDirDefault, P11D32.IniPathAndFile)
    p11d32.ReportPrint.PRNFileName = GetIniEntry(S_INI_SECTION_REPORTS, "PRNFileName", "Report", p11d32.IniPathAndFile)
    p11d32.ReportPrint.PRNFileExtension = GetIniEntry(S_INI_SECTION_REPORTS, "PRNFileExtension", ".prn", p11d32.IniPathAndFile)
    p11d32.ReportPrint.EmailTextAsHTML = GetIniEntry(S_INI_SECTION_REPORTS, "EmailTextAsHTML", False, p11d32.IniPathAndFile)
    p11d32.ReportPrint.HMITFieldTrim = GetIniEntry(S_INI_SECTION_REPORTS, "HMITFieldTrim", True, p11d32.IniPathAndFile)
    p11d32.ReportPrint.PRNLoopLength = GetIniEntry(S_INI_SECTION_REPORTS, "PRNLoopLength", p11d32.ReportPrint.PRNLoopLength, p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.RemeberEmployeeSelection = GetIniEntryBool(S_INI_SECTION_REPORTS, "RemeberEmployeeSelection", True)
    
    
    
    ' rdc
    p11d32.ReportPrint.DatesOnWorkingPaper = GetIniEntry(S_INI_SECTION_REPORTS, "DatesOnWorkingPaper", False, p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.ExportOption = GetIniEntry(S_INI_SECTION_REPORTS, "ExportOption", EXPORT_HTML_INTEXP5, p11d32.IniPathAndFile)   'km
    p11d32.ReportPrint.ExportAllowUserNameHTML = GetIniEntry(S_INI_SECTION_REPORTS, "ExportAllowUserNameHTML", False, p11d32.IniPathAndFile)
    p11d32.ReportPrint.ExportEmployeeFileNameFields = GetIniEntry(S_INI_SECTION_REPORTS, "ExportEmployeeFileNameFields", REPORT_EXPORT_EMPLOYEE_FILE_NAME_FIELDS.REEFNF_PERSONNEL_NUMBER, p11d32.IniPathAndFile)
    
    
    
    p11d32.ReportPrint.EmailSubject = GetIniEntry(S_INI_SECTION_REPORTS, "EmailSubject", "P11D Information", p11d32.IniPathAndFile)
    p11d32.ReportPrint.SeparatePrintJobs = GetIniEntry(S_INI_SECTION_REPORTS, "SeparatePrintJobs", False, p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.NotesIniFile = GetIniEntry(S_INI_SECTION_REPORTS, "NotesIniFile", FullPath(GetWindowsDirectory) & "Notes.ini", p11d32.IniPathAndFile)
    p11d32.ReportPrint.UseNotesForEmail = GetIniEntry(S_INI_SECTION_REPORTS, "UseNotesForEmail", False, p11d32.IniPathAndFile)
    
    p11d32.MagneticMedia.DataFormat = GetIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "DataFormat", MM_DF_ANY, p11d32.IniPathAndFile)
    p11d32.MagneticMedia.LastPathAndFileCreated = GetIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "LastPathAndFileCreated", "", p11d32.IniPathAndFile)
    p11d32.MagneticMedia.RecordViewID = GetIniEntry(S_INI_SECTION_MAGNETIC_MEDIA, "RecordViewID", "", p11d32.IniPathAndFile)

    p11d32.PAYEonline.LastPathAndFileCreated = GetIniEntry(S_INI_SECTION_PAYEONLINE, "LastPathAndFileCreated", "", p11d32.IniPathAndFile)
    p11d32.PAYEonline.P46Range = GetIniEntry(S_INI_SECTION_PAYEONLINE, "P46Range", PAYEonline.P46Range, p11d32.IniPathAndFile)
    If p11d32.PAYEonline.P46Range < P46_QUARTER1 Or p11d32.PAYEonline.P46Range > P46_QUARTER4 Then
      p11d32.PAYEonline.P46Range = P46_QUARTER1
    End If
    PAYEonline.P46DateFrom = TryConvertDateDMY(GetIniEntry(S_INI_SECTION_REPORTS, "P46DateFrom", DateStringEx(dQuarterStart, p11d32.Rates.value(TaxYearStart)), p11d32.IniPathAndFile), p11d32.Rates.value(TaxYearStart))
    PAYEonline.P46DateFrom = DateTaxYearFix(True, PAYEonline.P46DateFrom)
    PAYEonline.P46DateTo = TryConvertDateDMY(GetIniEntry(S_INI_SECTION_REPORTS, "P46DateTo", DateStringEx(dQuarterEnd, p11d32.Rates.value(TaxYearEnd)), p11d32.IniPathAndFile), p11d32.Rates.value(TaxYearEnd))
    PAYEonline.P46DateTo = DateTaxYearFix(False, PAYEonline.P46DateTo)

    p11d32.PAYEonline.P46UserDateFrom = GetIniEntry(S_INI_SECTION_PAYEONLINE, "P46UserDateFrom", DateStringEx(dQuarterStart, p11d32.Rates.value(TaxYearStart)), p11d32.IniPathAndFile)
    p11d32.PAYEonline.P46UserDateTo = GetIniEntry(S_INI_SECTION_PAYEONLINE, "P46UserDateTo", DateStringEx(dQuarterEnd, p11d32.Rates.value(TaxYearEnd)), p11d32.IniPathAndFile)
    'order of true and false in 'iif' statements  must not be changed
    
    p11d32.PAYEonline.Efiler_Test_Submission = GetIniEntryBool(S_INI_SECTION_PAYEONLINE, "Efiler_test_submission", False)
    p11d32.PAYEonline.Efiler_Mock_Submission = GetIniEntryBool(S_INI_SECTION_PAYEONLINE, "Efiler_mock_submission", False)
    p11d32.PAYEonline.Efiler_Proceed_Submission = GetIniEntryBool(S_INI_SECTION_PAYEONLINE, "Efiler_proceed_submission", True)
    p11d32.PAYEonline.Type_Selected = GetIniEntry(S_INI_SECTION_PAYEONLINE, "Type", PAYEonline_TYPES.POT_P11D, p11d32.IniPathAndFile)
    p11d32.PAYEonline.ViewProceedButtonIfErrors = GetIniEntryBool(S_INI_SECTION_PAYEONLINE, "ViewProceedButtonIfErrors", False)
    p11d32.PAYEonline.ByPassLocalSchemaCheck = GetIniEntryBool(S_INI_SECTION_PAYEONLINE, "ByPassLocalSchemaCheck", True)
    p11d32.PAYEonline.DisableCheckSubmissionWithAbatec = GetIniEntryBool(S_INI_SECTION_PAYEONLINE, "DisableCheckSubmissionWithAbatec", False)
    
    p11d32.BringForward.BringForwardType = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "BringForwardType", BRING_FORWARD_TYPE.BFT_OVERWRITE, p11d32.IniPathAndFile)
    p11d32.BringForward.HMITSChosen = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "HMITSChosen", -1, p11d32.IniPathAndFile)  'all is default
    p11d32.BringForward.CDCs = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "CDCs", vbChecked, p11d32.IniPathAndFile)
    p11d32.BringForward.CDBs = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "CDBs", vbChecked, p11d32.IniPathAndFile)
    p11d32.BringForward.Comments = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "comments", vbChecked, p11d32.IniPathAndFile)
    
    p11d32.BringForward.SharedVans = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "SharedVans", vbChecked, p11d32.IniPathAndFile)
    'AM To be removed
    p11d32.BringForward.NewFilesForAll = GetIniEntry(S_INI_SECTION_BRING_FORWARD, "NewFilesForAll", vbChecked, p11d32.IniPathAndFile)
    
    p11d32.NameOrder = GetIniEntry(S_INI_SECTION_GENERAL, "NameOrder", NO_SURNAME_TITLE_FN_INITIALS, p11d32.IniPathAndFile)
    
    p11d32.EmployeeSortOrderColumn = GetIniEntry(S_INI_SECTION_GENERAL, "EmployeeSortOrderColumn", LV_EE_NAME, p11d32.IniPathAndFile)
    p11d32.EmployeeSortOrder = GetIniEntry(S_INI_SECTION_GENERAL, "EmployeeSortOrder", ListSortOrderConstants.lvwAscending, p11d32.IniPathAndFile)
    
    p11d32.FindFilesSearchSubDirs = GetIniEntry(S_INI_SECTION_GENERAL, "FindFilesSearchSubDirs", False, p11d32.IniPathAndFile) 'JN
    p11d32.FindFilesDirList = GetIniEntry(S_INI_SECTION_GENERAL, "FindFilesDirList", "", p11d32.IniPathAndFile)   'SO
    
    p11d32.RegularPaymentsMethod = GetIniEntry(S_INI_SECTION_GENERAL, "RegularPaymentsMethod", RPM_MONTHLY, p11d32.IniPathAndFile)
    p11d32.RegularPaymentsForceEndOfMonth = GetIniEntry(S_INI_SECTION_GENERAL, "RegularPaymentsForceEndOfMonth", False, p11d32.IniPathAndFile)
    
    
    p11d32.KillBenefits = GetIniEntry(S_INI_SECTION_GENERAL, "KillBenefits", False, p11d32.IniPathAndFile)
    p11d32.CompanyCarCheckerCheck = GetIniEntry(S_INI_SECTION_GENERAL, "CompanyCarCheckerCheck", CCCC_DATES, p11d32.IniPathAndFile)
    
    p11d32.PayeReferenceAnyFormat = GetIniEntryBool(S_INI_SECTION_GENERAL, "PayeReferenceAnyFormat", False)
    
    p11d32.DataTypeListViewSorting = GetIniEntryBool(S_INI_SECTION_GENERAL, "DataTypeListViewSorting", True)
    
    p11d32.SortEmployeeReferenceAsNumber = GetIniEntryBool(S_INI_SECTION_GENERAL, "SortEmployeeReferenceAsNumber ", False)
    
    
    p11d32.MoveToNextEmployeeWithBenefit = GetIniEntryBool(S_INI_SECTION_GENERAL, "MoveToNextEmployeeWithBenefit", False)
    
    p11d32.Importing.Tracking = GetIniEntryBool(S_INI_SECTION_GENERAL, "ImportTracking", False)
    p11d32.ValidateNINumberOnEmployeeScreen = GetIniEntryBool(S_INI_SECTION_GENERAL, "ValidateNINumberOnEmployeeScreen", True)
    
    p11d32.VersionCheckEnabled = GetIniEntryBool(S_INI_SECTION_GENERAL, "VersionCheck", True)
    
    p11d32.Intranet.LoginUserNameSource = GetIniEntry(S_INI_SECTION_INTRANET, "LoginUserNameSource", INTRANET_LOGIN_USERNAME_SOURCE.ILUS_USERNAME, p11d32.IniPathAndFile)
    p11d32.Intranet.AuthenticationType = GetIniEntry(S_INI_SECTION_INTRANET, "AuthenticationType", INTRANET_AUTHENTICATION_TYPE.IAT_FULL, p11d32.IniPathAndFile)
    p11d32.Intranet.OutputType = GetIniEntry(S_INI_SECTION_INTRANET, "OutputType", INTRANET_OUTPUT_TYPE.IOT_P11D, p11d32.IniPathAndFile)
    p11d32.Intranet.BannerTitle = GetIniEntry(S_INI_SECTION_INTRANET, "BannerTitle", "Tax Solutions", p11d32.IniPathAndFile)
    p11d32.Intranet.UserInfoHTML = GetIniEntry(S_INI_SECTION_INTRANET, "UserInfoHTML", "Enter additional info here", p11d32.IniPathAndFile)
    p11d32.Intranet.CaseSensitiveOnFullAuthentication = GetIniEntryBool(S_INI_SECTION_INTRANET, "CaseSensitiveOnFullAuthentication", False)
    
    p11d32.Intranet.BannerBackColor = GetIniEntry(S_INI_SECTION_INTRANET, "BannerBackColor", "16744576", p11d32.IniPathAndFile)
    p11d32.Intranet.BannerForeColor = GetIniEntry(S_INI_SECTION_INTRANET, "BannerForeColor", "16777215", p11d32.IniPathAndFile)
    
    p11d32.ReportPrint.SortOtherTypeBenefitsAlphabetically = GetIniEntry(S_INI_SECTION_REPORTS, "SortOtherTypeBenefitsAlphabetically", False, p11d32.IniPathAndFile)

  End If
  
  IniSettings = True
  
IniSettings_END:
  Call xReturn("IniSettings")
  Exit Function
IniSettings_ERR:
  If (Err.Number = ERR_WORKING_DIRECTORY) Then
    Call Err.Raise(Err.Number, "IniSettings", Err.Description)
  End If
  If IRW = Ini_read Then
    Call ErrorMessage(ERR_ERROR, Err, "IniSettings", "Ini Settings", Err.Description)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "IniSettings", "Ini Settings", Err.Description)
  End If
  Resume IniSettings_END
  Resume
End Function

Private Function CreateUserDirectory(ByVal section As String, ByVal Key As String, ByVal sDefault As String, ByVal bYearSpecific As Boolean) As String
  Dim s As String
  Dim sOriginalValue As String
  
  Dim sLastYearDefault As String
  
  sOriginalValue = sDefault
  If bYearSpecific Then
    sDefault = p11d32.UserDataDirectoryYearSpecific & sDefault
  Else
    sDefault = p11d32.UserDataDirectory & sDefault
  End If
  
  If Not bYearSpecific Then
    sDefault = IniSettingLastYear(section, Key, sDefault, False)
  End If
  s = GetIniEntry(section, Key, sDefault, p11d32.IniPathAndFile)
  
  If Not bYearSpecific Then
   If StrComp(s, FullPath(p11d32.UserDataDirectoryYearSpecific & sOriginalValue), vbTextCompare) = 0 Then
     'force it to be non year specific..... !
     s = FullPath(p11d32.UserDataDirectory & sOriginalValue)
   End If
  Else
    
   
  End If
  'if it going to the wrong place is it is year specific when it should not be year specific
  
  If (Not FileExists(s, True)) Then
    If (Not FileExists(sDefault, True)) Then
      Call MkDirEx(sDefault)
    End If
    s = sDefault
  End If
  Call RemoveReadOnlyFolder(s, True)
  CreateUserDirectory = s
End Function

Public Sub CreateUserDirectoriesAndData()
  Dim s As String
  On Error GoTo err_err
  
  p11d32.Intranet.OutputDirectory = CreateUserDirectory(S_INI_SECTION_INTRANET, "OutputDirectory", S_USERDIR_INTRANET, True)
  p11d32.MagneticMedia.OutputDirectory = CreateUserDirectory(S_INI_SECTION_MAGNETIC_MEDIA, "OutPutDirectory", S_USERDIR_MMEDIA, True)
  
  p11d32.ReportPrint.ExportDirectory = CreateUserDirectory(S_INI_SECTION_REPORTS, "ExportDirectory", S_USERDIR_EXPORT, True)
  
  p11d32.ReportPrint.ReportPathUser = CreateUserDirectory(S_INI_SECTION_REPORTS, "UserReportDirectory", S_USERDIR_UREPORTS, False)
  
  p11d32.ReportPrint.UserLettersPath = CreateUserDirectory(S_INI_SECTION_EMPLOYEE_LETTER, "UserLettersDirectory", S_USERDIR_ULETTERS, False)
  
  
  p11d32.PAYEonline.OutputDirectory = CreateUserDirectory(S_INI_SECTION_PAYEONLINE, "OutPutDirectory", S_USERDIR_PAYEONLINE, True)
  
  p11d32.Importing.TrackingPath = CreateUserDirectory(S_INI_SECTION_GENERAL, "ImportTrackingDirectory", S_USERDIR_IMPORT_TRACKING, True)
  
          
  s = p11d32.UserDataDirectoryYearSpecific & S_PAYEONLINE_MDB
  If Not FileExists(s) Then
    If Not FileCopyEx(FullPath(AppPath) & S_PAYEONLINE_MDB, s) Then
      Call Err.Raise(ERR_WORKING_DIRECTORY, "CreateUserDirectoriesAndData", "Failed to copy the PAYE Online DB to " & s)
    End If
  End If
  
      
  
err_end:
 Exit Sub
err_err:
  
 Call Err.Raise(Err.Number, ErrorSource(Err, "CreateUserDirectories"), Err.Description)
End Sub
Public Sub CreateAndSetWorkingDirectory(frm As Form, ByVal sWorkingDirectory As String, Optional ByVal bChange As Boolean = False)
  Dim s As String
  Dim sCreateMessage As String
  Dim hwnd As Long
  On Error GoTo err_err
  
  'we want the main folder to exist and then the sub folders
  sWorkingDirectory = Trim$(sWorkingDirectory)
  If (bChange) Then GoTo JUST_CHANGE:
  
  If Len(sWorkingDirectory) = 0 Then
    sCreateMessage = "The working directory has not been set. The working directory" & vbCrLf & "is where your P11D data will be stored. It is advisable to" & vbCrLf & "choose a network folder so your data is backed up and others" & vbCrLf & "can share it." & vbCrLf & vbCrLf & "Please click OK and select a directory"
  ElseIf (Not FileExists(sWorkingDirectory, True)) Then
    sCreateMessage = "The working directory specified (" & sWorkingDirectory & ") does not exist." & vbCrLf & _
              "Click ok and select a directory"
  End If
  If (Len(sCreateMessage) > 0) Then
    If (MsgBox(sCreateMessage, vbOKCancel, "Working directory") = vbCancel) Then
      Call Err.Raise(ERR_WORKING_DIRECTORY, "CreateAndSetWorkingDirectories", "Working directory has not been set")
    End If
JUST_CHANGE:
    If Not frm Is Nothing Then
      hwnd = frm.hwnd
    End If
    
    sWorkingDirectory = BrowseForFolderEx(hwnd, CurDir$, "Select a working directory")
    If (Len(sWorkingDirectory) = 0) Then
      If (bChange) Then
        GoTo err_end
      End If
      Call Err.Raise(ERR_WORKING_DIRECTORY, "CreateWorkingDirectories", "Working directory not set")
    End If
    
  End If
  
  ChDriveUNC (sWorkingDirectory)
  ChDir (sWorkingDirectory)
  
  
  sWorkingDirectory = FullPath(sWorkingDirectory)
  p11d32.WorkingDirectory = sWorkingDirectory
    
  Call WriteIniEntry(S_INI_SECTION_GENERAL, "WorkingDirectory", sWorkingDirectory, p11d32.IniPathAndFile)
  Call FlushIniBuffer(p11d32.IniPathAndFile)
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, "CreateAndSetWorkingDirectories", Err.Description)
  Resume
End Sub

Public Property Get LetterFile() As String
  LetterFile = "Letter"
End Property
Public Property Get EmployeeLetterTemplateFile() As String
  EmployeeLetterTemplateFile = "Template"
End Property
Private Sub Class_Initialize()
  'no sub initialises can refer to P11d32 if inited inside init of p11d32
  'so p11d32.initilaise called instead
  
End Sub

Private Sub Class_Terminate()
  Set ReportPrint = Nothing
End Sub
Public Sub StartScreen()
  Call EmployerScreen
  Call LoadEmployers
End Sub
Public Function EmployerScreen() As Boolean
  Dim ibf As IBenefitForm2
  Dim lCurrentEmployerIndex As Long
  
  On Error GoTo EmployerScreen_Err
  
  Call xSet("EmployerScreen")
  Call UpdateBenefitFromTags
  
  lCurrentEmployerIndex = -1
  

  If Not LoadEmployer(Nothing) Then GoTo EmployerScreen_End
  MDIMain.ClearConfirmUndo
  Call ShowMaximized(F_Employers, CurrentForm, D_EMPLOYER_OFF)
  Set CurrentForm = F_Employers
  Set ibf = F_Employers
  
  If ibf.lv.SelectedItem Is Nothing Then
    lCurrentEmployerIndex = -1
  Else
    lCurrentEmployerIndex = ibf.lv.SelectedItem.Tag
  End If
  
  ibf.BenefitsToListView
  
  Call SelectBenefitByBenefitIndex(ibf, lCurrentEmployerIndex)
  MDIMain.sts.ClearCaptions
EmployerScreen_End:

  Call xReturn("EmployerScreen")
  Exit Function
  
EmployerScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EmployerScreen", "ERR_EMPR_SCREEN", "Error moving to the employer screen.")
  Resume EmployerScreen_End
  Resume
End Function
Public Function LoadEmployer(clsER As Employer, Optional bEmployeeScreen As Boolean = True, Optional bForceReload As Boolean = False, Optional bLoadExclusive As Boolean = False) As Boolean
  Dim ibc As IBenefitClass
  Dim ben As IBenefitClass
  
  On Error GoTo LoadEmployer_Err
  
  Call xSet("LoadEmployer")
  
  Call SetPanel2("Opening employer file")
  
  If Not CurrentEmployer Is Nothing Then
    If Not CurrentEmployer.LoadEmployeeEx(Nothing) Then GoTo LoadEmployer_End
  End If
  
  If clsER Is CurrentEmployer And Not bForceReload Then
    LoadEmployer = True
    GoTo LoadEmployer_End
  End If
  
  If Not CurrentEmployer Is Nothing Then
    Set ben = CurrentEmployer
    Call ben.Kill ' removes all employees, references and sets all rsbentables to nothing
  End If
  
  Set CurrentEmployer = clsER
  If Not clsER Is Nothing Then
    If Not clsER.PassWordOK() Then GoTo LoadEmployer_End
    Call clsER.OpenEmployer(bLoadExclusive)
    
    Set CurrentEmployer.rsBenTables(TBL_COMPANY_CARS) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_CARS), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_VANS) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_VANS), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_SHARED_VANS) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_SHARED_VANS), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_OTHER) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_OTHERS), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_CDB_LINKS) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_CDB_LINKS), dbOpenDynaset)
    
    Set CurrentEmployer.rsBenTables(TBL_PHONES) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_PHONE), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_ASSETSATDISPOSAL) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_ASSETSATDISPOSAL), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_SERVICESPROVIDED) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_SERVICESPROVIDED), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_BENLOAN) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_BENLOANS), dbOpenDynaset)
'MP DB confirm removal - not needed for any purpose?
'    Set CurrentEmployer.rsBenTables(TBL_HOMELOAN) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_HOMELOANS), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_ACCOM) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_ACCOM), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_RELOC) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_RELOC), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_EMPLOYEE_CARS) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_EECARS), dbOpenDynaset)
    Set CurrentEmployer.rsBenTables(TBL_TRANS) = CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_TRANS), dbOpenDynaset)
    Call BenDataLinksCreate(CurrentEmployer)
    Call CurrentEmployer.LoadEmployees
    Set ibc = CurrentEmployer.SharedVans
    ibc.ReadFromDB = False
    Call ibc.ReadDB
    Call CurrentEmployer.ReadFPCS
    Call CurrentEmployer.ReadCDCs
'MP DB confirm removal - not needed for any purpose?
'    Call CurrentEmployer.FindAccomExpenses
    If udm Is Nothing Then Set udm = New udm
    If bEmployeeScreen Then Call p11d32.CurrentEmployer.EmployeeScreen
  End If
  LoadEmployer = True
  
LoadEmployer_End:
  Set ibc = Nothing
  Call PrgStopCaption
  Call xReturn("LoadEmployer")
  Exit Function
  
LoadEmployer_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadEmployer", "Load Employer", "Unable to load employer file.")
  If Not CurrentEmployer Is Nothing Then
    Set ben = CurrentEmployer
    Call ben.Kill
    Set CurrentEmployer = Nothing
  End If
  Resume LoadEmployer_End
  Resume
End Function
Public Function EmployeeSelectedForPrint(benEE As IBenefitClass, Employer As Employer) As Boolean
  Dim es As EMPLOYEE_SELECTION
  Dim ee As Employee
  On Error GoTo EmployeeSelectedForPrint_err
  
  If Employer Is Nothing Then Call Err.Raise(ERR_ERROR, ERR_EMPLOYER_INVALID, "EmployeeSelectedForPrint", "The employer passed is nothing.")
  es = ReportPrint.EmployeeSelection
  
  If Not benEE Is Nothing Then
    Select Case es
      Case ES_ALL
        EmployeeSelectedForPrint = True
      Case ES_INVERSE_SELECTED
        If Not benEE.value(ee_Selected) Then
          EmployeeSelectedForPrint = True
        End If
      Case ES_SELECTED
        If benEE.value(ee_Selected) Then
          EmployeeSelectedForPrint = True
        End If
      Case ES_CURRENT
        Set ee = p11d32.CurrentEmployer.employees(F_Employees.LB.SelectedItem.Tag)
        If benEE Is ee Then
          EmployeeSelectedForPrint = True
        End If
      Case Else
        ECASE ("Invalid employee print selection")
    End Select
  End If

EmployeeSelectedForPrint_end:
  Exit Function
EmployeeSelectedForPrint_err:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeSelectedForPrint", "Employee Selected For Print", "Error determining if an employee is selected for printing.")
  Resume EmployeeSelectedForPrint_end
End Function
Public Function GetWS(lFileYear As Long) As Workspace
  
End Function
Public Property Get SystemMDWPath()
  SystemMDWPath = FullPath(AppPath) & "SYSTEM.MDW"
End Property

Private Function SetSystemMDA() As Boolean
  Dim sSystemPathAndFile As String
  
  On Error GoTo SetSystemMDA_err
  
  sSystemPathAndFile = SystemMDWPath
  
  If Not FileExists(sSystemPathAndFile) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "SetSystemMDA", "The system file " & sSystemPathAndFile & " does not exist.")
  DBEngine.SystemDB = sSystemPathAndFile
  DBEngine.DefaultUser = Me.DBUserName
  DBEngine.DefaultPassword = Me.DBPassword
  
  Set P11DWS = DBEngine.Workspaces(0)
  
  SetSystemMDA = True
SetSystemMDA_end:
  Exit Function
SetSystemMDA_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetSystemMDA", "P11D Initialisation", "Unable to initialise system workspace (" & DBEngine.SystemDB & ")")
  
  Resume SetSystemMDA_end
  Resume
End Function

Public Property Get DBUserName() As String
    DBUserName = "admin"
  
End Property
Public Property Get DBPassword()
    DBPassword = ""
 
End Property
Public Function LoadEmployers() As Long
  Dim empr As Employer
  Dim i As Long
  Dim s As String, q  As String
  
  On Error GoTo LoadEmployers_Err
  
  Call xSet("LoadEmployers")
    
  MDIMain.mnuFile.Enabled = False
  'WorkingDirectory = FullPath(CurDir()) ' EK Working Dir change, may need to comment out
  F_Employers.lblEmployersDirectory.Caption = S_EMPLOYER_DIRECTORY_LABEL & WorkingDirectory
  Call Employers.RemoveAll
  Employers.Compact
  'use this as use in BringForward and did not want to duplicate
  Call EnumEmployerFiles(Rates.FileExtensionCurrent, Me)
  Call BenScreenSwitchEnd(F_Employers)
  LoadEmployers = Employers.Count
  
LoadEmployers_End:
  MDIMain.mnuFile.Enabled = True
  Call PrgStopCaption
  Call xReturn("LoadEmployers")
  Exit Function

LoadEmployers_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadEmployers", "ERR_LOADEMPLOYERS", "Error reading the employer files.")
  Resume LoadEmployers_End
  Resume
End Function
Public Function IfNotCurrentEmployerThenLoad(empr As Employer) As Boolean
  Dim bEmployerLoaded As Boolean
  Dim benEmployer As IBenefitClass
  
  On Error GoTo IfNotCurrentEmployerThenLoad_END
  
  Call xSet("IfNotCurrentEmployerThenLoad")
  Set empr = CurrentEmployer
  
  IfNotCurrentEmployerThenLoad = Not empr Is Nothing
  If (Not IfNotCurrentEmployerThenLoad) And (CurrentForm Is F_Employers) Then
    If F_Employers.LB.SelectedItem Is Nothing Then Call Err.Raise(ERR_NO_EMPLOYER, "IfNotCurrentEmployerThenLoad", "Unable to load an employer.")
    Set empr = p11d32.Employers(F_Employers.LB.SelectedItem.Tag)
    Set benEmployer = empr
    If Not empr.ValidateEx(benEmployer.value(employer_PathAndFile), False, False, True) Then GoTo IfNotCurrentEmployerThenLoad_END
    IfNotCurrentEmployerThenLoad = True
  End If
  
IfNotCurrentEmployerThenLoad_END:
  Call xSet("IfNotCurrentEmployerThenLoad")
  Exit Function
IfNotCurrentEmployerThenLoad_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IfNotCurrentEmployerThenLoad", "If Not Current Employer Then Load", "Error loading an employer.")
  Resume IfNotCurrentEmployerThenLoad_END
End Function

Public Sub DebugSQL()
  Dim ibf As IBenefitForm2
  Dim sqld As SQLDebug
  Dim ey As Employer
  Dim db As Database
  Dim ben As IBenefitClass
  
  On Error GoTo DebugSQL_ERR
  Call xSet("DebugSQL")
  If Not CurrentForm Is F_Employers Then Err.Raise ERR_INVALID_FORM, "DebugSQL", "Current form is not the employer screen."
  Set sqld = New SQLDebug
  Set ibf = CurrentForm
  If IfNotCurrentEmployerThenLoad(ey) Then
    Set sqld.Workspace = P11DWS
    Call sqld.Show(ey.db)
    'Set ben = ey
    'If Not ey Is CurrentEmployer Then ben.Kill
  Else
    Call sqld.Show
  End If
  
DebugSQL_END:
  Call xReturn("DebugSQL")
  Exit Sub
DebugSQL_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DebugSQL", "Debug SQL", "Error starting debugging db SQL.")
  Resume DebugSQL_END
  Resume
End Sub
Public Sub SplitNamesRun()
  Dim empr As Employer
  Dim ben As IBenefitClass
  
  On Error GoTo SplitNamesRun_ERR
  
  Call xSet("SplitNamesRun")
  Call OnlyFromForm(F_Employers)
  
  If Not IfNotCurrentEmployerThenLoad(empr) Then GoTo SplitNamesRun_END
  Call empr.db.Execute(sql.Queries(UPDATE_NAME_PARTS_TO_NULL))
  Call empr.SplitFix(SELECT_EMPLOYEES_NO_CDB)
  Set ben = empr
  If Not empr Is CurrentEmployer Then ben.Kill
  
SplitNamesRun_END:
  Call xReturn("SplitNamesRun")
  Exit Sub
SplitNamesRun_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SplitNamesRun", "Split Names Run", "Error running split names.")
  Resume SplitNamesRun_END
  Resume
End Sub


Public Sub DeleteAllCDBS()
  On Error GoTo DeleteAllCDBS_ERR
  
  Call xSet("DeleteAllCDBS")
  
  Dim empr As Employer
  
  If Not IfNotCurrentEmployerThenLoad(empr) Then GoTo DeleteAllCDBS_END
  If empr.db Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "DeleteAllCDBS", "Database is nothing.")
  
  empr.db.Execute (sql.Queries(DELETE_CDBS_DEBUG))
  empr.db.Execute (sql.Queries(DELETE_CDB_LINKS_DEBUG))
  
DeleteAllCDBS_END:
  Call xReturn("DeleteAllCDBS")
  Exit Sub
DeleteAllCDBS_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DeleteAllCDBS", "Delete All CDBS", "Error deleting all the CDBs.")
  Resume DeleteAllCDBS_END
  Resume
End Sub
Public Sub UpdateFixLevel()
  
  Dim empr As Employer
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo UpdateFixLevel_ERR
  
  Call xSet("UpdateFixLevel")
  
  If Not IfNotCurrentEmployerThenLoad(empr) Then GoTo UpdateFixLevel_END
  Set ben = empr
  
  F_Input.ValText.Maximum = p11d32.TemplateFixlevel
  F_Input.ValText.Minimum = 0
  F_Input.ValText.TypeOfData = VT_LONG
  If F_Input.Start("Update fix level", "Current fix level = " & ben.value(employer_FixLevel_db) & ", min = " & 0 & ", max = " & p11d32.TemplateFixlevel, ben.value(employer_FixLevel_db)) Then
      Call empr.db.Execute(sql.Queries(UPDATE_FIX_LEVEL, F_Input.ValText.Text))
      ben.value(employer_FixLevel_db) = F_Input.ValText.Text
      Set ibf = F_Employers
      Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, empr)
  End If
  Set F_Input = Nothing
  
  If Not empr Is CurrentEmployer Then ben.Kill
  
UpdateFixLevel_END:
  Call xReturn("UpdateFixLevel")
  Exit Sub
UpdateFixLevel_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateFixLevel", "Update Fix Level", "Error setting the fix level.")
  Resume UpdateFixLevel_END
  Resume
End Sub



Public Function GetBenefitForm() As IBenefitForm2
  
  On Error GoTo GetBenefitForm_err
  
  Set GetBenefitForm = CurrentForm
  'If GetBenefitForm.benclass = BC_ALL Then Set GetBenefitForm = Nothing

GetBenefitForm_end:
  Exit Function
  
GetBenefitForm_err:
  Resume GetBenefitForm_end
End Function

Public Function EditEmployer(ByVal lEmployerIndex As Long) As Long
  Dim empr As Employer
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo EditEmployer_Err
  Call xSet("EditEmployer")
  
  If lEmployerIndex > 0 Then
    Set empr = Employers(lEmployerIndex)
    If Not empr.PassWordOK() Then GoTo EditEmployer_End
    
    
    Set ben = empr
    If ben.value(employer_FixLevel_db) < p11d32.TemplateFixlevel Then Call Err.Raise(ERR_FIX_LEVEL_LOW, "EditEmployer", "The fix level is to low to edit employer, please open the employer first, fixlevel = " & ben.value(employer_FixLevel_db))
    F_EmployerDetails.TxtBx(0).Text = ben.Name
    F_EmployerDetails.TxtBx(1).Text = ben.value(employer_Payeref_db)
    F_EmployerDetails.TxtBx(2).Text = Left$(ben.value(employer_FileName), Len(ben.value(employer_FileName)) - 6)
    F_EmployerDetails.TxtBx(2).Enabled = False
    F_EmployerDetails.TxtBx(3).Text = ben.value(employer_Signatory_db)
    F_EmployerDetails.TxtBx(4).Text = ben.value(employer_Contact_db)
    F_EmployerDetails.TxtBx(5).Text = ben.value(employer_contactnumber_db)
    F_EmployerDetails.chkbx(0) = BoolToChkBox(ben.value(employer_CT_db))
    F_EmployerDetails.chkbx(1) = BoolToChkBox(ben.value(employer_TreatEmployeeCarsAsUnderApprovedFPCS_db))
    F_EmployerDetails.chkbx(2) = BoolToChkBox(ben.value(employer_CheapTaxableLoansAllNon_db))
    F_EmployerDetails.chkbx(3) = BoolToChkBox(ben.value(employer_EerDeclSect2Chk1_db))
    F_EmployerDetails.chkbx(4) = BoolToChkBox(ben.value(employer_EerDeclSect2Chk2_db))
    F_EmployerDetails.chkbx(5) = BoolToChkBox(ben.value(employer_EerDeclSect2Chk3_db))
    
    F_EmployerDetails.TxtBx(6).Text = ben.value(employer_AddressLine1_db)
    F_EmployerDetails.TxtBx(7).Text = ben.value(employer_AddressLine2_db)
    F_EmployerDetails.TxtBx(8).Text = ben.value(employer_AddressLine3_db)
    F_EmployerDetails.TxtBx(9).Text = ben.value(employer_AddressLine4_db)
    F_EmployerDetails.TxtBx(10).Text = ben.value(employer_AddressPostCode_db)
    F_EmployerDetails.TxtBx(11).Text = ben.value(employer_SubmitterRef_db)
    F_EmployerDetails.TxtBx(12).Text = ben.value(employer_SubmitterName_db)
    F_EmployerDetails.TxtBx(14).Text = ben.value(employer_AddClass1AAmounts_db)
    F_EmployerDetails.TxtBx(15).Text = ben.value(employer_AddClass1ADescription_db)
    F_EmployerDetails.TxtBx(16).Text = ben.value(employer_deductClass1AAmounts_db)
    F_EmployerDetails.TxtBx(17).Text = ben.value(employer_deductClass1ADescription_db)
    F_EmployerDetails.TxtBx(18).Text = ben.value(employer_IRTaxOffice_db)
    F_EmployerDetails.TxtBx(19).Text = ben.value(employer_IRAddressLine1_db)
    F_EmployerDetails.TxtBx(20).Text = ben.value(employer_IRAddressLine2_db)
    F_EmployerDetails.TxtBx(21).Text = ben.value(employer_IRAddressLine3_db)
    F_EmployerDetails.TxtBx(22).Text = ben.value(employer_IRAddressLine4_db)
    F_EmployerDetails.TxtBx(23).Text = ben.value(employer_IRPostcode_db)
    F_EmployerDetails.TxtBx(24).Text = ben.value(employer_IRTelephoneNumber_db)
    F_EmployerDetails.TxtBx(25).Text = ben.value(employer_IRYourReference_db)
    F_EmployerDetails.TxtBx(26).Text = ben.value(employer_IRAccountsOfficeReference_db)
    F_EmployerDetails.TxtBx(27).Text = ben.value(employer_PayeOnlineID_db)
    F_EmployerDetails.TxtBx(28).Text = ben.value(employer_PayeOnlinePwd_db)
    F_EmployerDetails.TxtBx(29).Text = ben.value(employer_PayeOnlineUser_db)
    F_EmployerDetails.TxtBx(30).Text = ben.value(employer_PayeOnlineEmail_db)
    F_EmployerDetails.TxtBx(31).Text = ben.value(employer_GroupCode1Alias_db)
    F_EmployerDetails.TxtBx(32).Text = ben.value(employer_GroupCode2Alias_db)
    F_EmployerDetails.TxtBx(33).Text = ben.value(employer_GroupCode3Alias_db)
    
    
    
  Else
    Set empr = New Employer
    Set ben = empr
    ben.value(employer_EmployeeResponseDate_db) = p11d32.Rates.value(EmpLetEmployeeResponseDate)
    F_EmployerDetails.TxtBx(31).Text = S_GROUP_CODE1
    F_EmployerDetails.TxtBx(32).Text = S_GROUP_CODE2
    F_EmployerDetails.TxtBx(33).Text = S_GROUP_CODE3
    F_EmployerDetails.TxtBx(2).Enabled = True
  End If
  
  F_EmployerDetails.TxtBx(13).Text = DateValReadToScreen(ben.value(employer_EmployeeResponseDate_db))
  
retryedit:
'  F_EmployerDetails.Show vbModal
  Call p11d32.Help.ShowForm(F_EmployerDetails, vbModal)
  
  If F_EmployerDetails.m_ok = False Then GoTo EditEmployer_End
  
  If lEmployerIndex <= 0 Then If Not empr.CreateEmployer(p11d32.WorkingDirectory & F_EmployerDetails.TxtBx(2).Text & Rates.FileExtensionCurrent) Then GoTo retryedit
  
  Call empr.OpenEmployer(False)
  
  ben.value(employer_Name_db) = F_EmployerDetails.TxtBx(0).Text
  ben.value(employer_Payeref_db) = F_EmployerDetails.TxtBx(1).Text
  ben.value(employer_Signatory_db) = F_EmployerDetails.TxtBx(3).Text
  ben.value(employer_Contact_db) = F_EmployerDetails.TxtBx(4).Text
  ben.value(employer_contactnumber_db) = F_EmployerDetails.TxtBx(5).Text
  
  ben.value(employer_CT_db) = ChkBoxToBool(F_EmployerDetails.chkbx(0))
  ben.value(employer_CheapTaxableLoansAllNon_db) = ChkBoxToBool(F_EmployerDetails.chkbx(2))
  ben.value(employer_AddressLine1_db) = F_EmployerDetails.TxtBx(6).Text
  ben.value(employer_AddressLine2_db) = F_EmployerDetails.TxtBx(7).Text
  ben.value(employer_AddressLine3_db) = F_EmployerDetails.TxtBx(8).Text
  ben.value(employer_AddressLine4_db) = F_EmployerDetails.TxtBx(9).Text
  ben.value(employer_AddressPostCode_db) = F_EmployerDetails.TxtBx(10).Text
  ben.value(employer_SubmitterRef_db) = F_EmployerDetails.TxtBx(11).Text
  ben.value(employer_SubmitterName_db) = F_EmployerDetails.TxtBx(12).Text
  ben.value(employer_PayeOnlineID_db) = F_EmployerDetails.TxtBx(27).Text
  ben.value(employer_PayeOnlinePwd_db) = F_EmployerDetails.TxtBx(28).Text
  ben.value(employer_PayeOnlineUser_db) = F_EmployerDetails.TxtBx(29).Text
  ben.value(employer_PayeOnlineEmail_db) = F_EmployerDetails.TxtBx(30).Text
  ben.value(employer_AddClass1AAmounts_db) = F_EmployerDetails.TxtBx(14).Text
  ben.value(employer_AddClass1ADescription_db) = F_EmployerDetails.TxtBx(15).Text
  ben.value(employer_deductClass1AAmounts_db) = F_EmployerDetails.TxtBx(16).Text
  ben.value(employer_deductClass1ADescription_db) = F_EmployerDetails.TxtBx(17).Text
  ben.value(employer_EmployeeResponseDate_db) = ScreenToDateVal(F_EmployerDetails.TxtBx(13).Text, STDV_STRING)
  ben.value(employer_TreatEmployeeCarsAsUnderApprovedFPCS_db) = ChkBoxToBool(F_EmployerDetails.chkbx(1))
    
  ben.value(employer_EerDeclSect2Chk1_db) = ChkBoxToBool(F_EmployerDetails.chkbx(3))
  ben.value(employer_EerDeclSect2Chk2_db) = ChkBoxToBool(F_EmployerDetails.chkbx(4))
  ben.value(employer_EerDeclSect2Chk3_db) = ChkBoxToBool(F_EmployerDetails.chkbx(5))
    
  ben.value(employer_IRTaxOffice_db) = F_EmployerDetails.TxtBx(18).Text
  ben.value(employer_IRAddressLine1_db) = F_EmployerDetails.TxtBx(19).Text
  ben.value(employer_IRAddressLine2_db) = F_EmployerDetails.TxtBx(20).Text
  ben.value(employer_IRAddressLine3_db) = F_EmployerDetails.TxtBx(21).Text
  ben.value(employer_IRAddressLine4_db) = F_EmployerDetails.TxtBx(22).Text
  ben.value(employer_IRPostcode_db) = F_EmployerDetails.TxtBx(23).Text
  
  ben.value(employer_IRAddressLine2_db) = F_EmployerDetails.TxtBx(20).Text
  ben.value(employer_IRAddressLine3_db) = F_EmployerDetails.TxtBx(21).Text
  ben.value(employer_IRAddressLine4_db) = F_EmployerDetails.TxtBx(22).Text
  
  ben.value(employer_IRTelephoneNumber_db) = F_EmployerDetails.TxtBx(24).Text
  ben.value(employer_IRYourReference_db) = F_EmployerDetails.TxtBx(25).Text
  ben.value(employer_IRAccountsOfficeReference_db) = F_EmployerDetails.TxtBx(26).Text
    
  ben.value(employer_GroupCode1Alias_db) = F_EmployerDetails.TxtBx(31).Text
  ben.value(employer_GroupCode2Alias_db) = F_EmployerDetails.TxtBx(32).Text
  ben.value(employer_GroupCode3Alias_db) = F_EmployerDetails.TxtBx(33).Text
    
  ben.WriteDB
  
  Set ibf = F_Employers
  
  If lEmployerIndex > 0 Then
    Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ben)
  Else
    EditEmployer = Employers.Add(empr)
    Call ibf.BenefitToListView(ben, EditEmployer)
    Call SelectBenefitByBenefitIndex(ibf, EditEmployer)
  End If
  
  
EditEmployer_End:
  If Not ben Is Nothing Then ben.Kill
  Unload F_EmployerDetails
  Call xReturn("EditEmployer")
  Exit Function
  
EditEmployer_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EditEmployer", "", "Error editing or adding the employer.")
  Resume EditEmployer_End
  Resume
End Function
Public Sub Kill()
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo Kill_Err
  Call xSet("Kill")
  
  Call IniSettings(Ini_Write)
  Call LoadEmployer(Nothing)
  Set CurrentEmployer = Nothing
  
  For i = 1 To Employers.Count
    Set ben = Employers(i)
    If Not ben Is Nothing Then
      Call ben.Kill
      Set ben = Nothing
    End If
  Next i
  Set Help = Nothing
  Set Employers = Nothing
  Set P11DWS = Nothing
  Set ReportPrint = Nothing
  Set Importing = Nothing
  Set Rates = Nothing
  Set MagneticMedia = Nothing
  Set Intranet = Nothing
  Set udm = Nothing
  Set Help = Nothing

Kill_End:
  Call xReturn("Kill")
  Exit Sub

Kill_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Kill", "Kill", "Error in the P11D32 class kill.")
  Resume Kill_End
End Sub
Public Property Get AppYear() As Long
  AppYear = m_AppYear
End Property
Public Property Get Beta() As Boolean
  Beta = False
End Property

Public Property Get AppYearShortLastYear() As String
  AppYearShortLastYear = GetAppYearShort(m_AppYear - 1)
End Property
Public Property Get AppYearShortNextYear() As String
  AppYearShortNextYear = GetAppYearShort(m_AppYear + 1)
End Property

Public Property Get AppYearShort() As String
  AppYearShort = m_AppYearShort
End Property
Public Function TemplateFile(ey As Employer) As String
  Dim ben As IBenefitClass
  Dim sYear As String
  
  On Error GoTo TemplateFile_ERR
  
  Call xSet("TemplateFile")
  
  If ey Is Nothing Then 'assume appyearshort
    sYear = p11d32.AppYearShort
  Else
    Set ben = ey
    sYear = ben.value(employer_Year)
  End If
  
  TemplateFile = FullPath(AppPath) & "Blnk32" & S_DB_FILE_EXTENSION
  
TemplateFile_END:
  Call xReturn("TemplateFile")
  Exit Function
TemplateFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "TemplateFile", "Template File", "Error getting the 'Template File'.")
  Resume TemplateFile_END
End Function

'this is convoluted CAD 2004
Private Function IniPath(ByVal sRootPath As String, Optional ByVal sYear As String = "")
    
  sRootPath = FullPath(FullPath(sRootPath) & app.companyName)
  If Not FileExists(sRootPath, True) Then
    Call xMkdir(sRootPath)
  End If
  
  If (Len(sYear) = 0) Then
    sYear = CStr(p11d32.AppYear)
  End If
  
  sRootPath = FullPath(sRootPath) & "P11D" & sYear
  If Not FileExists(sRootPath, True) Then
    Call xMkdir(sRootPath)
  End If
  sRootPath = FullPath(sRootPath) & "P11D.ini"
  
  IniPath = sRootPath
End Function

Public Property Get IniPathAndFileGlobal() As String
  If Len(m_IniPathAndFileGlobal) = 0 Then
    m_IniPathAndFileGlobal = IniPath(GetSpecialFolderEx(CSIDL_COMMON_APPDATA))
  End If
  IniPathAndFileGlobal = m_IniPathAndFileGlobal
  m_IsGlobal = True
End Property
Public Property Get IniPathAndFile() As String
  Dim s As String
  
  If Len(m_IniPathAndFile) = 0 Then
    If p11d32.AppYear < 2004 Then
      m_IniPathAndFile = (FullPath(AppPath) & "P11D" & CStr((p11d32.AppYear - 1)) & ".ini")
    Else
      m_IniPathAndFile = IniPath(GetSpecialFolderEx(CSIDL_APPDATA), p11d32.AppYear)
    End If
  End If
  
  IniPathAndFile = m_IniPathAndFile
  m_IsGlobal = False

End Property
Public Property Get IniPathAndFileLastYearExists() As Boolean
  Static bFound As Boolean
  
  Dim s As String
  
  If (Not bFound) Then
    s = p11d32.IniPathAndFileLastYear
    If (FileExists(s, False)) Then
      m_IniPathAndFileLastYearExists = True
    End If
  End If
  IniPathAndFileLastYearExists = m_IniPathAndFileLastYearExists
End Property
Public Property Get IniPathAndFileLastYear() As String
  Dim sKey As String
  Dim s As String
  
  If Len(m_IniPathAndFileLastYear) = 0 Then
    If p11d32.AppYear < 2005 Then
      'try and get out of the registry as written in the install
      sKey = "HKEY_LOCAL_MACHINE\Software\DELOITTE\P11D" & p11d32.AppYear - 1 & "\INSTALLDIR"
      On Error GoTo err_key
      sKey = RegGetKeyValue("HKEY_LOCAL_MACHINE\Software\DELOITTE\P11D" & p11d32.AppYear - 1, "INSTALLDIR")
err_key:
      If (Not FileExists(sKey, True)) Then
        sKey = GetSpecialFolderEx(CSIDL_APPDATA) & "Deloitte\P11D" & p11d32.AppYear - 1 & "\"
        If ((Not FileExists(sKey, True))) Then
          sKey = "C:\Program Files\Deloitte\P11D" & p11d32.AppYear - 1
          If ((Not FileExists(sKey, True))) Then
            sKey = "C:\Program Files\Deloitte\P11D" & p11d32.AppYear - 1
            If ((Not FileExists(sKey, True))) Then
              GoTo err_end:
            End If
          End If
        End If
      End If
      s = FullPath(sKey) & "P11D" & p11d32.AppYear - 1 & ".ini"
      If (FileExists(s, False)) Then
        m_IniPathAndFileLastYear = s
      Else
        s = FullPath(sKey) & "P11D" & ".ini"
        If (FileExists(s, False)) Then
          m_IniPathAndFileLastYear = s
        End If
      
      End If
    Else
      m_IniPathAndFileLastYear = IniPath(GetSpecialFolderEx(CSIDL_APPDATA), p11d32.AppYear - 1)
    End If
  End If
err_end:
  IniPathAndFileLastYear = m_IniPathAndFileLastYear
  
End Property


Public Function Initialise() As Boolean
  Dim db As Database, rs As Recordset
  Dim errstring  As String
  Dim TemplateApplication As String
  Dim sTemplateFile As String
  
  On Error GoTo Initialise_ERR
  Call xSet("Initialise")
  
  m_AppYear = GetAppYear
  m_versionCheckResult = Empty
  m_AppYearShort = GetAppYearShort(m_AppYear)
  'P11D32.IniPathAndFile = (FullPath(AppPath) & AppExeName & ".ini")
  'this seems very dody as do not use last years ini file? I will always loos my settings??
  If Not SetSystemMDA Then GoTo Initialise_END
  Username = GetNetUser(False)
  Set Rates = New Rates
  Set Employers = New ObjectList
  Employers.Increment = 300
  Set Importing = New Importing
  Set BringForward = New BringForward
  Set ReportPrint = New ReportPrint
  ' help
  Set Help = New Help
  Help.DefaultLink = S_DEFAULT_HELP_PAGE
  'Call ReportPrint.Initialise
  Set MagneticMedia = New MagneticMedia
  Set PAYEonline = New PAYEonline
  
  Set Intranet = New Intranet
  Call IniSettings(Ini_read)
   
  Call AddF12Menus
  Call AddStatics
  Call InitBenDataLinks
  Call PAYEonline.Initialise
  Set KeyboardHook = Me
  
  
  
  sTemplateFile = TemplateFile(Nothing)
  If FileExists(sTemplateFile) Then
    If IsRunningInIDE Then Call RemoveReadOnlyFile(sTemplateFile) ' change file attribute from read-only to normal
    Set db = InitDB(P11DWS, sTemplateFile, "P11D template", , , False)
    If Not (db Is Nothing) Then
      Set rs = db.OpenRecordset(sql.Queries(SELECT_ID), dbOpenForwardOnly)
      If Not (rs.EOF And rs.BOF) Then
        TemplateApplication = rs.Fields("Application").value
        TemplateAppversion = rs.Fields("Appversion").value
        TemplateFixlevel = rs.Fields("Fixlevel").value
        If StrComp(Rates.TemplateApp, TemplateApplication, vbTextCompare) = 0 Then
          Initialise = True
          GoTo Initialise_END
        End If
      End If
    End If
  End If
  Call Err.Raise(ERR_P11DINIT, "Initialise", "Unable to find a valid template database " & sTemplateFile)

Initialise_END:
  Call xReturn("Initialise")
  Exit Function
  
Initialise_ERR:
  If Len(sTemplateFile) > 0 Then errstring = "Check " & sTemplateFile & " is not being edited"
  Call ErrorMessage(ERR_ERROR, Err, "Initialise", "Initialise", "Unable to initialise the P11D program." & vbCrLf & errstring)
  Resume Initialise_END
  Resume
End Function

Public Sub VersionCheck()
  On Error GoTo err_err
  
  If Not p11d32.VersionCheckEnabled Then GoTo err_end
  
  m_versionCheckResult = S_VERSION_CHECK_UNKNOWN
  m_versionCheckFinished = False
  Set m_versionCheck = New VersionCheck
  Call m_versionCheck.Check
  MDIMain.TimerVersionCheck.Enabled = True
  'cad 2010 enabled this!
err_end:
  Exit Sub
err_err:
  Resume err_end
End Sub
Public Property Get VersionCheckResult() As String
  VersionCheckResult = m_versionCheckResult
End Property
Public Property Let VersionCheckResult(ByVal value As String)
  m_versionCheckResult = value
End Property

Public Property Get VersionCheckFinished() As Boolean
  VersionCheckFinished = m_versionCheckFinished
End Property
Public Property Let VersionCheckFinished(ByVal value As Boolean)
  m_versionCheckFinished = True
End Property
Private Sub AddStatics()

  On Error GoTo AddStatics_Err
  
  Call xSet("AddStatics")

  Call AddStatic("StatusReady", "Ready", , True)
  

AddStatics_End:
  Call xReturn("AddStatics")
  Exit Sub
AddStatics_Err:
  Call Err.Raise(ERR_ADDSTATICS, "AddStatics", "Add Statics", "Error adding static values.")
End Sub
Private Sub AddF12Menus()
  On Error GoTo AddF12Menus_ERR
  
  Call xSet("AddF12Menus")
  
  Set CoreClass.DebugMenuInterface = Me
    
  
  Call AddMenuItem(S_MNU_F12_DEBUG_SQL, S_MNU_F12_DEBUG_SQL, MNU_DATABASE)
  Call AddMenuItem(S_MNU_F12_UPDATE_FIX, S_MNU_F12_UPDATE_FIX, MNU_DATABASE)
  Call AddMenuItem(S_MNU_F12_SPLIT_NAMES, S_MNU_F12_SPLIT_NAMES, MNU_DATABASE)
  Call AddMenuItem(S_MNU_F12_DELETE_ALL_CDBS, S_MNU_F12_DELETE_ALL_CDBS, MNU_DATABASE)
'MP RV TTP#320  Call AddMenuItem(S_MNU_F12_SET_ACTUAL_MILES, S_MNU_F12_SET_ACTUAL_MILES, MNU_DATABASE)
  
  Call AddMenuItem(S_MNU_F12_MAGNETIC_MEDIA_ERROR_LOGGING, S_MNU_F12_MAGNETIC_MEDIA_ERROR_LOGGING, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_MAGNETIC_MEDIA_USER_DATA_SIZE, S_MNU_F12_MAGNETIC_MEDIA_USER_DATA_SIZE, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_UPDATE_LIST_ITEM, S_MNU_F12_UPDATE_LIST_ITEM, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_SHOWEMPLOYERS_FIX_LEVEL, S_MNU_F12_SHOWEMPLOYERS_FIX_LEVEL, MNU_APPLICATION)
  
  'Call AddMenuItem(S_MNU_F12_ENABLE_EMAIL_REPORTS, S_MNU_F12_ENABLE_EMAIL_REPORTS, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_EMAIL_DEBUG, S_MNU_F12_EMAIL_DEBUG, MNU_APPLICATION)
  
  Call AddMenuItem(S_MNU_F12_DISPLAY_INVALIDFIELDS, S_MNU_F12_DISPLAY_INVALIDFIELDS, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_KILL_BENEFITS, S_MNU_F12_KILL_BENEFITS, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_ENTER_SERIAL_NUMBER, S_MNU_F12_ENTER_SERIAL_NUMBER, MNU_APPLICATION)
  
  Call AddMenuItem(S_MNU_F12_PAYE_ONLINE_SHOW_EXTRA_SUBMISSION_PROPERTIES_MENU, S_MNU_F12_PAYE_ONLINE_SHOW_EXTRA_SUBMISSION_PROPERTIES_MENU, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_VIEW_PROCEED_BUTTON_IF_ERRORS, S_MNU_F12_VIEW_PROCEED_BUTTON_IF_ERRORS, MNU_APPLICATION)
  
  
  
  
  Call AddMenuItem(S_MNU_F12_PAYE_REFERENCE_ANY_FORMAT, S_MNU_F12_PAYE_REFERENCE_ANY_FORMAT, MNU_APPLICATION)
  
  Call AddMenuItem(S_MNU_F12_REREAD_CONTEXT_SENSITIVE_HELP_LINKS, S_MNU_F12_REREAD_CONTEXT_SENSITIVE_HELP_LINKS, MNU_APPLICATION)
  
  
  Call AddMenuItem(S_MNU_F12_DATA_TYPE_LIST_VIEW_SORTING, S_MNU_F12_DATA_TYPE_LIST_VIEW_SORTING, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_F12_SETTINGS_FORM, S_MNU_F12_F12_SETTINGS_FORM, MNU_APPLICATION)


  Call AddMenuItem(S_MNU_F12_SORT_OTHER_ALPHABETICALLY, S_MNU_F12_SORT_OTHER_ALPHABETICALLY, MNU_APPLICATION)
  
  Call AddMenuItem(S_MNU_F12_QA_MANAGEMENT_REPORTS, S_MNU_F12_QA_MANAGEMENT_REPORTS, MNU_APPLICATION)
  Call AddMenuItem(S_MNU_F12_VERSION_CHECK, S_MNU_F12_VERSION_CHECK, MNU_APPLICATION)
  
  
  ' add HelpLink Checker menu item - only if no_serial licence type
  If LicenceType = LT_UNLICENSED Then
    'Call AddMenuItem(S_MNU_F12_CHECK_HELPLINKS, S_MNU_F12_CHECK_HELPLINKS, MNU_APPLICATION)
  End If
  
AddF12Menus_END:
  Call xReturn("AddF12Menus")
  Exit Sub
AddF12Menus_ERR:
  
End Sub
Public Function CurrentEmployeeIsNothing() As Boolean
  CurrentEmployeeIsNothing = True
  If Not CurrentEmployer Is Nothing Then
    CurrentEmployeeIsNothing = CurrentEmployer.CurrentEmployee Is Nothing
  End If
End Function

Private Sub IEnumEmployers_Count(ByVal Count As Long)
  Call PrgStartCaption(Count, "Searching for Employers...", "", Indicator.ValueOfMax)
  F_Employers.LB.listitems.Clear
End Sub

Private Sub IEnumEmployers_CurrentFile(ByVal sCurrentFile As String)
  Call PrgStepCaption("Checking file... " & sCurrentFile)
End Sub

Private Sub IEnumEmployers_Employer(ey As Object)
  Dim empr As Employer
  
  Set empr = ey
  If empr Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "IEnumEmployers_Employer", "Employer is nothing")
  Call Employers.Add(empr)
End Sub

Private Sub IKeyBoardHook_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim ibf As IBenefitForm2
  Dim bInFunc As Boolean
  
  On Error GoTo IKeyBoardHook_ERR
  
  Call xSet("IKeyBoardHook")
  
  If bInFunc Then GoTo IKeyBoardHook_END_PROC
  bInFunc = True
  'repeated calls to IsBenefitForm to reduce overhead on key down hook
  
  Select Case KeyCode
    Case vbKeyPageDown
      If Shift And vbCtrlMask Then
        If IsKeyBoardHookForm(CurrentForm) Then
          Call MoveEmployee(False)
          KeyCode = 0
        End If
      End If
    Case vbKeyPageUp
      If Shift And vbCtrlMask Then
        If IsKeyBoardHookForm(CurrentForm) Then
          Call MoveEmployee(True)
          KeyCode = 0
        End If
      End If
    Case vbKeyDelete
      If Shift And vbCtrlMask Then
        If IsKeyBoardHookForm(CurrentForm) Then
          Set ibf = CurrentForm
          If Not ibf.lv.SelectedItem Is Nothing Then
            Call ToolBarButton(TBR_REMOVE_BENEFIT, ibf.lv.SelectedItem.Tag)
            KeyCode = 0
          End If
        End If
      End If
    Case vbKeyInsert
      If Shift And vbCtrlMask Then
        If IsKeyBoardHookForm(CurrentForm) Then
          Call ToolBarButton(TBR_ADD_BENEFIT, 0)
          KeyCode = 0
        End If
      End If
    Case vbKeyX, vbKeyC, vbKeyV
      If ((vbCtrlMask And Shift) = vbCtrlMask) And ((Shift And vbShiftMask) = vbShiftMask) Then
        If IsBenefitFormReal(CurrentForm) Then
          Call LVKeyDown(KeyCode, Shift)
          KeyCode = 0
        End If
      End If
    Case vbKeyF
      If (Shift And vbCtrlMask) Then
        If CurrentForm Is F_Employees Then
          Call F_Employees.FoundEmployee(KeyCode, 0) 'EK 1/04 TTP#12
          KeyCode = 0
        ElseIf IsBenefitFormReal(CurrentForm) Then
          Call GotoScreen
          KeyCode = 0
        End If
      End If
    Case vbKeyF1
      Call p11d32.Help.ShowHelp("")
      KeyCode = 0
    Case vbKeyF11
      ' check if Ctrl key was pressed when this function was called
      If CBool(GetAsyncKeyState(VK_CONTROL)) = True And CBool(GetAsyncKeyState(VK_SHIFT)) = True Then
        If p11d32.LicenceType = LT_DEMO Or p11d32.LicenceType = LT_UNLICENSED Then
          Call p11d32.Help.ShiftControlF11
          KeyCode = 0
        End If
      End If
      
  End Select
  
IKeyBoardHook_END_PROC:
  bInFunc = False
IKeyBoardHook_END:
  Call xReturn("IKeyBoardHook")
  Exit Sub
IKeyBoardHook_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IKeyBoardHook", "IKeyBoard Hook", "Error processing the core keyboard hooks to navigate benefits.")
  bInFunc = False
  Resume IKeyBoardHook_END
End Sub


Public Function DirList(IRW As INI_READ_WRITE)
  Call IniSettings(IRW)

End Function

Public Function SavedFilesList(IRW As INI_READ_WRITE)
  Call IniSettings(IRW)
End Function

'MP RV TTP#320 this proc only used with F12 menu which no longer used - so have commented this proc
'Public Sub SetActualMiles()
' On Error GoTo SetActualMiles_ERR
'
'  Dim empr As Employer
'
'  Call xSet("SetActualMiles")
'
'  If Not IfNotCurrentEmployerThenLoad(empr) Then GoTo SetActualMiles_END
'  If empr.db Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "SetActualMiles", "Database is nothing.")
'
'  empr.db.Execute (sql.Queries(SET_ACTUAL_MILES))
'
'SetActualMiles_END:
'  Call xReturn("SetActualMiles")
'  Exit Sub
'SetActualMiles_ERR:
'  Call ErrorMessage(ERR_ERROR, Err, "SetActualMiles", "Set Actual Miles", "Error setting the actual miles to true.")
'  Resume SetActualMiles_END
'  Resume
'End Sub
Public Property Get GlobalCopySerialNumber() As String
  GlobalCopySerialNumber = m_GlobalCopySerialNumber
End Property
Public Property Let GlobalCopySerialNumber(sNewValue As String)
  m_GlobalCopySerialNumber = sNewValue
End Property
Public Property Get SerialNumber() As String
  SerialNumber = m_SerialNumber
End Property
Public Property Let SerialNumber(sNewValue As String)
  If ValidateSerialNumber(sNewValue) Then
     m_SerialNumber = sNewValue
  ElseIf m_IsGlobal Then
    If Not IsFormLoaded("F_SerialNumber") Then
      EnterSerialNumber
    Else
      F_SerialNumber.Visible = True
    End If
  End If
End Property
Public Property Get NoSerial() As Boolean
  NoSerial = GetCmdParam("/no_serial") Or True
End Property
Public Property Get JuneRelease() As Boolean
  If IsRunningInIDE Then
    JuneRelease = False
    Exit Property
  End If
  JuneRelease = False
End Property
Private Function ValidateSerialNumber(sSerialNumber As String) As Boolean
  Dim bOverride
  On Error GoTo ValidateSerialNumber_Err
  Call xSet("ValidateSerialNumber")
  
  If p11d32.NoSerial Then
    GoTo DEMO
  End If
  
  If StrComp(sSerialNumber, S_SERIAL_NUMBER_STANDARD, vbTextCompare) = 0 Then
     ValidateSerialNumber = True
     ReportPrint.AllReports = True
     ReportPrint.AbacusUDM = False
     If IsFormLoaded("MDIMain") Then
        MDIMain.mnuIntranet.Enabled = False
     End If
  ElseIf StrComp(sSerialNumber, S_SERIAL_NUMBER_INTRANET, vbTextCompare) = 0 Then
     ValidateSerialNumber = True
     ReportPrint.AllReports = True
     ReportPrint.AbacusUDM = True
     If IsFormLoaded("MDIMain") Then  'Enable on upgrade
       MDIMain.mnuIntranet.Enabled = True
     End If
  ElseIf StrComp(sSerialNumber, S_SERIAL_NUMBER_SHORT, vbTextCompare) = 0 Then
     ValidateSerialNumber = True
     ReportPrint.AllReports = False
     ReportPrint.AbacusUDM = False
     If IsFormLoaded("MDIMain") Then
        MDIMain.mnuIntranet.Enabled = False
        MDIMain.mnuFileErrorLogsMagneticMedia.Enabled = False
     End If
  ElseIf StrComp(sSerialNumber, S_SERIAL_NUMBER_DEMO, vbTextCompare) = 0 Then
DEMO:
     ValidateSerialNumber = True
     ReportPrint.AllReports = True
     ReportPrint.AbacusUDM = True
     If IsFormLoaded("MDIMain") Then
        MDIMain.mnuIntranet.Enabled = True
     End If
     
     sSerialNumber = S_SERIAL_NUMBER_DEMO
     
  Else
     ValidateSerialNumber = False
  End If

ValidateSerialNumber_End:
  Call xReturn("ValidateSerialNumber")
  Exit Function

ValidateSerialNumber_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateSerialNumber", "Error in ValidateSerialNumber", "Undefined error.")
  Resume ValidateSerialNumber_End
End Function
Public Property Get LicenceType() As LICENCE_TYPE
  
  If Len(p11d32.SerialNumber) = 0 Then
    LicenceType = LT_DEMO
  Else
    LicenceType = CheckLicenceType(p11d32.SerialNumber)
  End If
End Property
Private Function CheckLicenceType(sSerialNumber As String) As LICENCE_TYPE
  On Error GoTo CheckLicenceType_Err
  Call xSet("CheckLicenceType")
    
  If StrComp(sSerialNumber, S_SERIAL_NUMBER_STANDARD, vbTextCompare) = 0 Then
     CheckLicenceType = LT_STANDARD
  ElseIf StrComp(sSerialNumber, S_SERIAL_NUMBER_INTRANET, vbTextCompare) = 0 Then
     CheckLicenceType = LT_INTRANET
  ElseIf StrComp(sSerialNumber, S_SERIAL_NUMBER_SHORT, vbTextCompare) = 0 Then
     CheckLicenceType = LT_SHORT
  ElseIf StrComp(sSerialNumber, S_SERIAL_NUMBER_DEMO, vbTextCompare) = 0 Then
     CheckLicenceType = LT_DEMO
  Else
     CheckLicenceType = LT_UNLICENSED
  End If

CheckLicenceType_End:
  Call xReturn("CheckLicenceType")
  Exit Function

CheckLicenceType_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckLicenceType", "Error in CheckLicenceType", "Undefined error.")
  CheckLicenceType = LT_UNLICENSED
  Resume CheckLicenceType_End
End Function

Public Property Get WorkingDirectory() As String
  WorkingDirectory = m_WorkingDirectory
End Property
Public Property Let WorkingDirectory(ByVal sNewValue As String)
  m_WorkingDirectory = sNewValue
End Property




