VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MagneticMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IErrorFilter

Public Enum MMDATE_COMPARE
  MMDC_TAX_YEAR_START = 1
  MMDC_TAX_YEAR_END
  MMDC_NO_COMPARE
End Enum

Public Enum MM_DATA_FORMAT
  MM_DF_ANY
  [MM_DF_FIRSTITEM] = MM_DF_ANY
  MM_DF_35FLOPPY
  [MM_DF_LASTITEM] = MM_DF_35FLOPPY
End Enum

Public Enum MM_STANDARD_FIELD_SIZES
  MM_SFS_bool = 1
  MM_SFS_DATE = 8
  MM_SFS_DATA = 9
  MM_SFS_DESCRIPTION = 30
End Enum

Public SubReturn As Long
Public SubReturnOf As Long
Public DataFormat As MM_DATA_FORMAT
Public SubmitterRef As String
Public SubmitterName As String
Public ErrorCount As Long
Public LastPathAndFileCreated As String
Public RecordViewID As String
Public ErrorLogging As Boolean
Public UserDataSize As Long


Private m_TS As TextStream
Private m_lBytesWritten As Long

Private m_rsError As Recordset
Private m_db As Database
Private m_ws As Workspace
Private m_DataFormatSpace As Long
Private m_OutputDirectory As String
Private Const DEFAULT_DISK_SIZE As Single = (1024! * 64! * 1024!)  ' 64 MB

Private Sub CloseErrorLogDB()
  On Error GoTo CloseErrorLogDB_ERR
  
  Call xSet("CloseErrorLogDB")
  
  Set m_rsError = Nothing
  Set m_db = Nothing
  Set m_ws = Nothing
  
CloseErrorLogDB_END:
  Call xReturn("CloseErrorLogDB")
  Exit Sub
CloseErrorLogDB_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CloseErrorLogDB", "Close Error Log", "Error closing the magnetic media error log.")
  Resume CloseErrorLogDB_END

End Sub
Private Sub OpenErrorLogDB()
  Set m_ws = DBEngine.CreateWorkspace("ErrorWS", p11d32.DBUserName, p11d32.DBPassword)
  Set m_db = p11d32.PDDB(m_ws)
End Sub
Public Property Get MinDiskFreeSpace() As Long
  MinDiskFreeSpace = 896 '4 * size on largest record see MM manual
End Property
Private Sub GetDataFormatSpace()
  Dim i As MM_DATA_FORMAT
  
  If UsingUserDataSize Then
    m_DataFormatSpace = UserDataSize
    'see f12 debug
  Else
    For i = [MM_DF_FIRSTITEM] + 1 To [MM_DF_LASTITEM]
      Select Case i
        Case MM_DF_35FLOPPY
          m_DataFormatSpace = 1400000 - MinDiskFreeSpace
      End Select
    Next
  End If
  
  If m_DataFormatSpace <= MinDiskFreeSpace Then
    Call Err.Raise(ERR_INVALID_DATA_FORMAT_SPACE, "GetDataFormatSpace", "The data format space is below the minimum allowed of " & MinDiskFreeSpace & ".")
  End If
  
End Sub

Private Function Year() As String
  Year = Format$(p11d32.Rates.value(TaxYearEnd), "YYYY")
End Function

Public Function UpdateListViewItem(li As ListItem, ben As IBenefitClass, Optional ByVal BenefitIndex As Long = 0) As Boolean
  
  On Error GoTo UpdateListViewItem_ERR 'JN
  
  Call xSet("UpdateListViewItem") 'JN
  
  Dim emp As Employer
  
  If li Is Nothing Then Call Err.Raise(ERR_ERROR, "UpdateListViewItem", "The list item is nothing")
  If ben Is Nothing Then Call Err.Raise(ERR_ERROR, "UpdateListViewItem", "The beneift is nothing")
  
  If BenefitIndex > 0 Then li.Tag = BenefitIndex
  
  Set emp = ben
  
  If Not emp.MagneticMediaValid(False) Then
    li.SmallIcon = IMG_EXCLAMATION
  Else
    li.SmallIcon = IMG_OK
  End If
  
  li.Text = ben.Name
  li.SubItems(1) = ben.value(employer_Payeref_db)
  li.SubItems(2) = IIf(emp.PayeRefValid, "Yes", "No")
  li.SubItems(3) = ben.value(employer_EmployeesCount)
  li.SubItems(4) = ben.value(employer_SubmitterName_db)
  li.SubItems(5) = ben.value(employer_SubmitterRef_db)
  
  UpdateListViewItem = True
  
UpdateListViewItem_END: 'JN
  Call xReturn("UpdateListViewItem")
  Exit Function
  
UpdateListViewItem_ERR: 'JN
  Call ErrorMessage(ERR_ERROR, Err, "UpdateListViewItem", "Update List View Item", "Error Updating List Item")
  Resume UpdateListViewItem_END
End Function
Public Sub EmployersToListView()
  Dim i As Long
  Dim ben As IBenefitClass
  Dim li As ListItem

  On Error GoTo EmployersToListView_ERR
  
  Call xSet("EmployersToListView")

  For i = 1 To p11d32.Employers.Count
    Set ben = p11d32.Employers(i)
    If Not ben Is Nothing Then
      Set li = F_MM.lvMMEmployers.listitems.Add()
      Call UpdateListViewItem(li, ben, i)
    End If
  Next
  
  Call F_MM.lvMMEmployers_ItemCheck(Nothing)
  
EmployersToListView_END:
  Call xReturn("EmployersToListView")
  Exit Sub
EmployersToListView_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployersToListView", "Employers To List View", "Error placing the current employers to the Magnetic Media list view.")
  Resume EmployersToListView_END
  Resume
End Sub
Private Function EmployersSelected(Employers As ObjectList) As Long
  Dim i As Long
  Dim ben As IBenefitClass
  
  Set Employers = New ObjectList
  For i = 1 To p11d32.Employers.Count
    Set ben = p11d32.Employers(i)
    If Not ben Is Nothing Then
      If ben.value(employer_MagneticMediaSelected) Then
        If i = 1 Then
          SubmitterName = ben.value(employer_SubmitterName_db)
          SubmitterRef = ben.value(employer_SubmitterRef_db)
        End If
        Call Employers.Add(ben)
      End If
    End If
  Next
    
  EmployersSelected = Employers.Count
End Function
Private Function ValidEmployees(employees As ObjectList, ey As Employer) As Long
  Dim i As Long
  Dim ee As Employee
  
  Set employees = New ObjectList
  
  For i = 1 To ey.employees.Count
    Set ee = ey.employees(i)
    If Not ee Is Nothing Then
      Call employees.Add(ee)
    End If
  Next
    
  ValidEmployees = employees.Count

End Function

Private Function OpenFile(lVolNumber As Long, ByVal lblFile As Label) As Boolean
  Dim fs As FileSystemObject
  Dim sFileName As String
  
  On Error GoTo OpenFile_ERR
  
  Set fs = New FileSystemObject
  
  lVolNumber = lVolNumber + 1
  sFileName = PathAndFile(lVolNumber)
  If FileExists(sFileName) Then
    If MsgBox("The file " & sFileName & " already exists, overwrite?", vbOKCancel, "OpenFile") = vbCancel Then
      GoTo OpenFile_END
    Else
      Set m_TS = fs.CreateTextFile(sFileName, True)
    End If
  Else
    Set m_TS = fs.CreateTextFile(sFileName)
  End If
  LastPathAndFileCreated = sFileName
  lblFile = "Writing file: " & sFileName
  DoEvents
  OpenFile = True
  
OpenFile_END:
  Exit Function
  
OpenFile_ERR:
  If Len(sFileName) = 0 Then Err.Raise ERR_FILE_OPEN, "Open file", "Unable to open the file, could not create FileSystemObject" 'apf
  Err.Raise ERR_FILE_OPEN, "Open file", "Unable to open the file [" & sFileName & "] access may be restricted. " 'APF
  Resume
End Function
Private Function UsingUserDataSize() As Boolean
  If UserDataSize > 0 Then UsingUserDataSize = True
End Function
Private Function OutData(sData As String, lVolumeNumber As Long, lVolumeEmployerCount As Long, lVolumeEmployeeCount As Long, lblFile As Label, Optional bWritingEmployee As Boolean = False, Optional bWritingEmployer As Boolean = False) As Boolean
  Dim sngFreeSpace As Single
  Dim sngTotalSpace As Single
  
  On Error GoTo OutData_ERR
  
  If Not CoreClass.GetDiskSpace(OutputDirectory, sngTotalSpace, sngFreeSpace) Then
    sngTotalSpace = DEFAULT_DISK_SIZE
    sngFreeSpace = DEFAULT_DISK_SIZE
  End If
  
  m_lBytesWritten = m_lBytesWritten + Len(sData)
  
  If sngFreeSpace <= MinDiskFreeSpace Or ((DataFormat > MM_DF_ANY Or UsingUserDataSize) And m_lBytesWritten > m_DataFormatSpace) Then
    If Not m_TS Is Nothing Then
      'watch for changing volumes
      If bWritingEmployee Then lVolumeEmployeeCount = lVolumeEmployeeCount - 1
      If bWritingEmployer Then lVolumeEmployerCount = lVolumeEmployerCount - 1
      
      m_TS.Write (Type4Record(lVolumeEmployerCount, lVolumeEmployeeCount))
      OutData = True
      'watch for changing volumes
      If bWritingEmployee Then lVolumeEmployeeCount = 1 Else lVolumeEmployeeCount = 0
      If bWritingEmployer Then lVolumeEmployerCount = 1 Else lVolumeEmployerCount = 0
      Call CloseFile
    End If
    Do While sngFreeSpace <= MinDiskFreeSpace
      If MsgBox("There is not enough disk space in " & vbCrLf & vbCrLf & OutputDirectory & vbCrLf & vbCrLf & " Please free some space or change the disk.", vbOKCancel Or vbInformation) = vbCancel Then
        Call Err.Raise(ERR_NO_FREE_SPACE, "OutData", "Write of data aborted as not enough disk space.")
      End If
      If Not CoreClass.GetDiskSpace(OutputDirectory, sngTotalSpace, sngFreeSpace) Then
        'cannot determine disk space free
        sngTotalSpace = DEFAULT_DISK_SIZE
        sngFreeSpace = DEFAULT_DISK_SIZE
      End If
    Loop
  End If
    
  If (m_TS Is Nothing) Then
    If OpenFile(lVolumeNumber, lblFile) Then
      m_lBytesWritten = 0
      Call m_TS.Write(MMStr(Type0Record(lVolumeNumber)))
    End If
  End If
    
  
  m_TS.Write (sData)
  
  
OutData_END:
  Exit Function
OutData_ERR:
  Call Err.Raise(ERR_WRITING_DATA, "OutData", IIf(IsClientError(Err.Number), "Error writing the magnetic media data." & Err.Description, "Fatal error writing the magnetic media data."))
End Function
Private Sub CloseFile()
  If Not m_TS Is Nothing Then
    m_TS.Close
    Set m_TS = Nothing
  End If
End Sub


Private Function CheckBen(ey As IBenefitClass, ee As Employee, ben As IBenefitClass) As Boolean
  Dim sHMITSectionString As String
  Dim sBenefitFormCaption As String
  
  On Error GoTo CheckBen_ERR
  
  Call xSet("CheckBen")
      
  sHMITSectionString = p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION_STRING)
  sBenefitFormCaption = p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_FORM_CAPTION)
        
  If ben.value(ITEM_BENEFIT_REPORTABLE) Then
    If AllMMRequiredFieldsPresent(ey, ee, ben) Then
      If Not IsNumeric(ben.value(ITEM_BENEFIT)) Then
        Call Err.Raise(ERR_BEN_INCORRECT, "CheckBen", "Benefit is in error.")
      End If
      CheckBen = True
    End If
  Else
    If Not IsNumeric(ben.value(ITEM_BENEFIT)) Then
      Call Err.Raise(ERR_BEN_INCORRECT, "CheckBen", "Benefit is in error.")
    End If
  End If
  
CheckBen_END:
  Call xReturn("CheckBen")
  Exit Function
CheckBen_ERR:
  Call ErrorMessagePush(Err)
  If Not ben Is Nothing Then
    If (Err.Number = ERR_BEN_NOT_REPORTABLE) Or (Err.Number = ERR_BEN_INCORRECT) Then
      Call ErrorMessagePop(ERR_ERROR, Err, "CheckBen", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name, sBenefitFormCaption), "")
    Else
      Call ErrorMessagePop(ERR_ERROR, Err, "CheckBen", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name, sBenefitFormCaption), "Unable to sum benefits.")
    End If
  Else
    Call ErrorMessagePop(ERR_ERROR, Err, "CheckBen", FilterMessageTitle(ey.Name, ee.FullName, , "Unknown"), "Unable to sum benefits.")
  End If
  Resume CheckBen_END
  Resume
End Function
Private Function AllMMRequiredFieldsPresent(ey As IBenefitClass, ee As Employee, ben As IBenefitClass) As Boolean
  Dim i As Long
  Dim bc As BEN_CLASS
  Dim sHMITSectionString As String
  Dim sBenefitFormCaption As String
  Dim v As Variant
  
  On Error GoTo AllMMRequiredFieldsPresent_ERR
  
  Call xSet("AllMMRequiredFieldsPresent")
  
  bc = ben.BenefitClass
  sBenefitFormCaption = p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION)
  sHMITSectionString = p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING)
  For i = 1 To p11d32.Rates.BenClassTo(bc, BCT_BENITEMS_LAST_ITEM) - 1
    If Not p11d32.BenDataLinkMMFieldRequiredNotFilled(bc, i, ben.value(i)) Then
      Call Err.Raise(ERR_MMFIELD_REQUIRED_NOT_FILLED, "AllMMRequiredFieldsPresent", "Required field not entered or incorrect: benitem = " & i & ". ")
    End If
  Next
  
  AllMMRequiredFieldsPresent = True
  
AllMMRequiredFieldsPresent_END:
  Call xReturn("AllMMRequiredFieldsPresent")
  Exit Function
AllMMRequiredFieldsPresent_ERR:
  If Err.Number = ERR_MMFIELD_REQUIRED_NOT_FILLED Then
    Call ErrorMessage(ERR_ERROR, Err, "AllMMRequiredFieldsPresent", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name, sBenefitFormCaption), Err.Description)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "AllMMRequiredFieldsPresent", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name, sBenefitFormCaption), "Not all required fields are present in benefit.")
  End If
  Resume AllMMRequiredFieldsPresent_END
  Resume
End Function
Private Function Type0Record(lVolumeNumber As Long) As String
  On Error GoTo Type0Record_ERR

  Call xSet("Type0Record")


  Type0Record = RecordType(MM_REC_FILE_OPEN) & PadField("P11D", 35, TYPE_STR) & Year & PadField("DT", 3, TYPE_STR) & _
                PadField("", 7, TYPE_STR) & PadField(CLng(SubmitterRef), 8, TYPE_STR) & _
                PadField(SubmitterName, 35, TYPE_STR) & PadField(CStr(SubReturn), 2, TYPE_LONG) & _
                PadField(CStr(SubReturnOf), 2, TYPE_LONG) & PadField(CStr(lVolumeNumber), 2, TYPE_LONG) & vbCrLf
                
Type0Record_END:
  Call xReturn("Type0Record")
  Exit Function
Type0Record_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Type0Record", FilterMessageTitle(), "Open volume type 0 record, number=" & lVolumeNumber)
  Resume Type0Record_END
End Function

Private Function Type1Record(ey As Employer) As String
  Dim ben As IBenefitClass
  On Error GoTo Type1Record_ERR

  Call xSet("Type1Record")
  
  Set ben = ey
  Type1Record = RecordType(MM_REC_EMPLOYER_OPEN) & PadField(ey.TaxOfficeNumeral, 3, TYPE_LONG) & _
                PadField(ey.EmployerRefNo, 10, TYPE_STR) & _
                PadField("", 35, TYPE_STR) & PadField(ben.Name, 35, TYPE_STR) & _
                PadField(ben.value(employer_AddressLine1_db), 28, TYPE_STR) & PadField(ben.value(employer_AddressLine2_db), 28, TYPE_STR) & _
                PadField(ben.value(employer_AddressLine3_db), 28, TYPE_STR) & PadField(ben.value(employer_AddressLine4_db), 18, TYPE_STR) & _
                PadField(ben.value(employer_AddressPostCode_db), 8, TYPE_STR) & vbCrLf
                
Type1Record_END:
  Call xReturn("Type1Record")
  Exit Function
Type1Record_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Type1Record", FilterMessageTitle(ben.Name), "Employer type 1 record")
  Resume Type1Record_END
End Function

Private Function Type20Record(ee As Employee, ey As IBenefitClass, lEmployeesWithInvaidNI As Long) As String
  Dim ben As IBenefitClass
  Dim sNINumber As String
  Dim lError As Long
    
  On Error GoTo Type20Record_ERR
  
  Call xSet("Type20Record")
  
  Set ben = ee
  
  If Not ee.NINumberValid Then
    lEmployeesWithInvaidNI = lEmployeesWithInvaidNI + 1
    Call Err.Raise(ERR_NI_INVALID_BUT_ADDED, "Type20Record", "NI number '" & ben.value(ee_NINumber_db) & "' invalid but employee added.")
    'leave NI blank as IR will not acccept
  Else
    sNINumber = UCASE$(ben.value(ee_NINumber_db))
  End If
  
    'validate employee number, if changed then
  If Not ValidForEDI(ee.PersonnelNumber) Then Call Err.Raise(ERR_PERSONEL_NUMBER_CHANGED, "Type20Record", "Personnel number '" & ee.PersonnelNumber & "' contains invalid characters.")

  Type20Record = RecordType(MM_REC_EMPLOYEE) & PadFieldHelper(BC_EMPLOYEE, ee_PersonnelNumber_db, ben.value(ee_PersonnelNumber_db)) & _
                PadFieldHelper(BC_EMPLOYEE, ee_Surname_db, ben.value(ee_Surname_db)) & PadFieldHelper(BC_EMPLOYEE, ee_Firstname_db, ben.value(ee_Firstname_db)) & _
                PadField("", 7, TYPE_STR) & PadField(IIf(ben.value(ee_Director_db), "Y", "N"), 1, TYPE_STR) & _
                PadFieldHelper(BC_EMPLOYEE, ee_NINumber_db, sNINumber) & vbCrLf
  
Type20Record_END:
  Call xReturn("Type20Record")
  Exit Function
Type20Record_ERR:
  lError = Err.Number
  If (lError = ERR_NI_INVALID_BUT_ADDED) Or (lError = ERR_PERSONEL_NUMBER_CHANGED) Then
    Call ErrorMessage(ERR_ERROR, Err, "Type20Record", FilterMessageTitle(ey.Name, ee.FullName), Err.Description)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "Type20Record", FilterMessageTitle(ey.Name, ee.FullName), "Employee record.")
  End If
  If (lError = ERR_NI_INVALID_BUT_ADDED) Or (lError = ERR_PERSONEL_NUMBER_CHANGED) Then Resume Next
  
  Resume Type20Record_END
  Resume
End Function

Private Function Type3Record(lTotalEmployerCount As Long, lTotalEmployeeCount As Long) As String
  On Error GoTo Type3Record_END
  
  Call xSet("Type3Record")
  
    Type3Record = RecordType(MM_REC_EMPLOYER_CLOSE) & PadField(lTotalEmployerCount, 8, TYPE_LONG) & PadField(lTotalEmployeeCount, 8, TYPE_LONG) & vbCrLf
    
Type3Record_END:
  Call xReturn("Type3Record")
  Exit Function
Type3Record_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Type3Record", FilterMessageTitle(), "Closing overall type 3 record.")
  Resume Type3Record_END
End Function
Private Function Type4Record(lVolumeEmployerCount, lVolumeEmployeeCount As Long) As String
  On Error GoTo Type4Record_END
  
  Call xSet("Type4Record")
  
    Type4Record = RecordType(MM_REC_FILE_CLOSE) & PadField(lVolumeEmployerCount, 8, TYPE_LONG) & PadField(lVolumeEmployeeCount, 8, TYPE_LONG) & vbCrLf
    
Type4Record_END:
  Call xReturn("Type4Record")
  Exit Function
Type4Record_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Type4Record", FilterMessageTitle(), "Closing volume type 4 record.")
  Resume Type4Record_END
End Function
Public Sub CreateMagneticMediaFiles(ByVal prgEmployee As TCSProgressBar, ByVal lblFile As Label)
  Dim ey As Employer, ee As Employee
  Dim ben As IBenefitClass
  Dim Employers As ObjectList
  Dim bMMSelected As Boolean
  Dim bAnyEmployeeBenefitRecords As Boolean
  Dim i As Long, j As Long, lBytesWritten As Long
  Dim lVolumeNumber As Long, lEmployeesWritten As Long, lEmployersWritten As Long
  Dim lVolumeEmployerCount As Long, lVolumeEmployeeCount As Long, lEmployeesWithInvalidNI As Long
  Dim sBenefits As String, sEmployeeRec As String
  
  On Error GoTo CreateMagneticMediaFiles_ERR
  
  Call xSet("CreateMagneticMediaFiles")
  
  ErrorCount = 0
  
  Call OpenErrorLogDB
  Call GetDataFormatSpace
  
  prgEmployee.Visible = True
  prgEmployee.value = prgEmployee.Min
  prgEmployee.Indicator = None
  prgEmployee.Caption = ""
  
  Call EmployersSelected(Employers)
  
  If Employers.Count > 0 Then
    If ErrorLogging Then
      Set ErrorFilter = Me
    Else
      ErrorLogging = True
    End If
    Call m_db.Execute(sql.Queries(DELETE_MAGNETIC_MEDIA_ERRORS))
    Set m_rsError = m_db.OpenRecordset(sql.Queries(SELECT_MAGNETIC_MEDIA_ERRORS), dbOpenDynaset)
  End If
  
  Call p11d32.LoadEmployer(Nothing, False)
  For i = 1 To Employers.Count
    Set ey = Employers(i)
    Call p11d32.LoadEmployer(ey, False)
    Set ben = ey
    bAnyEmployeeBenefitRecords = False
            
    If ben.value(employer_EmployeesCount) = 0 Then GoTo NEXT_EMPLOYER
    
    For j = 1 To ey.employees.Count
      Set ee = ey.employees(j)
      Call ee.LoadBenefits(TBL_ALLBENEFITS, False)
      If ee.IterateBenefits(AnyReportable) Then
        bAnyEmployeeBenefitRecords = True
        Exit For
      End If
    Next
    
    If Not bAnyEmployeeBenefitRecords Then
      GoTo NEXT_EMPLOYER
    End If
    
    Call CheckDuplicateNINumbers(ey)
    
    lVolumeEmployerCount = lVolumeEmployerCount + 1
    lEmployersWritten = lEmployersWritten + 1
    
    
    Call OutData(Type1Record(ey), lVolumeNumber, lVolumeEmployerCount, lVolumeEmployeeCount, lblFile, , True)
    
    prgEmployee.Max = ey.employees.Count
    prgEmployee.value = 0
    prgEmployee.Indicator = ValueOfMax
    
    For j = 1 To ey.employees.Count
      Set ee = ey.employees(j)
      Call prgEmployee.StepCaption("Analysing employee: " & ee.FullName)
      If ee Is Nothing Then GoTo NEXT_EMPLOYEE
      Call ee.LoadBenefits(TBL_ALLBENEFITS, False)
      Call ee.CalculateBenefits
      
      sBenefits = GetEmployeesBenefits(ey, ee)
      If Len(sBenefits) > 0 Then
        If InStr(1, ee.PersonnelNumber, " ", vbTextCompare) Then
          Call Err.Raise(ERR_SPACE_IN_PAYE_NUM, , "Space found in employee PAYE reference " & ee.PersonnelNumber)
        End If
        sEmployeeRec = Type20Record(ee, ey, lEmployeesWithInvalidNI)
        If Len(sEmployeeRec) > 0 Then
          lEmployeesWritten = lEmployeesWritten + 1
          lVolumeEmployeeCount = lVolumeEmployeeCount + 1
          Call OutData(sEmployeeRec & sBenefits, lVolumeNumber, lVolumeEmployerCount, lVolumeEmployeeCount, lblFile, True)
        End If
      End If
      
      Call ee.KillBenefitsEx
NEXT_EMPLOYEE:
    Next
NEXT_EMPLOYER:
    Call p11d32.LoadEmployer(Nothing, False)
  Next
  
  If lEmployersWritten > 0 Then
    Call CheckNIRatio(lEmployeesWithInvalidNI, lEmployeesWritten)
    prgEmployee.Caption = "Completed magnetic media " & lVolumeNumber & " file(s) written, " & lEmployeesWritten & " employees written."
    prgEmployee.Indicator = None
    If Not m_TS Is Nothing Then
      m_TS.Write (Type3Record(lEmployersWritten, lEmployeesWritten))
      m_TS.Write (Type4Record(lVolumeEmployerCount, lVolumeEmployeeCount))
    End If
    Call CloseFile
  End If
  
CreateMagneticMediaFiles_END:
  UserDataSize = 0
  Set ErrorFilter = Nothing
  Call CloseErrorLogDB
  Call xReturn("CreateMagneticMediaFiles")
  Exit Sub
CreateMagneticMediaFiles_ERR:
  
  If Err.Number = ERR_SPACE_IN_PAYE_NUM Then
    Call ErrorMessage(ERR_SPACE_IN_PAYE_NUM, Err, "CheckEmployeePAYERef", FilterMessageTitle(ben.value(employer_Name_db), ee.FullName), Err.Description)
    Resume Next
  Else
    Call ErrorMessage(ERR_ERROR, Err, "CreateMagneticMediaFiles", FilterMessageTitle(), "Unable to create magnetic media files.")
  End If
  
  Resume CreateMagneticMediaFiles_END
  Resume
End Sub
Public Function Kill()
  Set m_db = Nothing
  Set m_ws = Nothing
End Function
Private Function BenLine(sbenline As String, ey As Employer, ee As Employee, benefits As ObjectList, BenArr() As BEN_CLASS, ByVal HS As HMIT_SECTIONS, ByVal MMS0 As MM_STANDARD_OUT, Optional bRecordHeader As Boolean = True) As Boolean
  Dim Description, value, MadeGood, benefit
  
  BenLine = Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, benefits, BenArr())
  sbenline = StandardOutLine(Description, value, MadeGood, benefit, ee, BenArr(1), HS, MMS0, bRecordHeader)
  
End Function

Private Function GetEmployeesBenefits(ey As IBenefitClass, ee As Employee) As String
  Dim Description, value, MadeGood, benefit
  Dim sbenline As String, s As String
  Dim HS As HMIT_SECTIONS
  Dim sHMITSectionString As String
  Dim BenArr() As BEN_CLASS, BenArr2() As BEN_CLASS
  Dim temp As String ' ek
   
  Dim i As Long
  Dim MM_SO As MM_STANDARD_OUT
  
  On Error GoTo GetEmployeesBenefits_ERR
  Call xSet("GetEmployeesBenefits")
  
  ReDim BenArr(1 To 1)
  BenArr(1) = -1
  
  For HS = HMIT_FIRST_ITEM To HMIT_LAST_ITEM
    sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
    Select Case HS
      Case HMIT_A
        If SectionA(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_B
        If SectionB2(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_C
        BenArr(1) = BC_VOUCHERS_AND_CREDITCARDS_C
        If AddCrLf(sbenline, BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_ASSETSTRANSFERRED_TYPE)) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_D
        BenArr(1) = BC_LIVING_ACCOMMODATION_D
        If AddCrLf(sbenline, BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VAN_TYPE)) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_E
        If SectionEandU(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_F
        If SectionF(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_G
        BenArr(1) = BC_NONSHAREDVANS_G
        If ee.AnyVanBenefit Then
          If AddCrLf(sbenline, BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VAN_TYPE)) Then
            GetEmployeesBenefits = GetEmployeesBenefits & sbenline
          End If
        End If
      Case HMIT_H
        If SectionH(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_I
        BenArr(1) = BC_PRIVATE_MEDICAL_I
        If AddCrLf(sbenline, BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_ASSETSTRANSFERRED_TYPE)) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_J
        BenArr(1) = BC_QUALIFYING_RELOCATION_J
        If AddCrLf(sbenline, BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VAN_TYPE)) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_K
        BenArr(1) = BC_SERVICES_PROVIDED_K
        If AddCrLf(sbenline, BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_ASSETSTRANSFERRED_TYPE)) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_L
        If SectionL(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
'      Case HMIT_M
'        BenArr(1) = BC_SHARES_M
'          If BenLine(sbenline, ey, ee, ee.benefits, BenArr(), HS, MM_SO_RECORD_HEADER, True) Then
'            GetEmployeesBenefits = GetEmployeesBenefits & sbenline & MMBool(True) & vbCrLf
'          End If
      Case HMIT_M
        If SectionM2(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
      Case HMIT_N
        If SectionN(sbenline, ey, ee) Then
          GetEmployeesBenefits = GetEmployeesBenefits & sbenline
        End If
    End Select
    
  Next

GetEmployeesBenefits_END:
  Call xReturn("GetEmployeesBenefits")
  Exit Function
GetEmployeesBenefits_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetEmployeesBenefits", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString), "Unable to get employee benefits.")
  Resume GetEmployeesBenefits_END
  Resume
End Function
Private Function SectionMTypePartialOutLine(ByVal bc As BEN_CLASS, IRDesc As Variant, Description As Variant, value As Variant, MadeGood As Variant, benefit As Variant) As String
  Dim s As String
  Dim iBenITem As Long
  
  SectionMPArtialOutLineError (bc)
  
  iBenITem = IRDescriptionBenItem(bc)
  s = s & PadFieldHelper(bc, iBenITem, IRDesc)
  s = s & PadFieldHelper(bc, ITEM_DESC, Description)
  s = s & PadFieldHelper(bc, ITEM_VALUE, value)
  s = s & PadFieldHelper(bc, ITEM_MADEGOOD_NET, MadeGood)
  s = s & PadFieldHelper(bc, ITEM_BENEFIT, benefit)
  
  SectionMTypePartialOutLine = s
End Function
Private Sub SectionMPArtialOutLineError(bc As BEN_CLASS)
  If (bc <> BC_CLASS_1A_M And bc <> BC_NON_CLASS_1A_M) Then
    Call Err.Raise(ERR_BENCLASS_INVALID, "SectionMPartialOutline", "Ben class can only be " & p11d32.Rates.BenClassTo(BC_CLASS_1A_M, BCT_UDM_BENEFIT_TITLE) & " or " & p11d32.Rates.BenClassTo(BC_NON_CLASS_1A_M, BCT_UDM_BENEFIT_TITLE) + " for a Section M benefit")
  End If
End Sub
Private Function SectionMTypePartialOutLineNoBenefits(ByVal bc As BEN_CLASS) As String
  Dim s As String
  Dim iPadField As Long
  Dim iBenITem As Long
  
  SectionMPArtialOutLineError (bc)
  
  iBenITem = IRDescriptionBenItem(bc)
  s = s & BlankField(bc, iBenITem)
  s = s & PadFieldHelper(bc, ITEM_DESC, "")
  s = s & PadFieldHelper(bc, ITEM_VALUE, 0)
  s = s & PadFieldHelper(bc, ITEM_MADEGOOD_NET, 0)
  s = s & PadFieldHelper(bc, ITEM_BENEFIT, 0)
  
  SectionMTypePartialOutLineNoBenefits = s
End Function

Private Function BlankFieldByRecordPosition(pStart As Long, pEnd As Long, Optional char As String = " ") As String
  BlankFieldByRecordPosition = String$(pEnd - pStart + 1, char)
End Function
Private Function BlankField(bc As BEN_CLASS, Item As Long, Optional char As String = " ") As String
  Dim iLen As Long
  iLen = p11d32.BenDataLinkMMFieldSize(bc, Item)
  BlankField = String$(iLen, char)
End Function

Private Function SectionM2(sbenline As String, ey As IBenefitClass, ee As Employee) As Boolean
  Dim BenArr() As BEN_CLASS
  Dim HS As HMIT_SECTIONS
  Dim Description As Variant, value As Variant, MadeGood As Variant, benefit As Variant
  Dim IRDesc As Variant
  Dim sbenline_class_1a As String, sbenline_non_class_1a As String, sbenline_2t_class_1a As String, sbenline_2t_non_class_1a As String
  Dim sbenline_2t As String, sbenline_2t_income_tax_not_dedcuted As String
  Dim sHMITSectionString As String
  Dim bAnyC1ABenefits As Boolean, bAnyNonClass1ABenefits As Boolean, bAnyIncomeTaxNotDeducted As Boolean
  
  On Error GoTo err_Err
  
  Call xSet("SectionM")
  
  sbenline = ""
  
  SectionM2 = False
  HS = HMIT_M
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
  
  
  'class 1A benefits, all stored under subs & professional
  ReDim BenArr(1 To 1)
  BenArr(1) = BC_CLASS_1A_M
  If (Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr(), IRDesc)) Then
    bAnyC1ABenefits = True
    sbenline_class_1a = SectionMTypePartialOutLine(BenArr(1), IRDesc, Description, value, MadeGood, benefit)
  Else
    sbenline_class_1a = SectionMTypePartialOutLineNoBenefits(BenArr(1))
  End If
  sbenline_2t_class_1a = PadFieldHelper(BC_CLASS_1A_M, ITEM_BENEFIT, benefit) 'get the zeros for the 2T line
  
  'class non class 1A benefits, all stored under nursery, and nother - do not include income tax not deducted
  'Charlie
  'Had a response from Pam.  If we have an "Income tax paid but not deducted", this should appear in field 6 of 2T, and nowhere on line 2N.  Therefore the cl1a total from N will be in field 4, non cl1a in 5 and income tax in field 6.  Only 4 and 5 are totals of line 2N.
  'If we do not have any cl1a or non cl1a, but have income tax, we need to have a 2T line appear.
  'Hope that makes sense.
  'Saida
  ReDim BenArr(1 To 1)
  BenArr(1) = BC_NON_CLASS_1A_M
    
  If (Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr(), IRDesc)) Then
    sbenline_non_class_1a = SectionMTypePartialOutLine(BenArr(1), IRDesc, Description, value, MadeGood, benefit)
    bAnyNonClass1ABenefits = True
  Else
    sbenline_non_class_1a = SectionMTypePartialOutLineNoBenefits(BenArr(1))
  End If
  
  sbenline_2t_non_class_1a = PadFieldHelper(BC_NON_CLASS_1A_M, ITEM_BENEFIT, benefit)
  
  ReDim BenArr(1 To 1)
  BenArr(1) = BC_INCOME_TAX_PAID_NOT_DEDUCTED_M
  If (Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr())) Then bAnyIncomeTaxNotDeducted = True
  sbenline_2t_income_tax_not_dedcuted = PadFieldHelper(BC_INCOME_TAX_PAID_NOT_DEDUCTED_M, ITEM_BENEFIT, benefit)
    
  'create the record entries
  If bAnyC1ABenefits Or bAnyNonClass1ABenefits Or bAnyIncomeTaxNotDeducted Then
    SectionM2 = True
    
    If (bAnyC1ABenefits Or bAnyNonClass1ABenefits) Then
      sbenline = StandardOutLine("", "", "", benefit, ee, -1, HS, MM_SO_RECORD_HEADER)
      sbenline = sbenline & sbenline_class_1a & sbenline_non_class_1a & vbCrLf
    End If
    
    '2T
    'header
    sbenline_2t = RecordType(MM_REC_BENEFIT_SECTION_N) & PadFieldHelper(BC_EMPLOYEE, ee_PersonnelNumber_db, ee.PersonnelNumber)
    sbenline_2t = sbenline_2t & sbenline_2t_class_1a & sbenline_2t_non_class_1a & sbenline_2t_income_tax_not_dedcuted & vbCrLf
    
    'final output
    sbenline = sbenline & sbenline_2t
  End If
  
  
err_End:
  Call xReturn("SectionM")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionM", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString), "Error getting the section M line.")
  Resume err_End
  Resume
End Function

Public Function MMBool(b As Variant) As String
  If b = True Then
    MMBool = "Y"
  ElseIf b = False Then
    MMBool = "N"
  Else
    MMBool = b
  End If
End Function
Private Function SectionN(sbenline As String, ey As IBenefitClass, ee As Employee) As Boolean
  Dim HS As HMIT_SECTIONS
  Dim BenArr() As BEN_CLASS
  Dim sHMITSectionString As String
  Dim bBenefitApplies
  Dim s As String
  Dim Description, value, MadeGood, benefit
  Dim IRDesc As Variant
  Dim bMultiple As Boolean
  Dim iBenITem As Long
  On Error GoTo SectionN_ERR
  
  Call xSet("SectionN")
  
  sbenline = ""
  HS = HMIT_N
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
  
  ReDim BenArr(1 To 1)
  'travel and subsistance
  BenArr(1) = BC_TRAVEL_AND_SUBSISTENCE_N
  SectionN = SectionN Or BenLine(s, ey, ee, ee.benefits, BenArr(), HS, MM_SO_ASSETSTRANSFERRED_TYPE)
  sbenline = sbenline & s
  'entertainment
  'RC 20040129  - Have added a new check to see whether the benefit applies, if it doesn't then add just a blank
  BenArr(1) = BC_ENTERTAINMENT_N
  bBenefitApplies = BenLine(s, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VALUE_MADEGOOD_BENEFIT, False)
  
  If bBenefitApplies = True Then
    sbenline = sbenline & MMBool(ey.value(employer_CT_db))
  Else
    sbenline = sbenline & " "
  End If
  
  SectionN = SectionN Or bBenefitApplies
  sbenline = sbenline & s
  'business travel
  BenArr(1) = BC_GENERAL_EXPENSES_BUSINESS_N
  SectionN = SectionN Or BenLine(s, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VALUE_MADEGOOD_BENEFIT, False)
  sbenline = sbenline & s
  'home telephone
  BenArr(1) = BC_PHONE_HOME_N
  SectionN = SectionN Or BenLine(s, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VALUE_MADEGOOD_BENEFIT, False)
  sbenline = sbenline & s
  'non qualifying relocation
  BenArr(1) = BC_NON_QUALIFYING_RELOCATION_N
  SectionN = SectionN Or BenLine(s, ey, ee, ee.benefits, BenArr(), HS, MM_SO_VALUE_MADEGOOD_BENEFIT, False)
  sbenline = sbenline & s
  'non other and chauffeur
  ReDim BenArr(1 To 2)
  BenArr(1) = BC_OOTHER_N
  BenArr(2) = BC_CHAUFFEUR_OTHERO_N
  
  SectionN = SectionN Or Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr(), IRDesc)
  'we seem to be able to have some common funtion here
   
  iBenITem = IRDescriptionBenItem(BC_OOTHER_N)
  sbenline = sbenline & PadFieldHelper(BC_OOTHER_N, iBenITem, IRDesc)
  sbenline = sbenline & PadFieldHelper(BC_OOTHER_N, ITEM_DESC, Description)
  
  sbenline = sbenline & PadFieldHelper(BC_OOTHER_N, ITEM_VALUE, value) & _
                PadFieldHelper(BC_OOTHER_N, ITEM_MADEGOOD_NET, MadeGood) & _
                PadFieldHelper(BC_OOTHER_N, ITEM_BENEFIT, benefit)
  
  sbenline = sbenline & vbCrLf
SectionN_END:
  Call xReturn("SectionN")
  Exit Function
SectionN_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SectionN", FilterMessageTitle(ey, ee, sHMITSectionString), "Error getting the section O line.")
  Resume SectionN_END
End Function
Private Function SectionF(sbenline As String, ey As Employer, ee As Employee) As Boolean
  Dim HS As HMIT_SECTIONS
  Dim sHMITSectionString As String
  Dim ben As IBenefitClass
  Dim ben_err As IBenefitClass
  Dim car As CompanyCar
  Dim FuelBen As IBenefitClass
  Dim BenArr() As BEN_CLASS
  Dim value, MadeGood, benefit, Description
  Dim i As Long, lErrorNum As Long
  Dim d As Date
  
  Dim temp As Long
  Dim temp2 As Boolean
  
  Dim bField6Used As Boolean
  Dim bField7Used As Boolean
  Dim bField8Used As Boolean
  Dim bField9Used As Boolean
  Dim sFuelType As String
  Dim sFuelReinstated As String
  Dim dFuelAvailableTo As Date
  ' Dim sRegistrationyear As String
    
  'EK as only one Datewitdrawn box in 2003, has to be altered 2004
  
  'EK accessories merged into one.
  Dim lTotalAccessories As Long
  
  Set ben_err = ey
    
  On Error GoTo SectionF_ERR
  
  Call xSet("SectionF")
  
  
  HS = HMIT_F
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HMIT_F)
  
  sbenline = ""
  
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = BC_COMPANY_CARS_F Then
        Set car = ben
'MP DB ToDo chk FuelBen usage in this proc - is it still valid?
        Set FuelBen = car.Fuel
        If CheckBen(ey, ee, ben) Then
                  
          
          Call CompanyCarAvailableDatesValid(ben)
          Call CompanyCarCheckDateFuelAvaialbleTo(ben)
                    sbenline = sbenline & StandardOutLine(0, 0, 0, ben, ee, -1, HS, MM_SO_RECORD_HEADER) & _
                     PadFieldHelper(BC_COMPANY_CARS_F, car_Make_db, ben.value(car_Make_db) & " " & ben.value(car_Model_db)) & _
                     PadFieldHelper(BC_COMPANY_CARS_F, car_Registrationdate_db, ben.value(car_Registrationdate_db))
                      
                     
          sFuelType = PadFieldHelper(BC_COMPANY_CARS_F, car_p46FuelTypeString, ben.value(car_p46FuelTypeString))
          
          
          'field6 is has any approved co2
          bField6Used = ((ben.value(car_p46CarbonDioxide_db) > 0) And ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE) And Not sFuelType = "E")
          bField7Used = IIf((Not bField6Used And ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE) And ben.value(car_p46NoApprovedCO2Figure_db)) Or sFuelType = "E", True, False)
          bField8Used = IIf(ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Or (ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE) And bField7Used), True, False)
          bField9Used = IIf(ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE), True, False)     'RC20040127 Reduced conditions according to new 2004 MM doc
          
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_p46CarbonDioxide_db, IIf(bField6Used, ben.value(car_p46CarbonDioxide_db), "   "))
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_p46NoApprovedCO2Figure_db, IIf(bField7Used, "Y", " "))
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_enginesize_db, IIf(bField8Used, IIf(sFuelType = "E", "0000", ben.value(car_enginesize_db)), "    "))
          
          
          sbenline = sbenline & IIf(bField9Used, sFuelType, " ")  'ek not sure if another condition is needed
         
          'RC 20040112 Have changed below to use car_CapitalContributionRestricted instead of car_capitalcontribution_db
          'EK accessories merged 2/04
          lTotalAccessories = ben.value(car_AccessoriesOriginal_db) + ben.value(car_AccessoriesNew_db) - ben.value(car_CheapAccessories_db)
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, Car_AvailableFrom_db, ben.value(Car_AvailableFrom_db), MMDC_TAX_YEAR_START)
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, Car_AvailableTo_db, ben.value(Car_AvailableTo_db), MMDC_TAX_YEAR_END)
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_ListPrice_db, ben.value(car_ListPrice_db))
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_AccessoriesOriginal_db, lTotalAccessories)
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_capitalcontribution_db, ben.value(car_CapitalContributionRestricted))
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, ITEM_MADEGOOD_NET, ben.value(ITEM_MADEGOOD_NET))
          
          If FuelBen.value(ITEM_BENEFIT) = 0 Then
            d = UNDATED
          Else
            d = ben.value(Car_FuelAvailableTo_calc)
          End If
          
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, Car_FuelAvailableTo_db, d)
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, car_fuelreinstated_db, IIf(ben.value(car_fuelreinstated_calc), "Y", " "))
          
          sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, ITEM_BENEFIT, ben.value(ITEM_BENEFIT))
          sbenline = sbenline & PadFieldHelper(BC_FUEL_F, ITEM_BENEFIT, IIf((bField8Used Or bField6Used) And FuelBen.value(ITEM_BENEFIT) > 0 And Not sFuelType = "E", FuelBen.value(ITEM_BENEFIT), "00000000")) & vbCrLf
          If ((bField8Used And ben.value(car_enginesize_db) = 0) Or (FuelBen.value(ITEM_BENEFIT) > 0 And (ben.value(car_enginesize_db) = 0 And ben.value(car_p46CarbonDioxide_db) = 0)) And sFuelType <> "E") Then
            Call Err.Raise(ERR_NEEDCO2ORENGINESIZE, , "Car '" & ben.value(car_Registration_db) & "' needs an engine size or C02 figure.")
          End If
        

          Call CheckCarListPrice(ben)
          
          'RC 20040219 only fuel type E permitted if car is reg. on or after 010198
          If sFuelType = "E" And ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
            Call Err.Raise(ERR_CAR_ELEC_PRE_98, , "Car " & ben.value(car_Registration_db) & "only fuel type E permitted if car is reg. on or after 010198")
          End If
          
          'removed as no longer checked for by PayeOnline
          '
          'If ben.value(car_p46FuelType_db) = CCFT_DIESEL And ben.value(car_Registrationdate_db) >= D_FUEL_TYPE_L_NO_LONGER_HAS_DISCOUNT Then
          '  Call Err.Raise(ERR_DIESEL_REGISTERED_AFTER_1_1_2006, , "Has fuel type D, change to L as registered after 31/12/2005")
          'End If
          
        End If
      End If
    End If
  Next
  
  If Len(sbenline) > 0 Then
    ReDim BenArr(1 To 1)
    BenArr(1) = BC_COMPANY_CARS_F
    Call Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr)
    sbenline = sbenline & RecordType(MM_REC_BENFIT_SECTION_F) & PadFieldHelper(BC_EMPLOYEE, ee_PersonnelNumber_db, ee.PersonnelNumber) & _
               PadFieldHelper(BC_COMPANY_CARS_F, ITEM_BENEFIT, benefit)
    BenArr(1) = BC_FUEL_F
    
    Call Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr)
    If benefit > 0 Then
      sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, ITEM_BENEFIT, benefit) & vbCrLf
    Else
      sbenline = sbenline & PadFieldHelper(BC_COMPANY_CARS_F, ITEM_BENEFIT, "", , TYPE_STR) & vbCrLf
    End If
    

    
    SectionF = True
  End If
  
  
SectionF_END:
  Call xReturn("SectionF")
  Exit Function
SectionF_ERR:
  If Err.Number = ERR_ROTARY Then
    Call ErrorMessage(ERR_ROTARY, Err, "SectionF", FilterMessageTitle(ben_err.value(employer_Name_db), , "Section F", "Company Cars"), Err.Description)
    Resume Next
  ElseIf Err.Number = ERR_CAR_FUEL_AVAILABLE_TO_INVALID Or Err.Number = ERR_CAR_LIST_PRICE Or Err.Number = ERR_NEEDCO2ORENGINESIZE Or Err.Number = ERR_MULTIPLEFUELWITHDRAWN Or Err.Number = ERR_CAR_DATES_INVALID Or Err.Number = ERR_CAR_ELEC_PRE_98 Or Err.Number = ERR_DIESEL_REGISTERED_AFTER_1_1_2006 Then
    Call ErrorMessage(Err.Number, Err, "SectionF", FilterMessageTitle(ben_err.value(employer_Name_db), ee.FullName, "Section F", "Company Cars"), Err.Description)
    Resume Next
  Else
    Call ErrorMessage(ERR_ERROR, Err, "SectionF", FilterMessageTitle(ey, ee, sHMITSectionString), "Error getting the section F line.")
    Resume SectionF_END
    Resume
  End If
End Function
'does not append crlf on summary
Private Function SectionAType(sbenline As String, ey As Employer, ee As Employee, BenArr() As BEN_CLASS, HS As HMIT_SECTIONS, mmrc As MM_RECORD_TYPE) As Boolean
  Dim sHMITSectionString As String
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  On Error GoTo err_Err
  
  Call xSet("SectionAType")
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
  
  sbenline = ""
    
  If Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr(), IRDesc) Then
    sbenline = StandardOutLine(Description, value, MadeGood, benefit, ee, BenArr(1), HS, MM_SO_IR_DESC_TYPE_BENEFIT_AND_VALUE, True, IRDesc) & vbCrLf
    sbenline = sbenline & StandardOutLine("", "", "", benefit, ee, BenArr(1), HS, MM_SO_SUMMARY_TYPE, , , mmrc) & vbCrLf
    SectionAType = True
  End If
  
    
err_End:
  Call xReturn("SectionAType")
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, "SectionAType", "Error generating section A type line")
  Resume err_End
  Resume
End Function

Private Function SectionB2(sbenline As String, ey As Employer, ee As Employee) As Boolean
  Dim sHMITSectionString As String
  Dim HS As HMIT_SECTIONS
  Dim BenArr() As BEN_CLASS
  Dim staxnotional As String
  Dim Description As Variant, MadeGood As Variant, benefit As Variant, value As Variant, IRDesc As Variant
  Dim bTaxOnNotional As Boolean
  On Error GoTo err_Err
  
  Call xSet("SectionB")
  
  HS = HMIT_B
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
  sbenline = ""
  
  ReDim BenArr(1 To 1)
  BenArr(1) = BC_TAX_NOTIONAL_PAYMENTS_B
  
  If (Validations.SumBenefit("", "", "", benefit, ey, ee, ee.benefits, BenArr())) Then
    bTaxOnNotional = True
    SectionB2 = True
  End If
  
  staxnotional = PadFieldHelper(BenArr(1), ITEM_BENEFIT, benefit)
  
  BenArr(1) = BC_PAYMENTS_ON_BEFALF_B
  
  If Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, ee.benefits, BenArr(), IRDesc) Then
    sbenline = StandardOutLine(Description, value, MadeGood, benefit, ee, BenArr(1), HS, MM_SO_IR_DESC_TYPE_BENEFIT, True, IRDesc) & vbCrLf
    SectionB2 = True
  End If
  
  If SectionB2 Then
    sbenline = sbenline & StandardOutLine("", "", "", benefit, ee, BenArr(1), HS, MM_SO_SUMMARY_TYPE, , , MM_REC_BENEFIT_SECTION_B) & staxnotional & vbCrLf
  End If
  

  
err_End:
  Call xReturn("SectionB")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionB", FilterMessageTitle(ey, ee, sHMITSectionString), "Error getting the section B line.")
  Resume err_End
  Resume
End Function
Private Function SectionA(sbenline As String, ey As Employer, ee As Employee) As Boolean
  Dim sHMITSectionString As String
  Dim HS As HMIT_SECTIONS
  Dim BenArr() As BEN_CLASS
  On Error GoTo err_Err
  
  Call xSet("SectionA")
  
  ReDim BenArr(1 To 1)
  BenArr(1) = BC_ASSETSTRANSFERRED_A
  HS = HMIT_A
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
  sbenline = ""
  SectionA = SectionAType(sbenline, ey, ee, BenArr(), HS, MM_REC_BENEFIT_SECTION_A)
  
err_End:
  Call xReturn("SectionA")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionA", FilterMessageTitle(ey, ee, sHMITSectionString), "Error getting the section A line.")
  Resume err_End
  Resume
End Function

Public Function MMDate(dValue As Variant, ByVal mmdc As MMDATE_COMPARE) As String

  Select Case mmdc
    Case MMDATE_COMPARE.MMDC_TAX_YEAR_END
      If dValue = p11d32.Rates.value(TaxYearEnd) Then
        dValue = ""
      End If
    Case MMDATE_COMPARE.MMDC_TAX_YEAR_START
      If dValue = p11d32.Rates.value(TaxYearStart) Then
        dValue = ""
      End If
    Case MMDATE_COMPARE.MMDC_NO_COMPARE
    Case Else
      ECASE ("Invalid MM Date compare")
  End Select
  If (Len(dValue) > 0) Then
    dValue = dValue
  End If
  If (dValue = UNDATED) Then
    MMDate = ""
  Else
    MMDate = DateString2(dValue, "", "ddmmyyyy")
  End If
  
  
End Function

Private Function PadNegative(ByVal FieldData As Variant, ByVal FieldLen As Long, MM_FT As DATABASE_FIELD_TYPES) As String
  If FieldData < 0 Then
    FieldLen = FieldLen - 1
    FieldData = FieldData * -1
    PadNegative = PadField(FieldData, FieldLen, TYPE_LONG)
    PadNegative = "-" & PadNegative
  Else
    PadNegative = PadField(FieldData, FieldLen, TYPE_LONG)
  End If
End Function
Private Function SectionEandU(sbenline As String, ey As IBenefitClass, ee As Employee) As Boolean
  Dim sHMITSectionString As String
  Dim ben As IBenefitClass
  Dim i As Long
  Dim bc As BEN_CLASS
  Dim lNetBenefit As Long
  Dim iCounter As Integer
  Dim sBenefitFormCaption As String
  Dim bReportable As Boolean
  
  On Error GoTo SectionEandU_ERR
  
  Call xSet("SectionEandU")
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HMIT_E)
  sBenefitFormCaption = p11d32.Rates.BenClassTo(BC_EMPLOYEE_CAR_E, BCT_FORM_CAPTION)
  
  sbenline = ""
    
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = BC_EMPLOYEE_CAR_E Then
        lNetBenefit = lNetBenefit + ben.value(ITEM_BENEFIT)
        iCounter = iCounter + 1
        'was bReportable = ben.value(eecar_MMSectionU_Reportable), this would only get the value for the last car
        bReportable = bReportable Or ben.value(eecar_MMSectionU_Reportable)
      End If
    End If
  Next
  If iCounter > 0 Then
    If lNetBenefit > 0 Then
      sbenline = StandardOutLine("", "", "", ben, ee, BC_EMPLOYEE_CAR_E, HMIT_E, MM_SO_RECORD_HEADER) & _
                            PadFieldHelper(BC_EMPLOYEE_CAR_E, ITEM_BENEFIT, lNetBenefit) & vbCrLf
      SectionEandU = True
    ElseIf lNetBenefit < 0 And bReportable Then
      If lNetBenefit < -999999 Then Call Err.Raise(ERR_MARORS, , "MARORS value < -999,999 " & lNetBenefit)
      sbenline = RecordType(MM_REC_BENEFIT_SECTION_U)
      sbenline = sbenline & PadFieldHelper(BC_EMPLOYEE, ee_PersonnelNumber_db, ee.PersonnelNumber)
      sbenline = sbenline & PadNegative(lNetBenefit, p11d32.BenDataLinkMMFieldSize(BC_EMPLOYEE_CAR_E, eecar_BenefitRepayableTaxable), p11d32.BenDataLinkDataType(BC_EMPLOYEE_CAR_E, eecar_BenefitRepayableTaxable)) & vbCrLf
      SectionEandU = True
    Else
      Call Err.Raise(ERR_BEN_NOT_REPORTABLE, "SectionEandU", "Benefit not reportable for magnetic media purposes.")
    End If
  End If
  
SectionEandU_END:
  Call xReturn("SectionEandU")
  Exit Function
SectionEandU_ERR:
  If Not ben Is Nothing Then
    If (Err.Number = ERR_BEN_NOT_REPORTABLE) Or (Err.Number = ERR_BEN_INCORRECT) Then
      Call ErrorMessagePop(ERR_ERROR, Err, "SectionEandU", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name, sBenefitFormCaption), "")
    ElseIf Err.Number = ERR_MARORS Then
      Call ErrorMessage(ERR_MARORS, Err, "CheckMARORS", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name), Err.Description)
      Resume Next
    Else
      Call ErrorMessagePop(ERR_ERROR, Err, "SectionEandU", FilterMessageTitle(ey.Name, ee.FullName, sHMITSectionString, ben.Name, sBenefitFormCaption), "Unable to sum benefits.")
    End If
  Else
    Call ErrorMessagePop(ERR_ERROR, Err, "SectionEandU", FilterMessageTitle(ey.Name, ee.FullName, , "Unknown"), "Unable to sum benefits.")
  End If
  Resume SectionEandU_END
End Function

Private Function SectionH(sbenline As String, ey As Employer, ee As Employee) As Boolean
    Dim sHMITSectionString As String
  Dim ben As IBenefitClass
  Dim BenArr(1) As IBenefitClass
  Dim i As Long
  Dim bc As BEN_CLASS
  Dim loans As loans

  On Error GoTo SectionH_ERR
  
  Call xSet("SectionH")
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HMIT_F)
  
  sbenline = ""
  i = ee.GetLoansBenefitIndex
  If i = 0 Then GoTo SectionH_END
  Set loans = ee.benefits(i)
  If loans Is Nothing Then GoTo SectionH_END
  bc = BC_LOAN_OTHER_H
  
  For i = 1 To loans.loans.Count
    Set ben = loans.loans(i)
    If Not ben Is Nothing Then
      If CheckBen(ey, ee, ben) Then
          sbenline = sbenline & StandardOutLine(0, 0, 0, ben, ee, -1, HMIT_H, MM_SO_RECORD_HEADER) & _
                     PadFieldHelper(bc, ln_nborrowers_db, ben.value(ln_nborrowers_db)) & _
                     PadFieldHelper(bc, ln_OpenOutstanding, ben.value(ln_OpenOutstanding)) & _
                     PadFieldHelper(bc, ln_CloseOutstanding, ben.value(ln_CloseOutstanding)) & _
                     PadFieldHelper(bc, ln_MaxOutstandingAtAnyPoint, ben.value(ln_MaxOutstandingAtAnyPoint)) & _
                     PadFieldHelper(bc, ln_InterestPaid_db, ben.value(ln_InterestPaid_db)) & _
                     PadFieldHelper(bc, ln_MadeDate, GetBenItem(ben, ln_MadeDate)) & _
                     PadFieldHelper(bc, ln_DischargedDate, GetBenItem(ben, ln_DischargedDate)) & _
                     PadFieldHelper(bc, ITEM_BENEFIT, ben.value(ITEM_BENEFIT)) & vbCrLf
      End If
    End If
  Next
  
  SectionH = Len(sbenline)
  
SectionH_END:
  Call xReturn("SectionH")
  Exit Function
SectionH_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SectionH", FilterMessageTitle(ey, ee, sHMITSectionString), "Error getting the section H line.")
  Resume SectionH_END
End Function
Private Function SectionL(sbenline As String, ey As Employer, ee As Employee) As Boolean
  Dim sHMITSectionString As String
  Dim HS As HMIT_SECTIONS
  Dim BenArr() As BEN_CLASS
  On Error GoTo err_Err
  
  Call xSet("SectionA")
  
  ReDim BenArr(1 To 1)
  BenArr(1) = BC_ASSETSATDISPOSAL_L
  HS = HMIT_L
  
  sHMITSectionString = p11d32.Rates.HMITSectionToHMITCode(HS)
  sbenline = ""
  SectionL = SectionAType(sbenline, ey, ee, BenArr(), HS, MM_REC_BENEFIT_SECTION_L)
  
err_End:
  Call xReturn("SectionL")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionL", FilterMessageTitle(ey, ee, sHMITSectionString), "Error getting the section L line.")
  Resume err_End
  Resume
End Function
Private Function AddCrLf(sLine As String, bAdd As Boolean) As Boolean
  If bAdd Then sLine = sLine & vbCrLf
  AddCrLf = bAdd
End Function
Private Property Get PathAndFile(ByVal lVolNumber As Long) As String
  PathAndFile = OutputDirectory & FileName(lVolNumber)
End Property
Private Property Get FileName(ByVal lVolNumber As Long) As String
  FileName = MMStr(PadField(CLng(SubmitterRef), 6, TYPE_STR)) & "." & PadField(lVolNumber, 3, TYPE_LONG)
End Property
Public Sub Start()
  Dim i As Long
  Dim emp As Employer
  Dim li As ListItem
  
  On Error GoTo Start_ERR
  Call xSet("Start")
  
  
  Call EmployersToListView
  
    F_MM.optDataFormat(DataFormat).value = True
    F_MM.Caption = "Magnetic Media " & p11d32.AppYear & "/" & p11d32.AppYear + 1   'RH
    Call p11d32.Help.ShowForm(F_MM, vbModal)

Start_END:
  Set p11d32.CurrentEmployer = Nothing
  Set F_MM = Nothing
  Call xReturn("Start")
  Exit Sub
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error starting Magnetic Media.")
  Resume Start_END
End Sub
Public Function StandardOutLine(Description, value, MadeGoodNet, benefit, ee As IBenefitClass, bc As BEN_CLASS, HS As HMIT_SECTIONS, ByVal MMS0 As MM_STANDARD_OUT, Optional bRecordHeader As Boolean = True, Optional IRDesc, Optional RecordHeaderType As MM_RECORD_TYPE = MM_REC_BENEFIT) As String
  Dim sRecordHeader As String
  Dim iBenITem As Long
  On Error GoTo StandardOutLine_ERR
  
  Call xSet("StandardOutLine")
  
  If bRecordHeader Then sRecordHeader = RecordType(RecordHeaderType, HS) & PadFieldHelper(BC_EMPLOYEE, ee_PersonnelNumber_db, ee.value(ee_PersonnelNumber_db))
  
  With p11d32
    If Not ee Is Nothing Then
      Select Case MMS0
        Case MM_SO_VAN_TYPE, MM_SO_VAN_TYPE_WITH_DESCRIPTION
          StandardOutLine = sRecordHeader & _
                            IIf(MMS0 = MM_SO_VAN_TYPE_WITH_DESCRIPTION, PadFieldHelper(bc, ITEM_DESC, Description), "") & PadFieldHelper(bc, ITEM_VALUE, benefit)
        Case MM_SO_ASSETSTRANSFERRED_TYPE, MM_SO_ASSETSTRANSFERRED_TYPE_WITH_DESCRIPTION
          StandardOutLine = sRecordHeader & _
                            IIf(MMS0 = MM_SO_ASSETSTRANSFERRED_TYPE_WITH_DESCRIPTION, PadFieldHelper(bc, ITEM_DESC, Description), "") & StandardOutLine(Description, value, MadeGoodNet, benefit, ee, bc, HS, MM_SO_VALUE_MADEGOOD_BENEFIT)
        Case MM_SO_VALUE_MADEGOOD_BENEFIT
          StandardOutLine = PadFieldHelper(bc, ITEM_VALUE, value) & _
                            PadFieldHelper(bc, ITEM_MADEGOOD_NET, MadeGoodNet) & PadFieldHelper(bc, ITEM_BENEFIT, benefit)
        Case MM_SO_RECORD_HEADER
          StandardOutLine = sRecordHeader
        Case MM_SO_ASSETSTRANSFERRED_TYPE_WITH_DESCRIPTION_NO_RECORD_HEADER
          StandardOutLine = PadFieldHelper(bc, ITEM_DESC, Description) & StandardOutLine(Description, value, MadeGoodNet, benefit, ee, bc, HS, MM_SO_VALUE_MADEGOOD_BENEFIT)
        Case MM_SO_NONE
          StandardOutLine = PadFieldHelper(bc, ITEM_VALUE, "") & _
                            PadFieldHelper(bc, ITEM_MADEGOOD_NET, "") & PadFieldHelper(bc, ITEM_BENEFIT, "")
        Case MM_SO_SECTION_B 'CAD CHECK, this can be removed
          StandardOutLine = sRecordHeader & PadFieldHelper(bc, ITEM_DESC, IRDesc) & PadFieldHelper(bc, ITEM_DESC, Description) & PadFieldHelper(bc, ITEM_BENEFIT, benefit)
        Case MM_SO_IR_DESC_TYPE_BENEFIT_AND_VALUE
          iBenITem = IRDescriptionBenItem(bc)
          StandardOutLine = sRecordHeader & PadFieldHelper(bc, iBenITem, IRDesc) & PadFieldHelper(bc, ITEM_DESC, Description) & PadFieldHelper(bc, ITEM_BENEFIT, value) & PadFieldHelper(bc, ITEM_BENEFIT, MadeGoodNet) & PadFieldHelper(bc, ITEM_BENEFIT, benefit)
        Case MM_SO_IR_DESC_TYPE_BENEFIT
          iBenITem = IRDescriptionBenItem(bc)
          StandardOutLine = sRecordHeader & PadFieldHelper(bc, iBenITem, IRDesc) & PadFieldHelper(bc, ITEM_DESC, Description) & PadFieldHelper(bc, ITEM_BENEFIT, benefit)
         Case MM_SO_SUMMARY_TYPE
          StandardOutLine = sRecordHeader & PadFieldHelper(bc, ITEM_BENEFIT, benefit)
        Case Else
          Call ECASE("Invalid MM Standard Out type.")
      End Select
    Else
      Call Err.Raise(ERR_NO_EMPLOYEE, "StandardOutLine", "The benefits parent is expected to be an employee but is nothing.")
    End If
  End With
  
  
StandardOutLine_END:
  Call xReturn("StandardOutLine")
  Exit Function
StandardOutLine_ERR:
  StandardOutLine = ""
  Call ErrorMessage(ERR_ERROR, Err, "StandardOutLine", "Magnetic Media Van Type", "Error composing a magnetic media van type output string.")
  Resume StandardOutLine_END
  Resume
End Function
Private Function PadFieldHelper(bc As BEN_CLASS, Item As Long, Data As Variant, Optional mmdc As MMDATE_COMPARE = MMDC_NO_COMPARE, Optional MM_FT_OVERIDE As DATABASE_FIELD_TYPES = TYPE_BLOB) As String
  Dim MM_FT As DATABASE_FIELD_TYPES
  
  
  If MM_FT_OVERIDE <> TYPE_BLOB Then
    MM_FT = MM_FT_OVERIDE
  Else
    MM_FT = p11d32.BenDataLinkDataType(bc, Item)
  
  End If
  
  PadFieldHelper = PadField(Data, p11d32.BenDataLinkMMFieldSize(bc, Item), MM_FT, mmdc)
End Function
Public Function PadField(ByVal FieldData As Variant, ByVal FieldLen As Long, ByVal MM_FT As DATABASE_FIELD_TYPES, Optional MMDATE_COMPARE = MMDC_NO_COMPARE) As String
  Dim s As String
  On Error GoTo PadField_ERR
  
  Call xSet("PadField")
  
  If FieldLen = 0 Then Call Err.Raise(ERR_FIELD_LEN_0, "PadField", "The fieldsize passed can not be 0.")
  Select Case MM_FT
    Case TYPE_STR
      PadField = xStrPad(MMStr(FieldData), " ", FieldLen)
    Case TYPE_DOUBLE, TYPE_LONG, TYPE_MEAN, TYPE_NOSUM, TYPE_SUM
      PadField = xStrPad(FieldData, "0", FieldLen, True)
    Case TYPE_BOOL
      PadField = MMBool(FieldData)
    Case TYPE_DATE
      PadField = xStrPad(MMDate(FieldData, MMDATE_COMPARE), " ", FieldLen)
    Case Else
      ECASE ("Invalid field type passed to Magnetic media pad field, field type =" & MM_FT & ".")
  End Select
  
PadField_END:
  Call xReturn("PadField")
  Exit Function
PadField_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PadField", "Pad Field", "Error padding a field , field data = " & FieldData & ", field length = " & FieldLen & ".")
  Resume PadField_END
  Resume
End Function
Public Function RecordType(MM_R_T As MM_RECORD_TYPE, Optional HS As HMIT_SECTIONS = -1) As String
  Select Case MM_R_T
    
    Case MM_REC_FILE_OPEN
      RecordType = "0"
    Case MM_REC_EMPLOYER_OPEN
      RecordType = "1"
    Case MM_REC_EMPLOYEE
      RecordType = "20"
    Case MM_REC_BENEFIT
      If HS < HMIT_FIRST_ITEM Or HS > HMIT_LAST_ITEM Then
        ECASE ("Invalid HMIT_SECTION in getting magnetic media record type for benefit, HMIT_SECTION =" & HS)
      Else
        RecordType = "2" & p11d32.Rates.HMITSectionToHMITCode(HS)
      End If
    Case MM_REC_BENFIT_SECTION_F
      RecordType = "2X"
    Case MM_REC_BENEFIT_SECTION_U
      RecordType = "2U"
    Case MM_REC_BENEFIT_SECTION_A 'AM
      RecordType = "2Q"
    Case MM_REC_BENEFIT_SECTION_B 'AM
      RecordType = "2R"
    Case MM_REC_BENEFIT_SECTION_L 'AM
      RecordType = "2S"
    Case MM_REC_BENEFIT_SECTION_N 'AM
      RecordType = "2T"
    Case MM_REC_EMPLOYER_CLOSE
      RecordType = "3"
    Case MM_REC_FILE_CLOSE
      RecordType = "4"
    Case Else
      ECASE ("invalid Magnetic media record type.")
  End Select

End Function
'IK 02/04 moved to validations module
'Public Function ValidForEDI(ByVal s As String) As Boolean
'  Dim r As String
'
'  ValidForEDI = Not (StrComp(s, MMStr(s), vbTextCompare) <> 0)
'
'End Function
'Public Function MMStr(v As Variant) As String
'  Dim j As Long, k As Long
'  Dim r As String
'
'  On Error GoTo MMStr_ERR
'
'  Call xSet("MMStr")
'
'  MMStr = UCase$(Trim$(v))
'  MMStr = ReplaceString(MMStr, r, Chr(13))
'  MMStr = ReplaceString(MMStr, r, Chr(10))
'
'  Do While j < Len(MMStr)
'    j = j + 1
'    r = Mid$(MMStr, j, 1)
'    k = AscB(r)
'    If k < vbKeyA Or k > vbKeyZ Then
'      If k < vbKey0 Or k > vbKey9 Then
'        If InStr(1, "/-,.'&() ", r) = 0 Then
'          MMStr = ReplaceString(MMStr, r, "")
'        End If
'      End If
'    End If
'  Loop
'
'  MMStr = Trim$(MMStr)
'
'MMStr_END:
'  Call xReturn("MMStr")
'  Exit Function
'
'MMStr_ERR:
'  Call ErrorMessage(ERR_ERROR, Err, "MMStr", "MMStr", "Error removing the invalid magnetic media characters from " & v & ".")
'  Resume MMStr_END
'  Resume
'
'End Function
Private Property Get ReportHeader() As String
  On Error GoTo ReportHeader_ERR
  
  Call xSet("ReportHeader")
  
  ReportHeader = "Magnetic media warnings and errors" & vbCrLf & vbCrLf
  
ReportHeader_END:
  Call xReturn("ReportHeader")
  Exit Property
ReportHeader_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportHeader", "Report Header", "Error getting the report header for import errors report.")
  Resume ReportHeader_ERR
End Property

Private Sub Class_Initialize()
  ErrorLogging = True
End Sub

Private Function IErrorFilter_FilterErrorMessage(ByVal Username As String, ByVal DateTime As Date, ByVal ErrorNumber As Long, ByVal ErrorName As String, ByVal ErrorText As String, ByVal SourceFunction As String) As Boolean
  Dim s() As String
  Dim lEqualsPos As Long
  
  On Error GoTo FilterError_err
   
  m_rsError.AddNew
  
  If ErrorNumber = 0 Then
    ECASE ("Error number in magnetic media error filter is 0")
  End If
  ErrorCount = ErrorCount + 1
   
  m_rsError.Sort = "errorNumber"
  
  m_rsError.Fields("Error Number") = ErrorNumber
  m_rsError.Fields("Description") = ErrorText
  
  Call GetDelimitedValues(s, ErrorName)
  m_rsError.Fields("Employer") = s(1)
  m_rsError.Fields("Employee") = s(2)
  m_rsError.Fields("HMIT Section") = s(3)
  m_rsError.Fields("Benefit") = s(4)
  m_rsError.Fields("Benefit type") = s(5)
  
  m_rsError.Update
  
FilterError_end:
  Exit Function
FilterError_err:
  Resume Next
End Function
Public Sub Errors(ByVal Dest As REPORT_TARGET)

  
  On Error GoTo Errors_ERR
  
  Call xSet("Errors")

  Call OpenErrorLogDB
  Call StartAutoSTD(sql.Queries(SELECT_MAGNETIC_MEDIA_ERRORS), m_db, ReportHeader, Dest)
  
  
Errors_END:
  Call CloseErrorLogDB
  Call xReturn("Errors")
  Exit Sub
Errors_ERR:
  If Not m_db Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "Errors", "Errors", "Error viewing/printing the magnetic media errors from " & m_db.Name & ".")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "Errors", "Errors", "Error viewing/printing the magnetic media errors.")
  End If
  Resume Errors_END
End Sub
  
Private Function IErrorFilter_PrintAvailable() As Boolean
End Function
Private Function IErrorFilter_PrintErrors() As Boolean
End Function
Private Function IErrorFilter_ViewAvailable() As Boolean
End Function
Private Function IErrorFilter_ViewErrors() As Boolean
End Function
Private Function bMultipleRecs(bc As BEN_CLASS, ee As Employee, ey As Employer) As Boolean
Dim ben As IBenefitClass
Dim i As Long, j As Long

  Call xSet("bMultipleRecs")

  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = bc Then
        If ben.value(ITEM_BENEFIT_REPORTABLE) Then
        
          If AllMMRequiredFieldsPresent(ey, ee, ben) Then
        
            If IsNumeric(ben.value(ITEM_BENEFIT)) Then
        
              j = j + 1
              If j > 1 Then
                bMultipleRecs = True
                Exit For
              End If
        
            End If
        
          End If
        
        End If
      End If
    End If
  Next
  
End Function

Public Property Get OutputDirectory() As String
  OutputDirectory = m_OutputDirectory
End Property

Public Property Let OutputDirectory(sNewValue As String)
  m_OutputDirectory = FullPath(sNewValue)
End Property



