VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITCSUDM

Private m_CurrentDataSet As Variant
Private Const S_EMPLOYER As String = "Employer"
Private Const S_EMPLOYEES As String = "Employees"
Private Const S_BENEFITS As String = "Benefits"

Private Const L_DB_MOVE_ERROR As Long = 3021
Private Const L_DB_BOF As Long = 0
Private Const L_DB_EOF As Long = -1
Private Const L_DB_RECORD_COUNT_NOT_INITALISED As Long = -1

Private Enum UDM_DATA_TYPE
  UDMDT_FIELD_DETAILS = 1
  UDMDT_DATA_SETS
  UDMDT_RECORD
End Enum

Private Enum UDM_RECORD_DATA
  UDMRD_MOVENEXT = 1
  UDMRD_MOVEFIRST
  UDMRD_RECORD_COUNT
End Enum

Private m_BCLast As BEN_CLASS
Private m_sLastDataSetRequested As String
Private m_RecordPointers(BC_FIRST_ITEM To [BC_REAL_BENEFITS_LAST_ITEM]) As Long
Private m_RecordCounts(BC_FIRST_ITEM To [BC_REAL_BENEFITS_LAST_ITEM]) As Long
Private m_Notify As IBaseNotify
Private Function ITCSUDM_OpenSession(FileName As String, Optional ByVal DestSelectOnly As Boolean = False) As Boolean
  Dim ey As Employer
  Dim ee As IBenefitClass
  Dim i As Long
  
  If App.StartMode <> vbSModeStandalone Then
    Set ey = New Employer
    If Not ey.Validate(FileName) Then Call Err.Raise(ERR_NO_EMPLOYER, "OpenSession", "Unable to open the session for " & FileName & ".")
    p11d32.ReportPrint.EmployeeSelection = ES_ALL
    Call p11d32.LoadEmployer(ey, False)
    Call Initialise
    m_RecordPointers(BC_EMPLOYER) = p11d32.Employers.Add(ey)
  End If
  ITCSUDM_OpenSession = True
End Function

Private Sub ITCSUDM_CloseSession()
  If App.StartMode <> vbSModeStandalone Then
    If Not p11d32.CurrentEmployer Is Nothing Then
      Call p11d32.LoadEmployer(Nothing)
      Call p11d32.Employers.RemoveAll
    End If
  End If
End Sub
Private Function ITCSUDM_AppendValues(ByVal DestFields As Variant, ByVal DestValues As Variant) As TCSBASE.OLE_ERRORTYPE

End Function

Private Property Get ITCSUDM_BOF() As Boolean
  Dim bc As BEN_CLASS
  
  bc = CurrentBenClass
  Call TableMovementData(UDMRD_RECORD_COUNT, bc)
  If m_RecordCounts(bc) = 0 Then
    ITCSUDM_BOF = True
  Else
    ITCSUDM_BOF = (m_RecordPointers(bc) = L_DB_BOF)
  End If
End Property

Private Property Let ITCSUDM_CurrentDataSet(RHS As Variant)
  m_CurrentDataSet = RHS
End Property

Private Property Get ITCSUDM_CurrentDataSet() As Variant
  ITCSUDM_CurrentDataSet = m_CurrentDataSet
End Property


Private Property Get ITCSUDM_Description() As String
  ITCSUDM_Description = "P11D UDM"
End Property

Private Function CurrentBenClass() As BEN_CLASS
   CurrentBenClass = BenClassFromDataSet(CStr(m_CurrentDataSet(UBound(m_CurrentDataSet))))

End Function

Private Function BenClassFromDataSet(sUBoundDataSet As String) As BEN_CLASS
  Dim i As Long
  
  On Error GoTo BenClassFromDataSet_ERR
  Call xSet("BenClassFromDataSet")
  If Len(sUBoundDataSet) = 0 Then ECASE ("String passed to get ben clas from is zero length.")
  
  If StrComp(m_sLastDataSetRequested, sUBoundDataSet) = 0 Then
    BenClassFromDataSet = m_BCLast
    GoTo BenClassFromDataSet_END
  End If
  
  For i = 1 To BC_LAST_ITEM
    If StrComp(p11d32.Rates.BenClassTo(i, BCT_FORM_CAPTION), sUBoundDataSet) = 0 Then
      BenClassFromDataSet = i
      m_BCLast = i
      m_sLastDataSetRequested = sUBoundDataSet
      Exit For
    End If
  Next
  
  If BenClassFromDataSet < BC_FIRST_ITEM Then ECASE ("Table offset not found")
 
BenClassFromDataSet_END:
  Call xReturn("BenClassFromDataSet")
  Exit Function
BenClassFromDataSet_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BenClassFromDataSet", "Ben Class From Data Set", "Error getting the ben class from a data set.")
  Resume BenClassFromDataSet_END
End Function

Private Property Get ITCSUDM_EOF() As Boolean
  Dim bc As BEN_CLASS
  
  bc = CurrentBenClass
  Call TableMovementData(UDMRD_RECORD_COUNT, bc) 'set BOF flag
  If m_RecordCounts(bc) = 0 Then
    ITCSUDM_EOF = True
  Else
    ITCSUDM_EOF = (m_RecordPointers(bc) = L_DB_EOF)
  End If
End Property
Private Function ITCSUDM_GetErrorDescription(ByVal ErrorCode As TCSBASE.OLE_ERRORTYPE) As String

End Function
Private Function UDMFieldDetails(ByVal bc As BEN_CLASS) As Variant
  
  On Error GoTo UDMFieldDetails_ERR
  
  Call xSet("UDMFieldDetails")
  
  If p11d32.BenDataLinkUDMUBound(bc) < 1 Then
    Call Err.Raise(ERR_BEN_HAS_NO_UDM_FIELDS, "UDMFieldDetails", "The benefit class has no field details.")
  End If
  UDMFieldDetails = p11d32.BenDataLinkUDMFieldDetails(bc)
  
UDMFieldDetails_END:
  Call xReturn("UDMFieldDetails")
  Exit Function
UDMFieldDetails_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UDMFieldDetails", "UDM Field Details", "Error getting the UDM field details for benclass = " & bc & ".")
  Resume UDMFieldDetails_END
  Resume
End Function
Private Function GetUDMData(ByVal UDMDT As UDM_DATA_TYPE, Optional bc As BEN_CLASS = -1) As Variant
  Dim v As Variant
  Dim i As Long, j As Long
  
  On Error GoTo GetUDMData_ERR
  
  Select Case UDMDT
    Case UDMDT_DATA_SETS
      Select Case bc
        Case -1
          ReDim v(1 To 1)
          v(1) = p11d32.Rates.BenClassTo(BC_EMPLOYER, BCT_FORM_CAPTION)
          
          p11d32.BenDataLinkAddUDMTableOffset (BC_EMPLOYER)
        Case Else
          Select Case bc
            Case BC_EMPLOYER
              ReDim v(1 To 1)
              v(1) = p11d32.Rates.BenClassTo(BC_EMPLOYEE, BCT_FORM_CAPTION)
              p11d32.BenDataLinkAddUDMTableOffset (BC_EMPLOYEE)
            Case BC_EMPLOYEE
              ReDim v(BC_FIRST_ITEM To p11d32.ReportPrint.BenClassLastUDM)
              j = BC_FIRST_ITEM - 1
              For i = BC_FIRST_ITEM To p11d32.ReportPrint.BenClassLastUDM
                 j = j + 1
                 v(j) = p11d32.Rates.BenClassTo(i, BCT_FORM_CAPTION)
                 Call p11d32.BenDataLinkAddUDMTableOffset(j)
              Next i
          End Select
      End Select
    Case UDMDT_FIELD_DETAILS
      GetUDMData = p11d32.BenDataLinkUDMFieldDetails(bc)
      v = p11d32.BenDataLinkUDMFieldDetails(bc)
    Case UDMDT_RECORD
      
  End Select
  
  If UDMDT = UDMDT_DATA_SETS Then GetUDMData = v
  
GetUDMData_END:
  Exit Function
GetUDMData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetUDMData", "Get UDM Data", "Error getting record data via the UDM.")
  Resume GetUDMData_END
  Resume
End Function
Private Function DebugFields(v)
  Dim IFD As FieldDetails
  Dim i As Long
  For i = LBound(v) To UBound(v)
    Set IFD = v(i)
    'Debug.Print IFD.Name & vbCrLf
  Next
End Function

Private Function ITCSUDM_GetFieldDetails(FieldDetailsArray As Variant) As TCSBASE.OLE_ERRORTYPE
  
  On Error GoTo GetFieldDetails_err
  Call xSet("GetFieldDetails")
  
  FieldDetailsArray = GetUDMData(UDMDT_FIELD_DETAILS, CurrentBenClass)
  
GetFieldDetails_end:
  Call xReturn("GetFieldDetails")
  Exit Function
GetFieldDetails_err:
  Call ErrorMessage(ERR_ERROR, Err, "GetFieldDetails", "Get Field Details", "Error in UDM getting field details")
  Resume GetFieldDetails_end
  Resume
End Function

Private Function ITCSUDM_GetValues(SourceFields As Variant, SourceValues As Variant, Optional KeyID As Variant) As TCSBASE.OLE_ERRORTYPE
  Dim i As Long
  Dim bc As BEN_CLASS
  Dim ben As IBenefitClass
  Dim Employer As Employer
  Dim Employee As Employee
  Dim loans As loans
  Dim v As Variant
  
  On Error GoTo ITCSUDM_GetValues_ERR
  
  bc = CurrentBenClass
  
  SourceFields = p11d32.BenDataLinkUDMFields(bc)
  
  Select Case bc
    Case BC_EMPLOYER
      Set ben = p11d32.Employers(m_RecordPointers(bc))
    Case BC_EMPLOYEE
      Set ben = GetCurrentEmployee
    Case Else
      Set Employee = GetCurrentEmployee
      
      If BenefitIsLoan(bc) Then
        'bloody loans again
        i = Employee.GetLoansBenefitIndex
        If i > 0 Then
          Set loans = Employee.benefits(i)
          If loans.LoansCount Then
            Set ben = loans.loans(m_RecordPointers(bc))
          End If
        End If
      Else
        Set ben = Employee.benefits(m_RecordPointers(bc))
      End If
  End Select
  
  'use array already sized correctly
  SourceValues = p11d32.BenDataLinkUDMFields(bc)
  
  If ben Is Nothing Then ECASE ("Ben is nothing in UDM GetValues")
  
  For i = 1 To p11d32.BenDataLinkUDMUBound(bc)
    'SO
    If p11d32.BenDataLinkUDMRequiresCalculate(bc, i) Then
      Call ben.Calculate
    End If
    
    SourceValues(i) = ben.value(p11d32.BenDataLinkUDMFieldID(bc, i))
  Next
  
ITCSUDM_GetValues_END:
  Exit Function
ITCSUDM_GetValues_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ITCSUDM_GetValues", "ITCSUDM Get Values", "Error getting values for UDM interface.")
  Resume ITCSUDM_GetValues_END
  Resume
End Function

Private Function ITCSUDM_IsGetDetailsAvailable() As Boolean

End Function

Private Function ITCSUDM_IsSetDetailsAvailable() As Boolean

End Function

Private Sub ITCSUDM_MoveFirst()
  Call TableMovementData(UDMRD_MOVEFIRST, CurrentBenClass)
End Sub
  
Private Sub ITCSUDM_MoveNext()
  Call TableMovementData(UDMRD_MOVENEXT, CurrentBenClass)
End Sub

Private Property Get ITCSUDM_Name() As String
  ITCSUDM_Name = AppName
End Property

Public Property Set ITCSUDM_Notify(RHS As TCSBASE.IBaseNotify)
  Set m_Notify = RHS
End Property

Public Property Get ITCSUDM_Notify() As TCSBASE.IBaseNotify
  Set ITCSUDM_Notify = m_Notify
End Property

Public Function Initialise()
  Dim i As Long
   
  'CAD need to open employer on file name not for now call p11d32.vlidateEx
  
  For i = BC_FIRST_ITEM To BC_REAL_BENEFITS_LAST_ITEM
    m_RecordPointers(i) = L_DB_BOF '-1 tells if EOF 0 for BOF
    m_RecordCounts(i) = L_DB_RECORD_COUNT_NOT_INITALISED
  Next i
End Function
Private Function ITCSUDM_QueryAllSubDataSets(DataSetArray As Variant) As Variant

  On Error GoTo QueryAllSubDataSets_ERR
  
  Call xSet("QueryAllSubDataSets")
  
  ITCSUDM_QueryAllSubDataSets = GetUDMData(UDMDT_DATA_SETS, BenClassFromDataSet(CStr(DataSetArray(UBound(DataSetArray)))))
  
QueryAllSubDataSets_END:
  Call xReturn("QueryAllSubDataSets")
  Exit Function
QueryAllSubDataSets_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "QueryAllSubDataSets", "Query All Sub Data Sets", "Error querying all sub data sets in UDM.")
  Resume QueryAllSubDataSets_END
  Resume
End Function

Private Function ITCSUDM_QueryAllTopLevelDataSets() As Variant
  ITCSUDM_QueryAllTopLevelDataSets = GetUDMData(UDMDT_DATA_SETS, -1)
End Function

Private Property Get ITCSUDM_RecordCount() As Long
  ITCSUDM_RecordCount = TableMovementData(UDMRD_RECORD_COUNT, CurrentBenClass)
End Property
Private Function GetCurrentEmployer() As Employer
  On Error GoTo GetCurrentEmployer_ERR
  
  Call xSet("GetCurrentEmployer")
  
  If ValidRecord(BC_EMPLOYER) Then
    If m_RecordPointers(BC_EMPLOYER) = L_DB_BOF Then
      Call TableMovementData(UDMRD_MOVENEXT, BC_EMPLOYER)
      Set GetCurrentEmployer = p11d32.Employers(m_RecordPointers(BC_EMPLOYER))
      m_RecordPointers(BC_EMPLOYER) = L_DB_BOF
    Else
      Set GetCurrentEmployer = p11d32.Employers(m_RecordPointers(BC_EMPLOYER))
    End If
  End If
  
GetCurrentEmployer_END:
  Call xReturn("GetCurrentEmployer")
  Exit Function
GetCurrentEmployer_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetCurrentEmployer", "Get Current Employer", "Error getting the current employer for the UDM.")
  Resume GetCurrentEmployer_END
End Function
Private Function GetEmployeeFromCurrentRecordPointer(Employer As Employer) As Employee
  Dim ibf As IBenefitForm2
  If App.StartMode = vbSModeStandalone Then
    Set ibf = F_Employees
    Set GetEmployeeFromCurrentRecordPointer = Employer.employees(ibf.lv.listitems(m_RecordPointers(BC_EMPLOYEE)).Tag)
  Else
    Set GetEmployeeFromCurrentRecordPointer = Employer.employees(m_RecordPointers(BC_EMPLOYEE))
  End If
End Function

Private Function GetCurrentEmployee() As Employee
  Dim Employer As Employer
  
  On Error GoTo GetCurrentEmployee_ERR
  
  Call xSet("GetCurrentEmployee")
  Set Employer = GetCurrentEmployer
  If Not Employer Is Nothing Then
    If ValidRecord(BC_EMPLOYEE) Then
      If m_RecordPointers(BC_EMPLOYEE) = L_DB_BOF Then
        Call TableMovementData(UDMRD_MOVENEXT, BC_EMPLOYEE)
        Set GetCurrentEmployee = GetEmployeeFromCurrentRecordPointer(Employer)
        m_RecordPointers(BC_EMPLOYER) = L_DB_BOF
      Else
        Set GetCurrentEmployee = GetEmployeeFromCurrentRecordPointer(Employer)
      End If
      Call GetCurrentEmployee.LoadBenefits(TBL_ALLBENEFITS, False)
    End If
  End If
  
  
GetCurrentEmployee_END:
  Call xReturn("GetCurrentEmployee")
  Exit Function
GetCurrentEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetCurrentEmployee", "Get Current Employee", "Error getting the current employee for the UDM.")
  Resume GetCurrentEmployee_END
  Resume
End Function
Private Sub MoveNextBenefit(ByVal bc As BEN_CLASS)
  Dim i As Long
  Dim Employer As Employer
  Dim Employee As Employee
  Dim OL As ObjectList
  Dim ben As IBenefitClass
  Dim b As Boolean
  Dim loans As loans
  Dim sError As String
  
  On Error GoTo MoveNextBenefit_ERR
  
  Call xSet("MoveNextBenefit")
  
  Set Employer = GetCurrentEmployer
  If Not Employer Is Nothing Then
    Set Employee = GetCurrentEmployee
    If Not Employee Is Nothing Then
      If m_RecordPointers(bc) <> L_DB_EOF Then
        If BenefitIsLoan(bc) Then
          'bloody loans again
          i = Employee.GetLoansBenefitIndex
          If i > 0 Then
            Set loans = Employee.benefits(i)
            Set OL = loans.loans
          End If
        Else
          Set OL = Employee.benefits
        End If
        If Not OL Is Nothing Then
          For i = m_RecordPointers(bc) + 1 To OL.Count
            Set ben = OL(i)
            If Not ben Is Nothing Then
              If ben.BenefitClass = bc Then
                Call ben.Calculate
                If ben.LinkBen And (Not ben.value(ITEM_BENEFIT_REPORTABLE)) Then GoTo NEXT_BEN
                b = True
                Exit For
              End If
            End If
NEXT_BEN:
          Next
        End If
      End If
    End If
  End If
  If b Then
    m_RecordPointers(bc) = i
  Else
    m_RecordPointers(bc) = L_DB_EOF
  End If
  
MoveNextBenefit_END:
  Call xReturn("MoveNextBenefit")
  Exit Sub
MoveNextBenefit_ERR:
  Call ErrorMessagePush(Err)
  sError = "Error moving to the next benefit via the UDM, ben class = " & p11d32.Rates.BenClassTo(bc, BCT_UDM_BENEFIT_TITLE) & "."
  If Not ben Is Nothing Then
    
    sError = sError & " Item = " & ben.value(ITEM_DESC)
    If Not Employee Is Nothing Then
      sError = sError & ", employee reference = " & Employee.PersonnelNumber
    End If
  End If
  Call ErrorMessagePopErr(Err)
  Call Err.Raise(Err.Number, ErrorSource(Err, "MoveNextBenefit"), sError)
  'Call ErrorMessage(ERR_ERROR, Err, "MoveNextBenefit", "Move Next Benefit", sError)
  Resume MoveNextBenefit_END
  Resume
End Sub

Private Sub KillBenefits()
  Dim ee As Employee
  Dim ey As Employer
  
  On Error GoTo KillBenefits_ERR
  
  If Not p11d32.KillBenefits Then GoTo KillBenefits_END ' not required in ee.KillBenfitsEx

  Set ey = GetCurrentEmployer
  If Not ey Is Nothing Then
    If ValidRecord(BC_EMPLOYEE) And m_RecordPointers(BC_EMPLOYEE) <> L_DB_BOF Then
      Set ee = GetEmployeeFromCurrentRecordPointer(ey)
      If Not ee Is Nothing Then ee.KillBenefitsEx
    End If
  End If
  
KillBenefits_END:
  Exit Sub
KillBenefits_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "KillBenefits", "Kill Benefits", "Error killing employee benefits inside UDM.")
  Resume KillBenefits_END
  Resume
End Sub
Private Sub MoveNextEmployee()
  Dim i As Long
  Dim Employer As Employer
  Dim b As Boolean
  Dim ibf As IBenefitForm2
  Dim li As ListItem
  Dim ben As IBenefitClass
  
  On Error GoTo MoveNextEmployee_ERR
  Call xSet("MoveNextEmployee")
  Set ibf = F_Employees
  
  'StartIndex of 1 denotes move first
  Set Employer = GetCurrentEmployer
  If Not Employer Is Nothing Then
    If ValidRecord(BC_EMPLOYEE) Then
      If App.StartMode = vbSModeStandalone Then
        For i = m_RecordPointers(BC_EMPLOYEE) + 1 To ibf.lv.listitems.Count 'change to lb to obtain ordering??
          Set li = ibf.lv.listitems(i)
          If p11d32.EmployeeSelectedForPrint(Employer.employees(li.Tag), Employer) Then
            Set ben = Employer.employees(li.Tag)
            b = True
            Exit For
          End If
        Next
      Else
        For i = m_RecordPointers(BC_EMPLOYEE) + 1 To Employer.employees.Count
          Set ben = Employer.employees(i)
          If Not ben Is Nothing Then
            b = True
            Exit For
          End If
        Next
      End If
    End If
  End If
  
  If b Then
    m_RecordPointers(BC_EMPLOYEE) = i
  Else
    m_RecordPointers(BC_EMPLOYEE) = L_DB_EOF
  End If
  
MoveNextEmployee_END:
  Call xReturn("MoveNextEmployee")
  Exit Sub
MoveNextEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "MoveNextEmployee", "Move Next Employee", "Error moving to the next employee via the UDM.")
  Resume MoveNextEmployee_END
  Resume
End Sub
Private Sub ResetBenefitsData()
  Dim i As Long
  
  On Error GoTo ResetBenefitsData_ERR
  
  Call xSet("ResetBenefitsData")
  
  For i = BC_FIRST_ITEM To p11d32.ReportPrint.BenClassLastUDM
    Call ResetBenefitData(i)
  Next
  
ResetBenefitsData_END:
  Call xSet("ResetBenefitsData")
  Exit Sub
ResetBenefitsData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ResetBenefitsData", "Reset Benefits Data", "Error resetting benefit data for the current UDM employee.")
  Resume ResetBenefitsData_END
End Sub
Private Sub ResetBenefitData(bc As BEN_CLASS)
  m_RecordPointers(bc) = L_DB_BOF
  m_RecordCounts(bc) = L_DB_RECORD_COUNT_NOT_INITALISED
End Sub

Private Function TableMovementData(ByVal UDMRD As UDM_RECORD_DATA, bc As BEN_CLASS) As Long
  Dim Employer As Employer
  Dim Employee As Employee
  Dim ben As IBenefitClass
  Dim loans As loans
  Dim i As Long
  Dim b As Boolean
  Dim OL As ObjectList
  
  On Error GoTo TableMovementData_ERR
  Select Case bc
    Case BC_EMPLOYER
      Select Case UDMRD
        Case UDM_RECORD_DATA.UDMRD_RECORD_COUNT
          If m_RecordCounts(bc) = L_DB_RECORD_COUNT_NOT_INITALISED Then
            m_RecordCounts(bc) = 1
          End If
          TableMovementData = m_RecordCounts(bc)
        Case UDM_RECORD_DATA.UDMRD_MOVEFIRST
          m_RecordPointers(BC_EMPLOYEE) = L_DB_BOF
          m_RecordPointers(bc) = L_DB_BOF
        Case UDM_RECORD_DATA.UDMRD_MOVENEXT
          If ValidRecord(bc) Then 'need to call record count
            If m_RecordPointers(bc) = L_DB_BOF Then
              m_RecordPointers(bc) = p11d32.Employers.ItemIndex(p11d32.CurrentEmployer)
              m_RecordPointers(BC_EMPLOYEE) = L_DB_BOF 'reset employees
            Else
              m_RecordPointers(bc) = L_DB_EOF
              m_RecordCounts(bc) = L_DB_RECORD_COUNT_NOT_INITALISED
            End If
          End If
      End Select
    Case BC_EMPLOYEE
      Select Case UDMRD
        'employees
        Case UDMRD_RECORD_COUNT
          If m_RecordCounts(bc) = L_DB_RECORD_COUNT_NOT_INITALISED Then
            m_RecordCounts(bc) = 0
            Set Employer = GetCurrentEmployer
            If Not Employer Is Nothing Then
              m_RecordCounts(bc) = p11d32.ReportPrint.NoOfEmployeesSelected(Employer)
            Else
              m_RecordCounts(bc) = 0
            End If
          End If
          TableMovementData = m_RecordCounts(bc)
        Case UDMRD_MOVEFIRST
          Call ResetBenefitsData
          m_RecordPointers(bc) = L_DB_BOF
       Case UDMRD_MOVENEXT
          Call ResetBenefitsData
          Call KillBenefits
          Call MoveNextEmployee
      End Select
    Case Else
      Select Case UDMRD
        Case UDMRD_RECORD_COUNT
          If m_RecordCounts(bc) = L_DB_RECORD_COUNT_NOT_INITALISED Then
            m_RecordCounts(bc) = 0
            Set Employer = GetCurrentEmployer
            If Not Employer Is Nothing Then
              Set Employee = GetCurrentEmployee
              If Not Employee Is Nothing Then
                If BenefitIsLoan(bc) Then
                  'bloody loans again
                  i = Employee.GetLoansBenefitIndex
                  If i > 0 Then
                    Set loans = Employee.benefits(i)
                    Set OL = loans.loans
                  End If
                Else
                  Set OL = Employee.benefits
                End If
                If Not OL Is Nothing Then
                  For i = 1 To OL.Count
                    Set ben = OL(i)
                    If Not ben Is Nothing Then
                      If ben.BenefitClass = bc Then
                        m_RecordCounts(bc) = m_RecordCounts(bc) + 1
                      End If
                    End If
                  Next
                End If
              End If
            End If
            'deal with loans if loans then take of one and add numebr of loans
          End If
          TableMovementData = m_RecordCounts(bc)
        Case UDMRD_MOVEFIRST
          m_RecordPointers(bc) = L_DB_BOF
        Case UDMRD_MOVENEXT
          Call MoveNextBenefit(bc)
      End Select
      

  End Select
  
  
TableMovementData_END:
  
  Exit Function
TableMovementData_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "TableMovementData"), Err.Description)
  'Call ErrorMessage(ERR_ERROR, Err, "TableMovementData", "Table Movement Data", "Error manipulating the UDM table data.")
  Resume TableMovementData_END
  Resume
End Function
Private Function ValidRecord(ByVal bc As BEN_CLASS) As Boolean
  If TableMovementData(UDMRD_RECORD_COUNT, bc) > 0 Then
    If Not EOF(bc) Then
      ValidRecord = True
    End If
  End If
End Function
Private Function EOF(ByVal bc As BEN_CLASS)
  EOF = (m_RecordPointers(bc) = L_DB_EOF)
End Function
Private Function ITCSUDM_SetValues(ByVal DestFields As Variant, ByVal DestValues As Variant, Optional KeyID As Variant) As TCSBASE.OLE_ERRORTYPE

End Function

Private Property Get ITCSUDM_Version() As String

End Property
