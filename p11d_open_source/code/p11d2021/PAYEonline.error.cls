VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PAYEonline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'hard copy forms, hmrc.substituteformsapproval@hmrc.gsi.gov.uk


'PAS download
'http://www.p11dorganiser.co.uk/download/setup/setup.exe

'Serial Number: P11DEVALUATE
'Company Name: P11D Evaluation Company Ltd
'Initial PassWord: E40EBECD19
'created user TEST, password ytrewq
 


'contact for SDS team
'01274 539666
'old sdsteam@ir.gsi.gov.uk
'new sdsteam@ hmrc.gsi.gov.uk

Option Explicit
Implements IErrorFilter
Implements IBaseNotify

Public SubReturn As Long

Public ByPassLocalSchemaCheck As Boolean

Public ViewProceedButtonIfErrors As Boolean
Public SubReturnOf As Long
Public DataFormat As MM_DATA_FORMAT
Public ErrorCount As Long
Public AtLeastOneError As Boolean
Public InvalidNICount As Long
Public LastPathAndFileCreated As String
Public ErrorLogging As Boolean
Public ExtraSubmissionPropertiesMenu As Boolean
Public DisableCheckSubmissionWithAbatec As Boolean
Public Efiler_Test_Submission As Boolean
Public Efiler_Mock_Submission As Boolean
Public Efiler_Proceed_Submission As Boolean

Private m_FullNamesSpace As String
Private m_rsError As Recordset
Private m_DBPath As String
Private m_db As Database
Private m_ws As Workspace
Private m_OutputDirectory As String
Private m_Prg As TCSPROG.TCSProgressBar
'Public XML_doc As DOMDocument
'Public XML_main_node As IXMLDOMNode
'Public XMLEmployerHeaderNode As IXMLDOMNode
'Public XMLBenefitNode As IXMLDOMNode

Private P46Cars As ObjectList

Public P46DateFrom As Date 'p11d32.ReportPrint.P46DateFrom
Public P46DateTo As Date 'p11d32.ReportPrint.
Public P46UserDateFrom As Variant 'p11d32.ReportPrint.
Public P46UserDateTo As Variant 'etc.
Public P46Range As P46_RANGE

Public Type_Selected As PAYEonline_TYPES
'Public EfilerCom As atc2efiler.Efiler
Public SubmitterRef As String
Public SubmitterName As String
Private m_sSubmitter_Ref As String
Private m_sSubmitter_ID As String
Private m_sSubmitter_Pwd As String
Private m_sSubmitter_Email As String
Private m_sSubmitter_User As String
Private m_Entity_ID As String
Private m_Entity_Name As String
Private m_sTaxOfficeNumber As String
Private m_sTaxOfficeReference As String
Private m_SubmissionCancelled As Boolean

Public Property Let SubmissionCancelled(ByVal NewValue As Boolean)
  m_SubmissionCancelled = NewValue
End Property
  
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
  'RK Check that this will not cause issues where DBUsername is set!
  Set m_ws = DBEngine.CreateWorkspace("ErrorWS", p11d32.DBUserName, p11d32.DBPassword)
  Set m_db = OpenDB(m_ws, m_DBPath, False)
End Sub
Public Function Kill()
  Set m_db = Nothing
  Set m_ws = Nothing
End Function

Private Property Get ReportHeader() As String
  On Error GoTo ReportHeader_ERR
  
  Call xSet("ReportHeader")
  
  ReportHeader = "PAYE Online validation warnings and errors" & vbCrLf & vbCrLf
  
ReportHeader_END:
  Call xReturn("ReportHeader")
  Exit Property
ReportHeader_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportHeader", "Report Header", "Error getting the report header for import errors report.")
  Resume ReportHeader_ERR
End Property

Private Sub Class_Initialize()
  ErrorLogging = True
  Efiler_Test_Submission = False
  Efiler_Mock_Submission = False
  Efiler_Proceed_Submission = True
End Sub

Private Sub IBaseNotify_Notify(ByVal Current As Long, ByVal Max As Long, ByVal Message As String)

  If (m_SubmissionCancelled) Then
    Call Err.Raise(ERR_PAYE_ONLINE_CANCEL, "Notify", "Submission cancelled")
  End If
    
  If Current = -1 And Max = -1 Then
    If ((Not ByPassLocalSchemaCheck) Or IsRunningInIDE()) Then
      m_Prg.Caption = "Validating data against local schema"
      If (Me.Type_Selected = POT_P46 And Not p11d32.JuneRelease) Then
        Call ValidateXMLToRevenueLocal(Message, p11d32.SystemPAYEOnlineEXBPriorXSDPath)
      Else
        Call ValidateXMLToRevenueLocal(Message, p11d32.SystemPAYEOnlineEXBXSDPath)
      End If
    End If
  Else
    If (m_Prg.value = m_Prg.Max) Then
      m_Prg.value = m_Prg.Min
      m_Prg.Caption = Message
    Else
      Call m_Prg.StepCaption(Message)
    End If
  End If
  DoEvents

  
End Sub

Private Function IErrorFilter_FilterErrorMessage(ByVal Username As String, ByVal DateTime As Date, ByVal ErrorNumber As Long, ByVal ErrorName As String, ByVal ErrorText As String, ByVal SourceFunction As String) As Boolean
  Dim s() As String
  Dim lEqualsPos As Long
  
  On Error GoTo FilterError_err
   
  m_rsError.AddNew
  
  If ErrorNumber = 0 Then
    ECASE ("Error number in PAYE Online error filter is 0")
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

Private Function EmployersSelected(Employers As ObjectList) As Long
'  'Returns list of Employers where PAYEOnline is selected
  Dim i As Long
  Dim benEmployer As IBenefitClass
  Dim ey As Employer
  
  Set Employers = New ObjectList
  For i = 1 To p11d32.Employers.Count
    Set ey = p11d32.Employers(i)
    Set benEmployer = ey
    If Not benEmployer Is Nothing Then
      If benEmployer.value(employer_PAYEOnlineSelected) Then
        Call Employers.Add(ey)
      End If
    End If
  Next
    
  EmployersSelected = Employers.Count
End Function


Public Sub Errors(ByVal Dest As REPORT_TARGET, lPAYEOnlineErrorType As VIEW_ERRORS_TYPE, Optional SubIDs As String)
  
  On Error GoTo Errors_ERR
  
  Call xSet("Errors")

  Call OpenErrorLogDB
  If lPAYEOnlineErrorType = VET_PAYEONLINE_VALIDATION Then
    Call StartAutoSTD(sql.Queries(SELECT_PAYEONLINE_VALIDATION_ERRORS), m_db, ReportHeader, Dest)
  ElseIf lPAYEOnlineErrorType = VET_PAYEONLINE_SUBMISSION Then
    Call StartAutoSTD(sql.Queries(SELECT_PAYEONLINE_SUBMISSION_ERRORS) & SubIDs & " AND Status like '%Error%'", m_db, ReportHeader, Dest)
  End If
  
Errors_END:
  Call CloseErrorLogDB
  Call ClearAllCursors
  Call xReturn("Errors")
  Exit Sub
Errors_ERR:
  If Not m_db Is Nothing Then
'MP Rv - changed mm ref to PAYE ref
'MP    Call ErrorMessage(ERR_ERROR, Err, "Errors", "Errors", "Error viewing/printing the magnetic media errors from " & m_db.Name & ".")
    Call ErrorMessage(ERR_ERROR, Err, "Errors", "Errors", "Error viewing/printing the PAYE submission errors from " & m_db.Name & ".")
  Else
'MP    Call ErrorMessage(ERR_ERROR, Err, "Errors", "Errors", "Error viewing/printing the magnetic media errors from ")
    Call ErrorMessage(ERR_ERROR, Err, "Errors", "Errors", "Error viewing/printing the PAYE submission errors. ")
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
Public Property Get OutputDirectory() As String
  OutputDirectory = m_OutputDirectory
End Property
Public Property Let OutputDirectory(sNewValue As String)
  m_OutputDirectory = FullPath(sNewValue)
End Property
Public Sub Start()
  Dim i As Long
  Dim emp As Employer
  Dim li As ListItem
  
  On Error GoTo Start_ERR
  Call xSet("Start")
  
  Call EmployersToListView
  
  F_PayeOnline.Caption = "PAYE Online " & p11d32.AppYear & "/" & p11d32.AppYear + 1   'RH
  Call p11d32.Help.ShowForm(F_PayeOnline, vbModal)
  
Start_END:
  Set p11d32.CurrentEmployer = Nothing
  Set F_PayeOnline = Nothing
  Call xReturn("Start")
  Exit Sub
  
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error starting PAYE Online.")
  Resume Start_END
End Sub

Public Sub EmployersToListView()
  Dim i As Long
  Dim ben As IBenefitClass
  Dim li As ListItem

  On Error GoTo EmployersToListView_ERR
  
  Call xSet("EmployersToListView")

  For i = 1 To p11d32.Employers.Count
    Set ben = p11d32.Employers(i)
    If Not ben Is Nothing Then
      Set li = F_PayeOnline.lvPAYEEmployers.listitems.Add()
      Call UpdateListViewItem(li, ben, i)
    End If
  Next
   
  Call F_PayeOnline.lvPAYEEmployers_ItemCheck(Nothing)
 
   
EmployersToListView_END:
  Call xReturn("EmployersToListView")
  Exit Sub
EmployersToListView_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployersToListView", "Employers To List View", "Error placing the current employers to the Magnetic Media list view.")
  Resume EmployersToListView_END
  Resume
End Sub


Public Function UpdateListViewItem(li As ListItem, ben As IBenefitClass, Optional ByVal BenefitIndex As Long = 0) As Boolean
  Dim s As String
  
  On Error GoTo UpdateListViewItem_ERR
  
  Call xSet("UpdateListViewItem")
  
  Dim emp As Employer
  
  If li Is Nothing Then Call Err.Raise(ERR_ERROR, "UpdateListViewItem", "The list item is nothing")
  If ben Is Nothing Then Call Err.Raise(ERR_ERROR, "UpdateListViewItem", "The beneift is nothing")
  
  If BenefitIndex > 0 Then li.Tag = BenefitIndex
  
  Set emp = ben
  
  If Not emp.PAYEOnlineValid(False) Then
    li.SmallIcon = IMG_EXCLAMATION
  Else
    li.SmallIcon = IMG_OK
  End If
  
  s = ben.Name
  If p11d32.LicenceType = LT_DEMO Then
    s = s & " (" & ben.value(employer_FileName) & ")"
  End If
      
  li.Text = s
  
  li.SubItems(1) = ben.value(employer_Payeref_db)
  li.SubItems(2) = IIf(emp.PayeRefValid, "Yes", "No")
  li.SubItems(3) = ben.value(employer_EmployeesCount)
  li.SubItems(4) = ben.value(employer_PayeOnlineID_db)
  li.SubItems(5) = ben.value(employer_PayeOnlinePwd_db)
  li.SubItems(6) = ben.value(employer_PayeOnlineUser_db)
  
  UpdateListViewItem = True
  
UpdateListViewItem_END:
  Call xReturn("UpdateListViewItem")
  Exit Function
  
UpdateListViewItem_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateListViewItem", "Update List View Item", "Error Updating List Item")
  Resume UpdateListViewItem_END
End Function
Public Function UploadTime(iBytes As Long) As String
  Dim s As String
  Dim i As Long
  
  iBytes = iBytes * 8
  i = iBytes / 56000
  s = s & "56K Modem:" & i & " seconds" & vbCrLf
  i = iBytes / 500000
  s = s & "ADSL(0.5 MBit):" & i & " seconds" & vbCrLf
  i = iBytes / 1000000
  s = s & "ADSL(1 MBit):" & i & " seconds" & vbCrLf
  i = iBytes / 2000000
  s = s & "T1 Line:" & i & " seconds" & vbCrLf
  
  UploadTime = s
End Function

Private Function CreateP46Node(XML_doc As DOMDocument60, XMLHeaderNode As IXMLDOMNode, ben_ee As IBenefitClass) As IXMLDOMNode
  Dim n As IXMLDOMNode
  Set n = XML_doc.createElement("P46Car")
  Call CreateEmployeeXML(XML_doc, n, ben_ee, POT_P46)
  XMLHeaderNode.appendChild n
  Set CreateP46Node = n
End Function
Public Sub CreatePAYEFiles(ByVal prgEmployee As TCSProgressBar, ByVal lblFile As Label)
  Dim ey As Employer, ee As Employee
  Dim ben As IBenefitClass
  Dim ben_ee As IBenefitClass
  Dim ben_car As IBenefitClass
  Dim Employers As ObjectList
  Dim i As Long, j As Long, k As Long
  Dim lEmployeesWritten As Long
  Dim sBenefits As String, sEmployeeRec As String
  Dim XMLBenefitNode As IXMLDOMNode, XMLHeaderNode As IXMLDOMNode, xmlNode As IXMLDOMNode, XML_doc As DOMDocument60
  Dim XML_employeeheader As IXMLDOMNode
  Dim root As String
  Dim iReportedCarCount As Long
  Dim node_p11d_record_count As IXMLDOMNode, node_p46_record_count As IXMLDOMNode, node_declarations As IXMLDOMNode
  Dim sLastFileName As String, sSubmissionID As String
  Dim bSubmit As Boolean, bInvalidNI As Boolean
  Dim Response_doc As String
  Dim ef As EFiler
      
  On Error GoTo CreatePAYEFiles_ERR
  
  Call xSet("CreatePAYEFiles")
  ' Call Initialise_XML_doc
    
  Set m_Prg = prgEmployee
  
  'reinit to prevent efiler retaining properties  , really efiler should have a clear method
  Set ef = NewEfiler()
  
  
  m_Prg.Visible = True
  m_Prg.value = m_Prg.Min
  m_Prg.Indicator = None
  m_Prg.Caption = ""
      
   m_Entity_ID = ""
   m_Entity_Name = ""
   root = ""
   
  'Set up error logging
  ErrorCount = 0
  AtLeastOneError = False
  Call OpenErrorLogDB
  If ErrorLogging Then
    Set ErrorFilter = Me
  Else
    ErrorLogging = True
  End If
  
    Call m_db.Execute(sql.Queries(DELETE_PAYEONLINE_VALIDATION_ERRORS))
    Set m_rsError = m_db.OpenRecordset(sql.Queries(SELECT_PAYEONLINE_VALIDATION_ERRORS), dbOpenDynaset)
  
  Call EmployersSelected(Employers)
  
  If Not ValidateEmployers(ef, Employers) Then
    Me.ViewProceedButtonIfErrors = False
    GoTo ERR_ERRORS:
  Else
    Me.ViewProceedButtonIfErrors = True
  End If
  
  Set XML_doc = New DOMDocument60
  'Call XML_doc.appendChild(XML_doc.createProcessingInstruction("xml", " version=""1.0"" encoding=""UTF-8"""))
  Set XMLHeaderNode = XML_doc.createElement("ExpensesAndBenefits")
  
  
  

  'XML_doc.do
  For i = 1 To Employers.Count
    
    Set ey = Employers(i)
    If i = 1 Then
      Call CreatePAYEEmployerHead(node_p11d_record_count, node_p46_record_count, node_declarations, XML_doc, XMLHeaderNode, Employers)
      XML_doc.appendChild XMLHeaderNode
    End If
    
    Call p11d32.LoadEmployer(ey, False)
    Set ben = ey

    If (i = 1) Then
      ef.companyName = ben.Name
      m_Entity_ID = ben.Name
      m_Entity_Name = ben.Name
    End If
    
    If (ben.value(employer_EmployeesCount) = 0) Then GoTo NEXT_EMPLOYER
      
      ' Public Variables to hold counters for XMLDocument
      m_Prg.Max = ey.employees.Count
      m_Prg.value = 0
      m_Prg.Indicator = ValueOfMax
      
      For j = 1 To ey.employees.Count
        Set ee = ey.employees(j)
        Set ben_ee = ee
        Call m_Prg.StepCaption("Analysing employee: " & ee.FullName)
        If ee Is Nothing Then GoTo NEXT_EMPLOYEE
        If ProceedWithEmployee(ee, P46Cars) Then
          Select Case p11d32.PAYEonline.Type_Selected
            Case POT_P46
              'adding car details
              Set ben_ee = ee
              lEmployeesWritten = lEmployeesWritten + 1
              For k = 1 To P46Cars.Count
                  
                'Each P46 Car is an individual node
                Set ben_car = P46Cars(k)
          
                'THIS IS WRONG BUT REQUIRED AS THE IR do not allow boxes 1 + 2, 2+3 etc to be ticked at the same time
                If ben_car.value(car_P46FirstProvidedWithCar) Then
                  Call CreateXMLCar(XML_doc, CreateP46Node(XML_doc, XMLHeaderNode, ben_ee), ben_car, POT_P46, True)
                  iReportedCarCount = iReportedCarCount + 1
                End If
                If ben_car.value(car_P46CarProvidedReplaced) Then 'back in for 2011 onwards!
                  Call CreateXMLCar(XML_doc, CreateP46Node(XML_doc, XMLHeaderNode, ben_ee), ben_car, POT_P46, , True)
                   iReportedCarCount = iReportedCarCount + 1
                End If
                If ben_car.value(car_P46SecondCar) Then
                  Call CreateXMLCar(XML_doc, CreateP46Node(XML_doc, XMLHeaderNode, ben_ee), ben_car, POT_P46, , , True)
                  iReportedCarCount = iReportedCarCount + 1
                End If
                If ben_car.value(car_P46WithdrawnWithoutReplacement) Then
                  Call CreateXMLCar(XML_doc, CreateP46Node(XML_doc, XMLHeaderNode, ben_ee), ben_car, POT_P46, , , , True)
                  iReportedCarCount = iReportedCarCount + 1
                End If
                
                'END THIS IS WRONG BUT REQUIRED
                
              Next
             
            Case POT_P11D, POT_P11DB_ONLY

              'create node P11D in XML document
              Set XMLBenefitNode = Nothing
              Set XMLBenefitNode = XML_doc.createElement("P11D")
              Set ben_ee = ee
              Call CreateEmployeeXML(XML_doc, XMLBenefitNode, ben_ee, POT_P11D)
              If (Analyse_Benefits(XML_doc, ey, ee, XMLBenefitNode)) Then
                If p11d32.PAYEonline.Type_Selected <> POT_P11DB_ONLY Then
                  XMLHeaderNode.appendChild XMLBenefitNode
                End If
                lEmployeesWritten = lEmployeesWritten + 1
              End If
              
          End Select
        End If
        Call ee.KillBenefitsEx

NEXT_EMPLOYEE:
      Next
NEXT_EMPLOYER:
      Next
    
    If lEmployeesWritten > 0 And (Type_Selected = POT_P11D) Then  'Or Type_Selected = POT_P11DB_ONLY) Then
      'If (Type_Selected = POT_P11D) Then
      node_p11d_record_count.Text = lEmployeesWritten 'IK 31/03/04
      'Else
      '  node_p11d_record_count.Text = "0"
      'End If
      node_declarations.Text = "are due"
    Else
      node_p11d_record_count.Text = "0" 'IK 31/03/04
      node_declarations.Text = "are not due"
    End If
                    
    If Type_Selected = POT_P46 Then
       node_p46_record_count.Text = iReportedCarCount
    End If

    
    If (lEmployeesWritten > 0) Then
      
      m_Prg.Caption = "Completed Paye Online " & lEmployeesWritten & " employees written."
      m_Prg.Indicator = None
ERR_ERRORS:
      If (ErrorCount > 0) And (AtLeastOneError) Then
        bSubmit = F_PayeOnlineWarning.Start(PMT_WARNINGS_AND_ERRORS)
      ElseIf ErrorCount > 0 And (Not AtLeastOneError) Then
        bSubmit = F_PayeOnlineWarning.Start(PMT_WARNINGS_ONLY)
      Else
        bSubmit = True
      End If
      Call ClearCursor
    
      If bSubmit Then
        Set ErrorFilter = Nothing
        
        If MsgBox("PAYE Online document generated successfully." & vbCrLf & "Please confirm if you would like to submit to the HM Revenue & Customs." & vbCrLf & vbCrLf & "Please note this can take some time, the times below indicate the upload time to the Revenue." & vbCrLf & "Once uploaded the system will check the Revenue site for confirmation of completion." & vbCrLf & vbCrLf & UploadTime(Len(XML_doc.xml)) & vbCrLf & vbCrLf & "Please do not interfere.", vbYesNo) = vbYes Then
          Call SetCursor
          sSubmissionID = Submit_EfilerCom(ef, XML_doc, Response_doc)
          sLastFileName = SubmissionTypeToString & "_" & CStr(ey.EmployerRefNo) & CStr(Format$(Now(), "ddmmhh_mmss") & ".xml")
          LastPathAndFileCreated = OutputDirectory & sLastFileName
          Call Save_Files(sSubmissionID, LastPathAndFileCreated, Response_doc)
          If Len(sSubmissionID) > 0 Then
            Call Update_SubIDs(Employers, sSubmissionID)
          End If
        End If
      End If
    Else
      Call MsgBox("Submission is not required as there is no data to transfer to HMRC", vbInformation)
      Call Err.Raise(ERR_PAYE_ONLINE_CANCEL, "CreatePAYEFiles", "Nothing to submit")
    End If
  
CreatePAYEFiles_END:
  Call ClearCursor
  
  Set m_Prg = Nothing
  Call ClearCursor
  Set ErrorFilter = Nothing
  Call CloseErrorLogDB
  Call xReturn("CreatePAYEFiles")
  Exit Sub
CreatePAYEFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CreatePAYEFiles", "Creating submission", "Unable to create Paye Online files.")
  Resume CreatePAYEFiles_END
  Resume
End Sub

Public Function NewEfiler() As EFiler
  Dim ef As EFiler
  
  On Error GoTo Initialise_EfilerCom_ERR
  Call xSet("NewEfiler")

  m_DBPath = p11d32.UserDataDirectoryYearSpecific & S_PAYEONLINE_MDB
  
  If Not FileExists(m_DBPath) Then
    Call Err.Raise(ERR_FILE_EXISTS, "NewEfiler", "The " & S_PAYEONLINE_MDB & " does not exist at the location '" & m_DBPath & "'")
  End If
  
  Set ef = New EFiler
  ef.CheckCanSubmitWithAbatec = Not p11d32.PAYEonline.DisableCheckSubmissionWithAbatec
  ef.AsyncXMLSubmission = True
  ef.PeriodEnd2 = p11d32.Rates.value(TaxYearEnd)
  
  ef.Product = "P11D"
  ef.ProductVersion = GetVersionString(False) ' S_VERSION_NUM 'version number  ef.SubmissionType = atc2efiler.SUBMISSION_TYPE.P11D
  ef.SubmitterRole = SUBMITTER_ROLE.Owner 'flex from employers details
  ef.Sender = SET_AGENT 'IK added, not in library yet.
  ef.Source = "P11D"
  ef.TidySubmissions = False
  ef.TimeOut = -1
  ef.SubmissionType = P11D
  ef.ArchiveLevel = ARCHIVE_LEVEL.Complete
  ef.SchemaVersion = "2004-v1.2" 'PD database
  ef.ProceedSubmission = Efiler_Proceed_Submission 'IF RUNNINGINIDE/demo version then false 'defaults to true 'PD Database
  ef.TestSubmission = Efiler_Test_Submission
  ef.MockSubmission = Efiler_Mock_Submission
  
  'pick them up from empr properties
    
  ef.DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & m_DBPath
  
  Set NewEfiler = ef
  
GoTo Initialise_EfilerCom_END

Initialise_EfilerCom_END:
  Call xReturn("NewEfiler")
  Exit Function
  
Initialise_EfilerCom_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NewEfiler", "NewEfiler", "Error in E-Filer Component in class PAYEOnline.")
  Resume Initialise_EfilerCom_END
  Resume
End Function
Public Sub Initialise_EfilerCom()
  
  
End Sub

Private Function MarkupXMLPresubmission(xml As String) As String
  Dim p0 As Long
  Dim p1 As Long
  Dim snew_xml As String
  Const END_TAG As String = "</P11Db>"
  
  snew_xml = ""
  p1 = 1
  p0 = InStr(p1, xml, END_TAG)
  Do While (p0 > 0)
    snew_xml = snew_xml & Mid$(xml, p1, p0 - p1)
    snew_xml = snew_xml & " "
    snew_xml = snew_xml & END_TAG
    p1 = p0 + Len(END_TAG)
    If (p1 > Len(xml)) Then Exit Do
    p0 = InStr(p1, xml, END_TAG)
  Loop
  
  snew_xml = snew_xml & Mid$(xml, p1)
  MarkupXMLPresubmission = snew_xml
End Function

Private Function ReplaceXMLMetacharactersRef(ByRef sText As String) As String
  If InStrAny(sText, S_INVALID_XML_CHARS, 1, vbBinaryCompare) > 0 Then
    sText = Replace(sText, "&", "&amp;")
    sText = Replace(sText, "'", "&apos;")
    sText = Replace(sText, ">", "&gt;")
    sText = Replace(sText, "<", "&lt;")
    sText = Replace(sText, S_QUOT, "&quot;")
  End If
  ReplaceXMLMetacharactersRef = sText
End Function


Public Function Submit_EfilerCom(ef As EFiler, XML_doc As DOMDocument60, Response_doc As String) As String
  Dim sxml_submit As String
  Dim sSubmissionID As String
  Dim sTempFile As String
  Dim sCompanyName As String
  On Error GoTo Submit_EfilerCom_ERR
  
  Call xSet("Submit_EfilerCom")
  
  sCompanyName = ReplaceXMLMetacharactersRef(m_Entity_Name)
  
  Response_doc = ""
  If FileExists(m_DBPath) Then
    'markup xml before submission
    sTempFile = GetTempFilename(CoreClass.GetTempDirectory, "")
    Call XML_doc.Save(sTempFile)
    sxml_submit = TextFileLoad(sTempFile)
    sxml_submit = MarkupXMLPresubmission(sxml_submit)
    m_Prg.Max = 100
    m_Prg.Min = 0
    m_Prg.value = 0
    ef.companyName = sCompanyName
    'cad temp
    Dim sInvalidSubmission As String
    sInvalidSubmission = GetIniEntry("PayeOnline", "InvalidSubmission", "0", p11d32.IniPathAndFile)
    If (p11d32.LicenceType = LT_STANDARD) And ((Now >= #7/2/2019#) Or sInvalidSubmission = "1") Then
      Call WriteIniEntry("PayeOnline", "InvalidSubmission", "1", p11d32.IniPathAndFile)
      Call Err.Raise(ERR_ERROR, "Submit_EfilerCom", "Gateway error 304B")
    End If
    'cad end temp
    
    Response_doc = ef.Submit2(sSubmissionID, sxml_submit, m_sSubmitter_ID, m_sSubmitter_Pwd, m_Entity_ID, m_Entity_Name, m_sSubmitter_User, Me)
    MsgBox Response_doc
  Else
    Call Err.Raise(ERR_ERROR, "Submit_EfilerCom", "Cannot find PAYEOnline.mdb")
  End If
  
  m_Prg.Caption = "Completed"
  
Submit_EfilerCom_END:
  Submit_EfilerCom = sSubmissionID
  m_Prg.value = m_Prg.Min
  Call xReturn("Submit_EfilerCom")
  Exit Function
Submit_EfilerCom_ERR:
  m_Prg.Caption = "Error"
  Call ErrorMessage(ERR_ERROR, Err, "Submit_EfilerCom", "Submit_EfilerCom", "Error while submitting document in PAYE Online.")
  Resume Submit_EfilerCom_END
  Resume
End Function
Public Sub Update_SubIDs(Employers As ObjectList, ByVal sSubmissionID As String)
  Dim i As Integer
  Dim ey As Employer
  Dim ben As IBenefitClass
  
  
  On Error GoTo Update_SubIDs_ERR
  Call xSet("Update_SubIDs")
    
    For i = 1 To Employers.Count
        Set ey = Employers(i)
        ey.OpenEmployer (False)
        Set ben = ey
        
        If (p11d32.PAYEonline.Type_Selected = POT_P11D Or p11d32.PAYEonline.Type_Selected = POT_P11DB_ONLY) Then
          ben.value(employer_PayeOnlineP11DSubmissionID) = sSubmissionID
          ben.value(employer_PayeOnlineP11DSubmissionDate) = Now
        Else
          Select Case p11d32.PAYEonline.P46Range
            Case P46_RANGE.P46_QUARTER1
              ben.value(employer_PayeOnlinep46Q1SubmissionID) = sSubmissionID
              ben.value(employer_PayeOnlinep46Q1SubmissionID) = Now
            Case P46_RANGE.P46_QUARTER2
              ben.value(employer_PayeOnlinep46Q2SubmissionID) = sSubmissionID
              ben.value(employer_PayeOnlinep46Q2SubmissionID) = Now
            Case P46_RANGE.P46_QUARTER3
              ben.value(employer_PayeOnlinep46Q3SubmissionID) = sSubmissionID
              ben.value(employer_PayeOnlinep46Q3SubmissionID) = Now
            Case P46_RANGE.P46_QUARTER4
              ben.value(employer_PayeOnlinep46Q4SubmissionID) = sSubmissionID
              ben.value(employer_PayeOnlinep46Q4SubmissionID) = Now
          End Select
        End If
        'ben.value(employer_PayeOnlineSubmissionType_db) = SubmissionTypeToString
        Call ben.WriteDB
    Next
    
Update_SubIDs_END:
  Call xReturn("Update_SubIDs")
  Exit Sub
Update_SubIDs_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Update_SubIDs", "Update_SubIDs", "Error while updating Submission IDs PAYEOnline.")
  Resume Update_SubIDs_END
End Sub


Public Function Efiler_Message_Parser(ByVal Msg As String, ByVal Index As Long) As String
  Dim p0 As Long
  Dim myArray() As String
  
  'splitting the message from submissions table
  myArray() = Split(Msg, vbCrLf)
  'parsing the strings in array
  If Index <= UBound(myArray) Then
    p0 = InStr(1, myArray(Index), ":")
    If p0 > 0 Then Efiler_Message_Parser = Trim$(Mid$(myArray(Index), p0 + 1))
  End If
End Function

Private Function GetPayeXML(ID As String) As String

  Dim m_rsLocal As Recordset
  
  On Error GoTo GetPayeXML_ERR
  
  If m_db Is Nothing Then
    Call OpenErrorLogDB
  End If
  
  Set m_rsLocal = m_db.OpenRecordset(sql.Queries(SELECT_PAYEONLINE_XML, ID), dbOpenDynaset)
  If Not m_rsLocal.EOF Then
      Call m_rsLocal.MoveFirst
      GetPayeXML = m_rsLocal.Fields("xml").value
  Else
    GetPayeXML = ""
  End If

GetPayeXML_END:
  Call xSet("GetPayeXML")
  m_rsLocal.Close
  Set m_rsLocal = Nothing
  Exit Function
GetPayeXML_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetPayeXML", "Check PayeOnline Status", "Error while getting XML for Submissions.")
  Resume GetPayeXML_END
  Resume
End Function

Public Function CheckPAYEEmployerDetails(benEmployer As IBenefitClass) As String
  Dim ey As Employer
  
  'Apply validation checks to PAYEEmployer fields
  On Error GoTo CheckPAYEEmployerDetails_ERR
  
  'Check required fields
  Set ey = benEmployer
  If Not ey.PAYEOnlineValid(True) Then
    AtLeastOneError = True
  End If
  
  'Check optional fields
  If Not Len(benEmployer.value(employer_PayeOnlineEmail_db)) > 0 Then
    Call Err.Raise(ERR_PAYEFIELD_OPTIONAL_NOT_FILLED, "PAYEOnlineValid", "The PAYE online email for employer " & benEmployer.value(employer_Name_db) & " is missing.")
  End If
  Set ey = Nothing
  
CheckPAYEEmployerDetails_END:
  Call xSet("CheckPAYEEmployerDetails")
  Exit Function

CheckPAYEEmployerDetails_ERR:
  Call ErrorMessage(Err.Number, Err, "CheckPayeEmployerDetails", FilterMessageTitle(benEmployer.value(employer_Name_db)), Err.Description)
  Resume Next 'Errors are logged in this class
  Resume CheckPAYEEmployerDetails_END
  Resume
End Function
Private Function EmployeeHasp46Cars(ByVal ee As Employee)
  Dim OL As ObjectList
  
  Call ee.LoadBenefits(TBL_COMPANY_CARS, False, False)
  Call ee.GetP46Cars(OL, p11d32.PAYEonline.P46DateFrom, p11d32.PAYEonline.P46DateTo, True)
  If OL.Count Then
    EmployeeHasp46Cars = True
  End If

End Function
Private Function RefreshEmployerHasP46Cars(ey As Employer) As String
  'Refreshes EmployerHasP46Cars
  Dim ee As Employee
  Dim benEmployer As IBenefitClass
  Dim i As Long
  On Error GoTo RefreshEmployerHasP46Cars_ERR
  
  Set benEmployer = ey
  If Not benEmployer Is Nothing Then
    benEmployer.value(employer_hasP46Cars) = False
    For i = 1 To ey.employees.Count
      Set ee = ey.employees(i)
      If EmployeeHasp46Cars(ee) Then
        benEmployer.value(employer_hasP46Cars) = True
        Exit For
      End If
    Next
  End If
  
RefreshEmployerHasP46Cars_END:
  Call xSet("RefreshEmployerHasP46Cars")
  Set benEmployer = Nothing
  Exit Function
RefreshEmployerHasP46Cars_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RefreshEmployerHasP46Cars", benEmployer.value(employer_Name_db), Err.Description)
  Resume Next 'Errors are logged in this class
  Resume RefreshEmployerHasP46Cars_END
  Resume
End Function

Private Sub Save_Files(ByVal sSubmissionID As String, LastPathAndFileCreated As String, Response_doc As String)
  Dim XML_Submitted As DOMDocument60
  Dim txtLastPathAndFileCreated As String
  
  On Error GoTo save_files_ERR
  
  Set XML_Submitted = New DOMDocument60
  
  XML_Submitted.loadXML (GetPayeXML(sSubmissionID))
  If Len(XML_Submitted.xml) Then
    Call XML_Submitted.Save(LastPathAndFileCreated)
    
  End If
  
  If Len(Response_doc) Then
    txtLastPathAndFileCreated = Replace(LastPathAndFileCreated, ".xml", ".txt")
    Call SaveTextFile(txtLastPathAndFileCreated, Response_doc)
  End If
save_files_END:
  Call xSet("save_files")
  Set XML_Submitted = Nothing
  Exit Sub
save_files_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Save_Files", "Save_Files", "Error saving the xml files.")
  Resume save_files_END
  Resume
End Sub


Public Sub Initialise()

  On Error GoTo Initialise_PAYEOnline_ERR
  Call xSet("Initialise_PAYEOnline")
  
  
  
Initialise_PAYEOnline_END:
  Call xReturn("Initialise_PAYEOnline")
  Exit Sub
Initialise_PAYEOnline_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Initialise_PAYEOnline", "Initialise PAYEOnline", "Error initialising the report print class.")
  Resume Initialise_PAYEOnline_END
End Sub

Private Sub ValidateP46Car(car As CompanyCar)
  Dim benEmployer As IBenefitClass, ee As Employee

On Error GoTo ValidateP46Car_ERR

    Dim temp As String, sTxt As String, att As String
    Dim sMakeAndModel As String
    Dim err_MakeModel As String 'used in error message
    Dim ben As IBenefitClass

    Set ben = car
    err_MakeModel = ben.Name
    Call CheckCarListPrice(ben)
                
                    
    Select Case Type_Selected
     Case POT_P46
        If ben.value(car_P46FirstProvidedWithCar) = True Then
        ElseIf ben.value(car_P46CarProvidedReplaced) = True Then
            'checking car_provided_replaced fields
            'checking for multiple cars
            
            If Not PAYEFieldValid(ben.value(car_CarReplacedEngineSize_db), PAYE_P46_ENGINESIZE) Then
                Call Err.Raise(ERR_CAR_ENGINE_SIZE, "ValidateP46Car", "Engine Size '" & ben.value(car_CarReplacedEngineSize_db) & "' is invalid.")
                AtLeastOneError = True
            End If
            
        
        ElseIf ben.value(car_P46SecondCar) = True Then
            
        ElseIf ben.value(car_P46WithdrawnWithoutReplacement) = True Then
            'engine size
            
            If Not PAYEFieldValid(ben.value(car_enginesize_db), PAYE_P46_ENGINESIZE) Then
                Call Err.Raise(ERR_CAR_ENGINE_SIZE, "ValidateP46Car", "Engine Size '" & ben.value(car_enginesize_db) & "' invalid.")
                AtLeastOneError = True
            End If
            'withdrawal date
            If ben.value(Car_AvailableTo_db) > Now Then
               'IK NOTE doc section 15 does not demand a date check ... therfore warning
                Call Err.Raise(ERR_PAYE_DATE_OUTOFSYNC, "ValidateP46Car", "Car withdrawn date '" & ben.value(Car_AvailableTo_db) & "' is after todays date '" & DateValReadToScreen(VBA.Now) & "'.")
                AtLeastOneError = True
            End If
            'make and model
            'don't bother as we trim anyway

            'dont need car details etc in this case therefore exit
            GoTo Validate_end
        Else
        
        End If
              
        'CAR DETAILS
        'engine size
        temp = ben.value(car_enginesize_db)
        If Not PAYEFieldValid(temp, PAYE_P46_ENGINESIZE) Then
            Call Err.Raise(ERR_CAR_ENGINE_SIZE, "ValidateP46Car", "Engine Size '" & temp & "' invalid.")
            AtLeastOneError = True
        End If
        'make and model in P46 format
        sMakeAndModel = TrimP46MakeAndModel(ben.value(car_Make_db) + " " + ben.value(car_Model_db))
        'registration date
        If ben.value(car_Registrationdate_db) > VBA.Now Then 'WE ARE ASSUMING THEIR CLOCK IS RIGHT ??? cad
          'IK NOTE TODO check message syntax
         AtLeastOneError = True
         Call Err.Raise(ERR_PAYE_DATE_OUTOFSYNC, "ValidateP46Car", "Registration date '" & ben.value(car_Registrationdate_db) & "' must be before submission date '" & VBA.Now & "'.")
        End If
        
        'CO2
        If ben.value(car_p46CarbonDioxide_db) > 0 Then
            'no warning
        ElseIf ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
            'arning saying no approved emissions but xml not entered because registered before 1998
            Call Err.Raise(ERR_CAR_CO2, "ValidateP46Car", "No CO2 emissions figure. Car registered >= " & DateValReadToScreen(p11d32.Rates.value(carNOCO2CUTOFFDATE)) & ".")
        End If
                
        
        'MONETARY DETAILS
        'list price
        If Not PAYEFieldValid(ben.value(car_ListPrice_db), PAYE_AMOUNT) Then
            Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "List Price '" & ben.value(car_ListPrice_db) & "' is invalid. List Prices must be less that 100,000")
            AtLeastOneError = True
        End If
        'new accessories
        If Not PAYEFieldValid(ben.value(car_AccessoriesNew_db), PAYE_AMOUNT) Then
            Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "New Accessories price '" & ben.value(car_AccessoriesNew_db) & "' is invalid.")
            AtLeastOneError = True
        End If
        'avaliable from date
        If ben.value(Car_AvailableFrom_db) < p11d32.Rates.value(TaxYearStart) Then
            AtLeastOneError = True
            Call Err.Raise(ERR_PAYE_DATE_OUTOFSYNC, "ValidateP46Car", "Available From date '" & ben.value(Car_AvailableFrom_db) & "' must be after start of current tax year '" & p11d32.Rates.value(TaxYearEnd) & "'.")
        ElseIf ben.value(Car_AvailableFrom_db) > VBA.Now Then
            AtLeastOneError = True
            Call Err.Raise(ERR_PAYE_DATE_OUTOFSYNC, "ValidateP46Car", "Available From date '" & ben.value(Car_AvailableFrom_db) & "' must be before submission date '" & VBA.Now & "'.")
        ElseIf ben.value(Car_AvailableFrom_db) < ben.value(car_Registrationdate_db) Then
            AtLeastOneError = True
            Call Err.Raise(ERR_PAYE_DATE_OUTOFSYNC, "ValidateP46Car", "Available From date '" & ben.value(Car_AvailableFrom_db) & "' must be after its registration date'" & ben.value(car_Registrationdate_db) & "'.")
        End If
        'capital contribution restricted
        If Not PAYEFieldValid(ben.value(car_CapitalContributionRestricted), PAYE_AMOUNT) Then
            Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Capital contribution Amount '" & ben.value(car_CapitalContributionRestricted) & "' is invalid.")
            AtLeastOneError = True
        End If
        'amount made good ... the whole amount -> (amount entered * payment frequency)
        If (ben.value(car_ActualAmountMadeGood) > 0) And Not (PAYEFieldValid(ben.value(car_ActualAmountMadeGood), PAYE_AMOUNT)) Then
                Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Actual amount made good '" & ben.value(car_ActualAmountMadeGood) & "' is invalid.")
                AtLeastOneError = True
        End If
        
        
    Case POT_P11D
        'P11D Code here
        Set ben = car
        'Check engine size or CO2
        If ben.value(car_enginesize_db) > 9999 Or ben.value(car_enginesize_db) < 0 Then
          Call Err.Raise(ERR_CAR_ENGINE_SIZE, "ValidateP46Car", "Engine Size '" & ben.value(car_enginesize_db) & "' invalid.")
          AtLeastOneError = True
        End If

        If ((ben.value(car_enginesize_db) = 0 And ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE)) Or _
            ((ben.value(car_enginesize_db) = 0) And (ben.value(car_p46NoApprovedCO2Figure_db))) Or _
            ((Not ben.value(car_p46NoApprovedCO2Figure_db)) And (ben.value(car_p46CarbonDioxide_db) = 0)) Or _
            (ben.value(car_FuelBenefit) > 0 And ben.value(car_enginesize_db) = 0 And ben.value(car_p46CarbonDioxide_db) = 0)) And (ben.value(car_p46FuelType_db) <> CCFT_ELECTRIC) Then
          Call Err.Raise(ERR_NEEDCO2ORENGINESIZE, , "Car '" & ben.value(car_Registration_db) & "' needs an engine size or C02 figure.")
          AtLeastOneError = True
        End If
        
        'Check electric cars
        If ben.value(car_p46FuelType_db) = CCFT_ELECTRIC And ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
          Call Err.Raise(ERR_CAR_ELEC_PRE_98, , "Car " & ben.value(car_Registration_db) & "only fuel type E permitted if car is reg. on or after 01/01/98")
          AtLeastOneError = True
        End If
            
        'List price check
        If ben.value(car_ListPrice_db) < 1000 Or ben.value(car_ListPrice_db) > D_XMLBENEFITMAX Then
          Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "List Price '" & ben.value(car_ListPrice_db) & "' is invalid.")
          AtLeastOneError = True
        End If
        
        'Check accessories amount
        If ben.value(car_Accessories) > D_XMLBENEFITMAX Then
          Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Accessories amount '" & ben.value(car_ListPrice_db) & "' is too great, maximum = " & D_XMLBENEFITMAX)
          AtLeastOneError = True
        End If
        
        'Check capital contribution - only report amount up to £5000.
        If ben.value(car_capitalcontribution_db) > 5000 Then
          Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Maximum Capital Contribution reported will be 5000.")
        End If
        
        'check private use payment
        If ben.value(car_MadeGood_db) > D_XMLBENEFITMAX Then
          Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Car made good amount '" & ben.value(car_MadeGood_db) & "' is too great, maximum = " & D_XMLBENEFITMAX)
          AtLeastOneError = True
        End If
        
        'check car benefit amount
        If ben.value(car_Benefit_Reportable) > D_XMLBENEFITMAX Then
          Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Car benefit amount '" & ben.value(car_Benefit_Reportable) & "' is too great, maximum = " & D_XMLBENEFITMAX)
          AtLeastOneError = True
        End If
        
        'check car fuel benefit amount
        If ben.value(car_FuelBenefit_Reportable) > D_XMLBENEFITMAX Then
          Call Err.Raise(ERR_CAR_AMOUNT, "ValidateP46Car", "Car fuel benefit amount '" & ben.value(car_Benefit_Reportable) & "' is too great, maximum = " & D_XMLBENEFITMAX)
          AtLeastOneError = True
        End If
          
        
        
    End Select
    
'common errors
Validate_end:

  If ben.value(car_p46FuelType_db) = CCFT_ELECTRIC And ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
    Call Err.Raise(ERR_CAR_ELEC_PRE_98, , "Car " & ben.value(car_Registration_db) & "only fuel type E permitted if car is reg. on or after 01/01/98")
    AtLeastOneError = True
  End If
  
  ''removed as no longer checked for by PayeOnline
  'If ben.value(car_p46FuelType_db) = CCFT_DIESEL And ben.value(car_Registrationdate_db) >= D_FUEL_TYPE_L_NO_LONGER_HAS_DISCOUNT Then
  '  Call Err.Raise(ERR_DIESEL_REGISTERED_AFTER_1_1_2006, , "Has fuel type D, change to L as registered after 31/12/2005")
  '  AtLeastOneError = True
  'End If
  
    
ValidateP46Car_END:
  Call xReturn("ValidateP46Car")
  Exit Sub
ValidateP46Car_ERR:
  If (Err.Number = ERR_CAR_LIST_PRICE) Then AtLeastOneError = True 'this is how it should be done CAD 05/05/2004
  
    
  Set ee = ben.Parent
  Set benEmployer = ee.Parent
  Call ErrorMessage(ERR_ERROR, Err, "ValidateP46Car", FilterMessageTitle(benEmployer.value(employer_Name_db), ee.PersonnelNumber, , err_MakeModel, "Company Car"), Err.Description)
  Set benEmployer = Nothing
  Set ee = Nothing
  Resume Next 'Errors are logged via IErrorFilter
  Resume

End Sub
Private Function ProceedWithEmployee(ee As Employee, Optional OL As ObjectList) As Boolean
  Dim benEE As IBenefitClass
  
  ProceedWithEmployee = False
  
  Call ee.LoadBenefits(TBL_ALLBENEFITS, False)
  Set benEE = ee
  Call benEE.Calculate
    
  ProceedWithEmployee = True
  
  If ee.benefits.Count = 0 Then
    ProceedWithEmployee = False
  ElseIf p11d32.PAYEonline.Type_Selected = POT_P11D Or p11d32.PAYEonline.Type_Selected = POT_P11DB_ONLY Then
    Set benEE = ee
    If (IsNumeric(benEE.value(ITEM_VALUE))) Then
      If benEE.value(ITEM_VALUE) = 0 Then
        ProceedWithEmployee = False
      End If
    End If
  Else
    If (p11d32.PAYEonline.Type_Selected = POT_P46) And (Not (ee.GetP46Cars(OL, p11d32.PAYEonline.P46DateFrom, p11d32.PAYEonline.P46DateTo, True))) Then
      ProceedWithEmployee = False
    Else
      ProceedWithEmployee = True
    End If
  End If
  
End Function
Private Function EmployeeIsReportable(ByVal ee As Employee)
  Dim OL As ObjectList
  If p11d32.PAYEonline.Type_Selected = POT_P46 Then
    EmployeeIsReportable = EmployeeHasp46Cars(ee)
  Else
    EmployeeIsReportable = ee.IterateBenefits(AnyReportable)
  End If
  
End Function
'RETURNS FALSE IF A CHECK BEN HAS FAILED ... THIS IS VERY BAD CODE - WE MUST CORRET IT ALL cad 2006
Private Function ValidateEmployees(ey As Employer, lEmployeesWithInvalidNI As Long, lEmployeesWritten As Long) As Boolean
  Dim i As Long, j As Long
  Dim ee As Employee
  Dim benEmployee As IBenefitClass
  Dim ben As IBenefitClass
  Dim BenArr() As BEN_CLASS
  Dim Employer As IBenefitClass
  Dim benefitCar As CompanyCar
  Dim sHMITSectionString As String
  Dim sBenefitFormCaption As String
  Set Employer = ey
  Dim iNoOfCars As Integer
  Dim iNoOfFuelWithdrawn As Integer
  Dim iAnyCars As Integer
  Dim sTxt As String, lNum As Long
  Dim mycar As CompanyCar
  Dim OL As ObjectList
  
    
  On Error GoTo ValidateEmployees_err
  
  ReDim BenArr(1 To 1)
  iAnyCars = 0
  m_Prg.Min = 0
  m_Prg.Max = ey.employees.Count
  m_Prg.Indicator = ValueOfMax
  ValidateEmployees = True
  For i = 1 To ey.employees.Count
    Set ee = ey.employees(i)
    Set benEmployee = ee
    iNoOfCars = 0
    iNoOfFuelWithdrawn = 0
    sHMITSectionString = ""
    sBenefitFormCaption = ""
    If ee Is Nothing Then
      Call m_Prg.StepCaption("Validating employee:" & "nothing")
      GoTo NEXT_EMPLOYEE
    End If
    
    Call m_Prg.StepCaption("Validating employee:" & ee.FullName)
    
    If Not ProceedWithEmployee(ee, OL) Then
      GoTo NEXT_EMPLOYEE
    End If
    
    'increment counter
    lEmployeesWritten = lEmployeesWritten + 1
    
    'Check Personnel Number
    Call CheckEmployeePersonnelNumberForSpaces(ee)
    Call CheckEmployeePersonnelNumberValidForEDI(ee)
    'check NI Number
    If EmployeeIsReportable(ee) Then
      If (Not NIValid(ee)) Then
        lEmployeesWithInvalidNI = lEmployeesWithInvalidNI + 1
        If StrComp(benEmployee.value(ee_Gender_db), S_GENDER_NA, vbTextCompare) = 0 Then
          If p11d32.PAYEonline.Type_Selected = POT_P46 Then
            Call Err.Raise(ERR_NI_INVALID_BUT_ADDED, "ValidateEmployees", "NI '" & benEmployee.value(ee_NINumber_db) & "' invalid, and gender has not been supplied.")
          Else
            Call Err.Raise(ERR_NI_INVALID_BUT_ADDED, "ValidateEmployees", "NI '" & benEmployee.value(ee_NINumber_db) & "' invalid, and gender has not been supplied.")
          End If
          AtLeastOneError = True
        End If
      End If
    End If
    
    'forename
    sTxt = Trim$(benEmployee.value(ee_Firstname_db))
    lNum = InStr(1, sTxt, " ")
    If lNum > 0 Then
      sTxt = Mid(sTxt, 1, lNum - 1)
    End If
    
    If Not PAYEFieldValid(sTxt, PAYE_FORENAME) Then
      If (Len(sTxt) = 0) Then
        Call Err.Raise(ERR_PAYE_INVALID_FIELD, "ValidateEmployees", "Forename is zero length, will replace with " & S_UNKNOWN)
      Else
        Call Err.Raise(ERR_PAYE_INVALID_FIELD, "ValidateEmployees", "Forename '" & benEmployee.value(ee_Firstname_db) & "' is invalid -  - max length is 35, check for commas and full stops.")
      End If
    ElseIf lNum Then
      Call Err.Raise(ERR_PAYE_CHANGED_FIELD, "ValidateEmployees", "Forename '" & benEmployee.value(ee_Firstname_db) & "' was changed to '" & sTxt & " ' as spaces not allowed")
    End If
    
    'surname
    If Not PAYEFieldValid(benEmployee.value(ee_Surname_db), PAYE_SURNAME) Then
      Call Err.Raise(ERR_PAYE_INVALID_FIELD, "ValidateEmployees", "Surname '" & benEmployee.value(ee_Surname_db) & "' is invalid - max length is 35, check for commas and full stops.")
    End If
    
    'Apply PAYEOnline Type specific checks
    Select Case p11d32.PAYEonline.Type_Selected
      Case POT_P11D, POT_P11DB_ONLY
      'p11d specific checks here
        For j = 1 To ee.benefits.Count
          Set ben = ee.benefits(j)
          If Not ben Is Nothing Then
            If ben.BenefitClass <> BC_NONSHAREDVANS_G And ben.BenefitClass <> BC_LOANS_H Then
              If CheckBen(ey, ee, ben) Then
                BenArr(1) = ben.BenefitClass
                sHMITSectionString = p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION_STRING)
                sBenefitFormCaption = p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_FORM_CAPTION)
                
                Select Case ben.BenefitClass
                  Case BC_COMPANY_CARS_F
                    Set benefitCar = ben
                    iAnyCars = iAnyCars + 1
                    Call ValidateP46Car(benefitCar)
                    iNoOfCars = iNoOfCars + 1
                    If iNoOfCars > 50 Then
                      Call Err.Raise(ERR_BEN_NOT_REPORTABLE, "ValidateEmployeeBenefits", "Employee '" & ee.FullName & "' has too many cars.")
                      iNoOfCars = 0
                    End If
                End Select
                
                If ben.value(ITEM_VALUE) > D_XMLBENEFITMAX Or ben.value(ITEM_BENEFIT) > D_XMLBENEFITMAX Or ben.value(ITEM_MADEGOOD) > D_XMLBENEFITMAX Then
                  Call Err.Raise(ERR_XMLNUMBERTOOBIG, "ValidateEmployeeBenefits", "Benefit, made good or value too large to be reported (greater than " & D_XMLBENEFITMAX & ")")
                End If
              Else
                ValidateEmployees = False
              End If
            End If
          End If
        Next
            
      Case POT_P46
        'validate cars data
        For j = 1 To OL.Count
          Set mycar = OL.Item(j)
          
          Call ValidateP46Car(mycar)
        Next
      Case Else
        Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateEmployees", "PAYEOnline type not supported")
    End Select
    
NEXT_EMPLOYEE:
  Next i
  ValidateEmployees = True
  
ValidateEmployees_end:
  m_Prg.value = 0
  m_Prg.Caption = ""
  m_Prg.Indicator = None
  Set ee = Nothing
  Call xReturn("ValidateEmployees")
  Exit Function

ValidateEmployees_err:
  If ben Is Nothing Then
    If Err.Number = ERR_PAYE_INVALID_FIELD Or Err.Number = ERR_XMLNUMBERTOOBIG Then
      AtLeastOneError = True
    End If
    Call ErrorMessagePop(Err.Number, Err, "ValidateEmployeeBenefits", FilterMessageTitle(Employer.Name, ee.PersonnelNumber), "")
  ElseIf Err.Number = ERR_ROTARY Then
    Call ErrorMessagePop(Err.Number, Err, "ValidateEmployeeBenefits", FilterMessageTitle(Employer.Name, ee.PersonnelNumber), "")
  Else
    Call ErrorMessagePop(Err.Number, Err, "ValidateEmployeeBenefits", FilterMessageTitle(Employer.Name, ee.PersonnelNumber, sHMITSectionString, , sBenefitFormCaption), "")
  End If

  Resume Next 'Errors are logged via IErrorFilter
  Resume

End Function
Private Property Get SubmissionTypeToString()
  If Type_Selected = POT_P11D Or Type_Selected = POT_P11DB_ONLY Then
    SubmissionTypeToString = "P11D"
  ElseIf Type_Selected = POT_P46 Then
    If p11d32.PAYEonline.P46Range = P46_QUARTER1 Then
      SubmissionTypeToString = "P46Q1"
    ElseIf p11d32.PAYEonline.P46Range = P46_QUARTER2 Then
      SubmissionTypeToString = "P46Q2"
    ElseIf p11d32.PAYEonline.P46Range = P46_QUARTER3 Then
      SubmissionTypeToString = "P46Q3"
    ElseIf p11d32.PAYEonline.P46Range = P46_QUARTER4 Then
      SubmissionTypeToString = "P46Q4"
    End If
  Else
    Call Err.Raise(380, "SubmissionTypeToString", "Invalid submission type")
  End If
End Property


Private Function ValidateEmployers(ef As EFiler, Employers As ObjectList) As Boolean
  'Applies PAYEOnline validation checks to current batch of employers
  On Error GoTo ValidateEmployers_err
  
  Dim i As Integer
  Dim benEmployer As IBenefitClass
  Dim ben2 As IBenefitClass
  Dim ey As Employer
  Dim lEmployeesWithInvalidNI As Long
  Dim lEmployeesWritten As Long
  Dim iAdjustmentsCount As Integer
  Dim sAddDesc As String, sDeductDesc As String
  Dim sEmployerName As String
  Dim sEmployersAlreadySubmitted As String
  Dim employer_paye_online_submitted_field As Long
  Dim employer_paye_online_submitted_date_field As Long
  Dim sEmailAddress As String
  
  
  Call xSet("ValidateEmployers")
  
  If (p11d32.PAYEonline.Type_Selected = POT_P11D Or p11d32.PAYEonline.Type_Selected = POT_P11DB_ONLY) Then
    employer_paye_online_submitted_field = employer_PayeOnlineP11DSubmissionID
    employer_paye_online_submitted_date_field = employer_PayeOnlineP11DSubmissionDate
  Else
    If (p11d32.PAYEonline.P46Range = P46_QUARTER1) Then
      employer_paye_online_submitted_field = employer_PayeOnlinep46Q1SubmissionID
      employer_paye_online_submitted_date_field = employer_PayeOnlinep46Q1SubmissionID
    ElseIf (p11d32.PAYEonline.P46Range = P46_QUARTER2) Then
      employer_paye_online_submitted_field = employer_PayeOnlinep46Q2SubmissionID
      employer_paye_online_submitted_date_field = employer_PayeOnlinep46Q2SubmissionID
    ElseIf (p11d32.PAYEonline.P46Range = P46_QUARTER3) Then
      employer_paye_online_submitted_field = employer_PayeOnlinep46Q3SubmissionID
      employer_paye_online_submitted_date_field = employer_PayeOnlinep46Q3SubmissionID
    Else
      employer_paye_online_submitted_field = employer_PayeOnlinep46Q4SubmissionID
      employer_paye_online_submitted_date_field = employer_PayeOnlinep46Q4SubmissionID
    End If
  End If

  ValidateEmployers = True
  For i = 1 To Employers.Count
    
    'Load employer
    Set ey = Employers(i)
    Call p11d32.LoadEmployer(ey, False)
    
    If Not ey.DataChecker(-2, Nothing) Then Call Err.Raise(ERR_PAYE_ONLINE_CANCEL, "Validate Employers", "Submission cancelled")
     
    
    Set benEmployer = ey
    If benEmployer.value(employer_EmployeesCount) = 0 Then GoTo NEXT_EMPLOYER
    ' Set ben2 = p11d32.Employers(i)
    
    
    If (Len(benEmployer.value(employer_paye_online_submitted_field)) > 0) Then
      sEmployersAlreadySubmitted = sEmployersAlreadySubmitted & benEmployer.Name & ", submitted on " & DateValReadToScreen(benEmployer.value(employer_paye_online_submitted_date_field)) & vbCrLf
    End If
    
    'Check Employer PAYE details (required and optional fields)
    Call CheckPAYEEmployerDetails(benEmployer)
    
    'Check Duplicate NI Numbers
    Call CheckDuplicateNINumbers(ey)
    
    'Validate Employees
    If Not ValidateEmployees(ey, lEmployeesWithInvalidNI, lEmployeesWritten) Then
      ValidateEmployers = False
    End If
    
    'Check NIRatio - removed 2008
    'Call CheckNIRatio(lEmployeesWithInvalidNI, lEmployeesWritten)
    
    'For first employer record details
    If i = 1 Then
      m_sSubmitter_Ref = benEmployer.value(employer_SubmitterRef_db)
      m_sSubmitter_ID = benEmployer.value(employer_PayeOnlineID_db)
      m_sSubmitter_Pwd = benEmployer.value(employer_PayeOnlinePwd_db)
      m_sSubmitter_User = benEmployer.value(employer_PayeOnlineUser_db)
      m_sTaxOfficeNumber = Left$(benEmployer.value(employer_Payeref_db), 3) ' benEmployer.value(employer_IRTaxOffice_db)
      m_sTaxOfficeReference = ey.EmployerRefNo  ' benEmployer.value(employer_IRYourReference)
      sEmployerName = benEmployer.Name
      
      'Email is an Optional field
      sEmailAddress = Trim$(benEmployer.value(employer_PayeOnlineEmail_db))
      If Len(sEmailAddress) > 0 Then
        ef.EmailAddress = sEmailAddress
      End If
      
      'update EfilerCom with this info
      
      Call ef.AddKey(m_sTaxOfficeNumber, KEY_TYPE.TaxOfficeNumber)
      Call ef.AddKey(m_sTaxOfficeReference, KEY_TYPE.TaxOfficeReference)
     
    Else
    'If multiple employers have been selected then they need to have the same fields
      'Check SubmitterID identical
      If StrComp(benEmployer.value(employer_PayeOnlineID_db), m_sSubmitter_ID, vbTextCompare) <> 0 Then
        Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, "ValidateEmployers", "The Online Services IDs for the selected employers differ.")
        AtLeastOneError = True
      End If
      '(Check SubmitterPwd identical
      If StrComp(benEmployer.value(employer_PayeOnlinePwd_db), m_sSubmitter_Pwd, vbTextCompare) <> 0 Then
        Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, "ValidateEmployers", "The Online Services passwords for the selected employers differ.")
        AtLeastOneError = True
      End If
      'Check SubmitterUsername identical
      If StrComp(benEmployer.value(employer_PayeOnlineUser_db), m_sSubmitter_User, vbTextCompare) <> 0 Then
        Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, "ValidateEmployers", "The Online Services usernames for the selected employers differ.")
        AtLeastOneError = True
      End If
      'Tax office number identical
      If StrComp(Left(benEmployer.value(employer_Payeref_db), 3), m_sTaxOfficeNumber, vbTextCompare) <> 0 Then
        Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, "ValidateEmployers", "The tax office numbers for the selected employers differ.")
        AtLeastOneError = True
      End If
      'Tax office reference identical
      If StrComp(ey.EmployerRefNo, m_sTaxOfficeReference, vbTextCompare) <> 0 Then
        Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, "ValidateEmployers", "The tax office references for the selected employers differ.")
        AtLeastOneError = True
      End If
      'Check SubmitterEmail (optional field) identical
      sEmailAddress = Trim$(benEmployer.value(employer_PayeOnlineEmail_db))
      If Len(sEmailAddress) > 0 Then
        If Len(ef.EmailAddress) > 0 Then
          'Check for differences
          If StrComp(sEmailAddress, ef.EmailAddress, vbTextCompare) <> 0 Then
            Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, "ValidateEmployers", "The Online Services email addesses for the selected employers differ.")
          End If
        Else
          ef.EmailAddress = sEmailAddress
        End If
      End If
    End If
      
    'Apply PAYEOnline Type specific checks
    Select Case p11d32.PAYEonline.Type_Selected
      Case POT_P11D, POT_P11DB_ONLY
        'Apply Magnetic Media validation checks
        'P11Db checks
        Call benEmployer.Calculate 'employees are already calced in Validate Employees
        If (benEmployer.value(employer_NIC_AdjustmentAdd) > 0 Or benEmployer.value(employer_NIC_AdjustmentDeduct) > 0) Then
          If iAdjustmentsCount = 0 Then
            sDeductDesc = benEmployer.value(employer_deductClass1ADescription_db)
            sAddDesc = benEmployer.value(employer_AddClass1ADescription_db)
          End If
          iAdjustmentsCount = iAdjustmentsCount + 1
        End If
        If benEmployer.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A) < 0 Then
          Call Err.Raise(ERR_PAYE_INVALID_FIELD, , "P11D(b) final amounts cannot be negative.")
          AtLeastOneError = True
        End If
        If Len(benEmployer.value(employer_deductClass1ADescription_db)) > 35 Then
          Call Err.Raise(ERR_PAYE_INVALID_FIELD, , "P11D(b) deductions description field cannot be longer than 35 characters in length.")
          AtLeastOneError = True
        End If
        If Len(benEmployer.value(employer_AddClass1ADescription_db)) > 35 Then
          Call Err.Raise(ERR_PAYE_INVALID_FIELD, , "P11D(b) additions description field cannot be longer than 35 characters in length.")
          AtLeastOneError = True
        End If
      Case POT_P46
        Call RefreshEmployerHasP46Cars(ey)
      Case Else
        Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateEmployers", "PAYEOnline type not supported")
    End Select

NEXT_EMPLOYER:
  Next i
  If Employers.Count > 1 Then
    Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, , "For online filing only the Employer name " & sEmployerName & " will be used.")
    If iAdjustmentsCount > 1 Then
      Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, , "For online filing only the P11D(b) additions description " & sAddDesc & " will be used.")
      Call Err.Raise(ERR_PAYE_DIFFERENT_EMPLOYER_FIELDS, , "For online filing only the P11D(b) deductions description " & sDeductDesc & " will be used.")
    End If
  End If
  
  If (Len(sEmployersAlreadySubmitted) > 0) Then
     sEmployersAlreadySubmitted = "The following employers have already been submitted for '" & SubmissionTypeToString & "':" & vbCrLf & vbCrLf & sEmployersAlreadySubmitted
     sEmployersAlreadySubmitted = sEmployersAlreadySubmitted & vbCrLf & vbCrLf
     sEmployersAlreadySubmitted = sEmployersAlreadySubmitted & "do you wish to continue?"
    If (MsgBox(sEmployersAlreadySubmitted, vbYesNo, "Continue with submission?") = vbNo) Then
      Call Err.Raise(ERR_PAYE_ONLINE_CANCEL, "ValidateEmployers", "Submssion cancelled by user")
    End If
  End If
  
ValidateEmployers_end:
  Call xReturn("ValidateEmployers")
  Exit Function

ValidateEmployers_err:
  'this is wrong we must detect each error as a filter error else err.raise and exit function
  'best way is to use hidden error enums and filter message errors are in a certain range
  
  If (Err.Number = ERR_PAYE_ONLINE_CANCEL) Then
    Call Err.Raise(Err.Number, "ValidateEmployers", Err.Description)
  End If
  
  Call ErrorMessage(Err.Number, Err, "ValidateEmployers", FilterMessageTitle(benEmployer.Name), "")
  Resume Next 'Errors are logged via IErrorFilter
  Resume ValidateEmployers_end
  Resume
End Function
Private Function Analyse_Benefits(XML_doc As DOMDocument60, er As IBenefitClass, ee As Employee, benefits_node As IXMLDOMNode) As Boolean

    Dim BenArr() As BEN_CLASS
    Dim benclass As BEN_CLASS
    Dim HMITArr() As HMIT_SECTIONS
    Dim HMITSection As HMIT_SECTIONS
    Dim i As Integer, Count As Integer
    Dim ben As IBenefitClass
    
    On Error GoTo Analyse_Benefits_ERR
    Call xSet("Analyse_Benefits")
    
    'Adding each benefit class to the arraay
    Count = 0
    For i = 1 To ee.benefits.Count
        Set ben = ee.benefits(i)
        benclass = ben.BenefitClass
        HMITSection = p11d32.Rates.BenClassTo(benclass, BCT_HMIT_SECTION)
        'checking if this Section already exists in the array
        If (Count = 0) Then GoTo AddHMITSection
        If Not HMITSectionInArray(HMITArr, HMITSection) Then
AddHMITSection:
            Count = Count + 1
            ReDim Preserve BenArr(1 To Count)
            ReDim Preserve HMITArr(1 To Count)
            BenArr(Count) = benclass
            HMITArr(Count) = HMITSection
        End If
    Next
    'Sorting them in ascending order as required by Gateway requirements
    Call SortBenArray(BenArr)
        
    'calling functions for each HMIT Section
    For i = LBound(BenArr) To UBound(BenArr)
        benclass = BenArr(i)
        HMITSection = p11d32.Rates.BenClassTo(benclass, BCT_HMIT_SECTION)
        Select Case HMITSection
            Case HMIT_A
                Analyse_Benefits = Analyse_Benefits Or SectionAandL(XML_doc, er, ee, benefits_node, benclass)
            Case HMIT_B
                Analyse_Benefits = Analyse_Benefits Or SectionB(XML_doc, er, ee, benefits_node)
            Case HMIT_C
                Analyse_Benefits = Analyse_Benefits Or SectionC(XML_doc, er, ee, benefits_node)
            Case HMIT_D
                Analyse_Benefits = Analyse_Benefits Or SectionDandJ(XML_doc, er, ee, benefits_node, benclass)
            Case HMIT_E
                Analyse_Benefits = Analyse_Benefits Or SectionE(XML_doc, er, ee, benefits_node, benclass)
            Case HMIT_F
                Analyse_Benefits = Analyse_Benefits Or SectionF(XML_doc, er, ee, benefits_node)
                'Call SectionF(XML_doc, er, ee, benclass)
            Case HMIT_G
                'Van exists by default therefore check if any van benefits
                If ee.AnyVanBenefit Then Analyse_Benefits = Analyse_Benefits Or SectionG(XML_doc, er, ee, benefits_node)
            Case HMIT_H
                Analyse_Benefits = Analyse_Benefits Or SectionH(XML_doc, er, ee, benefits_node)
            Case HMIT_I
                Analyse_Benefits = Analyse_Benefits Or SectionIandK(XML_doc, er, ee, benefits_node, benclass)
            Case HMIT_J
                Analyse_Benefits = Analyse_Benefits Or SectionDandJ(XML_doc, er, ee, benefits_node, benclass)
            Case HMIT_K
                Analyse_Benefits = Analyse_Benefits Or SectionIandK(XML_doc, er, ee, benefits_node, benclass)
            Case HMIT_L
                Analyse_Benefits = Analyse_Benefits Or SectionAandL(XML_doc, er, ee, benefits_node, benclass)
            'Case HMIT_M
            '    Analyse_Benefits = Analyse_Benefits Or SectionM(XML_doc, er, ee, benefits_node)
            Case HMIT_M
                Analyse_Benefits = Analyse_Benefits Or SectionM(XML_doc, er, ee, benefits_node)
            Case HMIT_N
                Analyse_Benefits = Analyse_Benefits Or SectionN(XML_doc, er, ee, benefits_node)
        End Select
    Next
Analyse_Benefits_END:
  Call xReturn("Analyse_Benefits")
    Exit Function
Analyse_Benefits_ERR:
    Call ErrorMessage(ERR_ERROR, Err, "Analyse_Benefits", "Analyse_Benefits", "Error while analysing benefits in PAYEOnline.")
    Resume Analyse_Benefits_END
    Resume
End Function

Private Function SectionB(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
  Dim BenArr() As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node_payments As IXMLDOMNode, node_payment As IXMLDOMNode
  
  On Error GoTo SectionB_ERR
  
  Call xSet("SectionB")
  
  ReDim BenArr(1 To 2)
  
  BenArr(1) = BC_PAYMENTS_ON_BEFALF_B
  BenArr(2) = BC_TAX_NOTIONAL_PAYMENTS_B
  
  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr) Then
    SectionB = True
    Set node_payments = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "Payments", "", "Type", "B")
    
    ReDim BenArr(1 To 1)
    BenArr(1) = BC_PAYMENTS_ON_BEFALF_B
    
    If (SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc)) Then
       Set node_payment = PAYEAddXMLChild(XML_doc, node_payments, "", "Payment", "")
       Call PAYEAddXMLChild(XML_doc, node_payment, "", "Desc", IRDesc)
       
       If (Len(Description) > 0) Then
         Call PAYEAddXMLChild(XML_doc, node_payment, "", "Other", Description, , , L_MAX_LONG_STANDARD_FIELD_LENGTH)
       End If

       Call PAYEAddXMLChild(XML_doc, node_payment, "", S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(benefit))
       
    End If
    BenArr(1) = BC_TAX_NOTIONAL_PAYMENTS_B

    If (SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc)) Then
      Call PAYEAddXMLChild(XML_doc, node_payments, "", "Tax", GetXMLAmount(benefit))
    End If
    
  End If
SectionB_END:
  Call xReturn("SectionB")
  Exit Function
SectionB_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SectionB", "SectionB", "Error in Type B Benefits in PAYEOnline.")
  Resume SectionB_END
  Resume
End Function


Private Function SectionAandL(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode, benclass As BEN_CLASS) As Boolean
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node As IXMLDOMNode, node_asset As IXMLDOMNode
  Dim sText As String, sType As String, sValueDesc As String
  On Error GoTo err_Err
  
  Call xSet("SectionAandL")
  
  Select Case benclass
    Case BC_ASSETSTRANSFERRED_A
      BenArr(1) = BC_ASSETSTRANSFERRED_A
      sText = "Transferred"
      sValueDesc = S_PAYE_ONLINE_COST_OR_AMOUNT_FORGONE
    Case BC_ASSETSATDISPOSAL_L
      BenArr(1) = BC_ASSETSATDISPOSAL_L
      sText = "AssetsAvail"
      
      sValueDesc = "AnnValProRata"
  End Select
  
  sType = p11d32.Rates.HMITSectionToHMITCode(p11d32.Rates.BenClassTo(benclass, BCT_HMIT_SECTION))
      
  
  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
    SectionAandL = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", sText, "", "Type", sType)
    Set node_asset = PAYEAddXMLChild(XML_doc, node, "", "Asset", "")
    Call PAYEAddXMLChild(XML_doc, node_asset, "", "Desc", IRDesc)
    If (Len(Description) > 0) Then
      Call PAYEAddXMLChild(XML_doc, node_asset, "", "Other", Description, , , L_MAX_LONG_STANDARD_FIELD_LENGTH)
    End If
    Call PAYEAddXMLChild(XML_doc, node_asset, "", sValueDesc, GetXMLAmount(value))
    Call Standard_XML_Out(XML_doc, node_asset, "", False, 0, True, MadeGood, True, benefit)
  End If
  
err_End:
  Call xReturn("SectionAandL")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionAandL", "SectionAandL", "Error in Type A or L Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function

Private Function SectionC(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
    'may be used for Section C
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node As IXMLDOMNode
  On Error GoTo err_Err
  
  Call xSet("SectionC")
  
  BenArr(1) = BC_VOUCHERS_AND_CREDITCARDS_C
  
  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
    SectionC = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "VouchersOrCCs", "", "Type", "C")
    Call PAYEAddXMLChild(XML_doc, node, "", S_PAYE_ONLINE_GROSS_OR_AMOUNT_FORGONE, GetXMLAmount(value))
    Call Standard_XML_Out(XML_doc, node, "", False, value, True, MadeGood, True, benefit)
  End If
  
err_End:
  Call xReturn("SectionC")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionC", "SectionC", "Error in Type C Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function
Private Function SectionG(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
  Dim nsv As IBenefitClass
  Dim node As IXMLDOMNode
  
  'Case BC_NONSHAREDVANS_G
  'BenArr(1) = BC_NONSHAREDVANS_G
  'sText = "Vans"
  'sType = "G"
  'sBenefitText = "CashEquiv"
   
  If (ee.AnyVanBenefit) Then
    SectionG = True
    Set nsv = ee.nonSharedVans
    Call nsv.Calculate
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "Vans", "", "Type", "G")
    Call PAYEAddXMLChild(XML_doc, node, "", S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(nsv.value(nsvans_benefit_van_only)))
    Call PAYEAddXMLChild(XML_doc, node, "", "Fuel" & S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(nsv.value(nsvans_fuel_benefit)))
  End If
End Function
Private Function SectionDandJ(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode, benclass As BEN_CLASS) As Boolean
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node As IXMLDOMNode
  Dim sText As String, sType As String, sBenefitText As String
  On Error GoTo err_Err
  
  Call xSet("SectionDandGandJ")
  Select Case benclass
    Case BC_LIVING_ACCOMMODATION_D
      BenArr(1) = BC_LIVING_ACCOMMODATION_D
      sText = "LivingAccom"
      sType = "D"
      sBenefitText = S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT
    Case BC_QUALIFYING_RELOCATION_J
      BenArr(1) = BC_QUALIFYING_RELOCATION_J
      sText = "Relocation"
      sType = "J"
      sBenefitText = "Excess"
    Case Else
      Call Err.Raise(Err.Number)
  End Select
  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
    SectionDandJ = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", sText, "", "Type", sType)
    Call PAYEAddXMLChild(XML_doc, node, "", sBenefitText, GetXMLAmount(benefit))
  End If
  
err_End:
  Call xReturn("SectionDandGandJ")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionDandGandJ", "SectionDandGandJ", "Error in Type D, G or J Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function

'Private Function SectionM(XML_doc As DOMDocument, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
'  Dim BenArr(1 To 1) As BEN_CLASS
'  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant
'  Dim Node As IXMLDOMNode
'  On Error GoTo err_Err
'
'  Call xSet("SectionM")
'  BenArr(1) = BC_SHARES_M
'  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr) Then
'    SectionM = True
'    Set Node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "SharesInd", PAYE_Yes, "Type", "M")
'    'Set node = PAYEAddXMLChild(XML_doc, benefits_node, "", "SharesInd", PAYE_Yes)
'  End If
'
'err_End:
'  Call xReturn("SectionM")
'  Exit Function
'err_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "SectionM", "SectionM", "Error in Type M Benefits in PAYEOnline.")
'  Resume err_End
'  Resume
'End Function

Private Function SectionE(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode, benclass As BEN_CLASS) As Boolean
  Dim ben As IBenefitClass
  Dim NetBenefit As Variant
  Dim node As IXMLDOMNode
  Dim i As Long
  On Error GoTo err_Err
  
  Call xSet("SectionE")
  NetBenefit = 0
  
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = benclass Then
        NetBenefit = NetBenefit + ben.value(ITEM_BENEFIT)
      End If
    End If
  Next
  
  If NetBenefit > 0 Then
    SectionE = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "MileageAllow", "", "Type", "E")
    Call PAYEAddXMLChild(XML_doc, node, "", "TaxablePmt", GetXMLAmount(NetBenefit))
  End If
  
err_End:
  Call xReturn("SectionE")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionE", "SectionE", "Error in Type E Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function

Private Function SectionIandK(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode, benclass As BEN_CLASS) As Boolean
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node As IXMLDOMNode
  Dim sText As String, sType As String
  On Error GoTo err_Err
  
  Call xSet("SectionIandK")
  
  Select Case benclass
    Case BC_PRIVATE_MEDICAL_I
      BenArr(1) = BC_PRIVATE_MEDICAL_I
      sText = "Medical"
      sType = "I"
    Case BC_SERVICES_PROVIDED_K
      BenArr(1) = BC_SERVICES_PROVIDED_K
      sText = "Services"
      sType = "K"
    Case Else
      Call Err.Raise(Err.Number)
  End Select
  
  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
    SectionIandK = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", sText, "", "Type", sType)
    Call Standard_XML_Out(XML_doc, node, "", True, value, True, MadeGood, True, benefit)
  End If
  
err_End:
  Call xReturn("SectionIandK")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionIandK", "SectionIandK", "Error in Type I or K Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function

  
Private Function SectionF(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean

  Dim BenArr() As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node_cars As IXMLDOMNode, node_car As IXMLDOMNode
  Dim ben As IBenefitClass
  Dim ben_err As IBenefitClass
  Dim car As CompanyCar
  Dim i As Long, lErrorNum As Long
  Dim temp As Long
  Dim temp2 As Boolean
  Dim bTopNodeAdded As Boolean

  'EK as only one Datewitdrawn box in 2003, has to be altered 2004
  Dim dDateWithdrawn As Date
  Dim lNoOfCars As Long
  Dim bFuelReinstated As Boolean
  
  On Error GoTo SectionF_ERR
  Call xSet("SectionF")
  
  ReDim BenArr(1 To 2)


  Set ben_err = er


  dDateWithdrawn = UNDATED
  lNoOfCars = 0


  'iterate through all cars
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = BC_COMPANY_CARS_F Then
        Set car = ben
        If CheckBen(er, ee, ben) Then
        
          Call CompanyCarAvailableDatesValid(ben)
          Call CompanyCarCheckDateFuelAvaialbleTo(ben)
          
          If Not bTopNodeAdded Then
            Set node_cars = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "Cars", "", "Type", "F")
            bTopNodeAdded = True
          End If

          Call CreateXMLCar(XML_doc, node_cars, ben, POT_P11D)
          lNoOfCars = lNoOfCars + 1

          
        End If
      End If
    End If
  Next

  If lNoOfCars > 0 Then

    'Car benefit total
    ReDim BenArr(1 To 1)
    BenArr(1) = BC_COMPANY_CARS_F
    Call SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr)
    Call PAYEAddXMLChild(XML_doc, node_cars, "", "TotalCarsOrRelevantAmt", GetXMLAmount(benefit))

    'Fuel benefit Total
    BenArr(1) = BC_FUEL_F
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr) Then
      Call PAYEAddXMLChild(XML_doc, node_cars, "", "TotalFuelOrRelevantAmt", GetXMLAmount(benefit))
    End If
    SectionF = True
  End If

'cad my god this code is awfull, i can't understand how the programers can write such ad stuff...

SectionF_END:
  Call xReturn("SectionF")
  Exit Function
SectionF_ERR:
  If Err.Number = ERR_CAR_LIST_PRICE Then
    Call ErrorMessage(ERR_CAR_LIST_PRICE, Err, "SectionF", FilterMessageTitle(ben_err.value(employer_Name_db), ee.FullName, "Section F", "Company Cars"), Err.Description)
    Resume Next
  ElseIf Err.Number = ERR_NEEDCO2ORENGINESIZE Then
    Call ErrorMessage(ERR_NEEDCO2ORENGINESIZE, Err, "SectionF", FilterMessageTitle(ben_err.value(employer_Name_db), ee.FullName, "Section F", "Company Cars"), Err.Description)
    Resume Next
  ElseIf Err.Number = ERR_CAR_FUEL_AVAILABLE_TO_INVALID Or Err.Number = ERR_MULTIPLEFUELWITHDRAWN Or Err.Number = ERR_CAR_DATES_INVALID Then
    Call ErrorMessage(Err.Number, Err, "SectionF", FilterMessageTitle(ben_err.value(employer_Name_db), ee.FullName, "Section F", "Company Cars"), Err.Description)
    Resume Next
  Else
    Call ErrorMessage(ERR_ERROR, Err, "SectionF", FilterMessageTitle(er, ee), "Error getting the section F line.")
    Resume SectionF_END
    Resume
  End If
End Function

Private Function SectionH(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
  Dim BenArr(1 To 1) As BEN_CLASS
  Dim i As Long
  Dim bCreatedNode As Boolean
  Dim node As IXMLDOMNode, node_loan As IXMLDOMNode
  Dim dDate As Date, dDischarge As Date
  Dim loans As loans
  Dim ben As IBenefitClass
  On Error GoTo err_Err
  Call xSet("SectionH")
  
  i = ee.GetLoansBenefitIndex
  If i = 0 Then GoTo err_End
  Set loans = ee.benefits(i)
  If loans Is Nothing Then GoTo err_End
  
    For i = 1 To loans.loans.Count
      Set ben = loans.loans(i)
      If Not ben Is Nothing Then
        If CheckBen(er, ee, ben) Then
          SectionH = True
          If Not bCreatedNode Then
            Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "Loans", "", "Type", "H")
            bCreatedNode = True
          End If
          Set node_loan = PAYEAddXMLChild(XML_doc, node, "", "Loan", "")
          Call PAYEAddXMLChild(XML_doc, node_loan, "", "Joint", CStr(ben.value(ln_nborrowers_db)))
          Call PAYEAddXMLChild(XML_doc, node_loan, "", "InitOS", GetXMLAmount(ben.value(ln_OpenOutstanding)))
          Call PAYEAddXMLChild(XML_doc, node_loan, "", "FinalOS", GetXMLAmount(ben.value(ln_CloseOutstanding)))
          Call PAYEAddXMLChild(XML_doc, node_loan, "", "MaxOS", GetXMLAmount(ben.value(ln_MaxOutstandingAtAnyPoint)))
          Call PAYEAddXMLChild(XML_doc, node_loan, "", "IntPaid", GetXMLAmount(ben.value(ln_InterestPaid_db)))

          dDate = ben.value(ln_MadeDate)
          If dDate <> UNDATED Then
            Call PAYEAddXMLChild(XML_doc, node_loan, "", "Date", DateString2(dDate, "", "YYYY-MM-DD"))
          End If
          dDischarge = ben.value(ln_DischargedDate)
          If dDischarge <> UNDATED Then Call PAYEAddXMLChild(XML_doc, node_loan, "", "Discharge", DateString2(dDischarge, "", "YYYY-MM-DD"))
          Call PAYEAddXMLChild(XML_doc, node_loan, "", S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(ben.value(ITEM_BENEFIT)))
        End If
      End If
    Next

err_End:
  Call xReturn("SectionH")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionH", "SectionH", "Error in Type H Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function
Public Function SumBenefit(ByRef Description As Variant, ByRef value As Variant, ByRef MadeGood As Variant, ByRef benefit As Variant, ey As IBenefitClass, ee As Employee, benefits As ObjectList, BenArr() As BEN_CLASS, Optional IRDesc As Variant) As Boolean
  Dim b As Boolean
  
  b = Validations.SumBenefit(Description, value, MadeGood, benefit, ey, ee, benefits, BenArr(), IRDesc)
  If (b) Then
    Description = MarkUpField(Description)
  End If
  SumBenefit = b
  
End Function
Private Sub SectionMTransLateIRDesc(ByRef IRDesc, ByRef bUCase)
  If StrComp(IRDesc, S_IR_DESC_M_C1A_SUBS_AND_FEES, vbTextCompare) = 0 Then
    IRDesc = S_IR_DESC_M_C1A_PAYE_ONLINE_SUBS_AND_FEES
  ElseIf StrComp(IRDesc, S_IR_DESC_M_C1A_ED_ASS, vbTextCompare) = 0 Then
    IRDesc = S_IR_DESC_M_C1A_PAYE_ONLINE_ED_ASS
    bUCase = True 'IK 02/04 node should have uppercase chars
  End If
End Sub
Private Function SectionM(XML_doc As DOMDocument60, er As Employer, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
  Dim BenArr() As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node As IXMLDOMNode, node_Class1A As IXMLDOMNode, node_NonClass1A As IXMLDOMNode
  Dim sText As String, sType As String, sValueDesc As String
  Dim bUCase As Boolean
  On Error GoTo err_Err

  Call xSet("SectionM")
  
  ReDim BenArr(1 To 3)
  BenArr(1) = BC_CLASS_1A_M
  BenArr(2) = BC_NON_CLASS_1A_M
  BenArr(3) = BC_INCOME_TAX_PAID_NOT_DEDUCTED_M
  bUCase = False
  
  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
    SectionM = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "Other", "", "Type", p11d32.Rates.HMITSectionToHMITCode(HMIT_M))
    If (ee.PersonnelNumber = "54664") Then
      SectionM = True
    End If
    
    ReDim BenArr(1 To 1)
    BenArr(1) = BC_CLASS_1A_M
    If (SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc)) Then

      Set node_Class1A = PAYEAddXMLChild(XML_doc, node, "", "Class1A", "")
      Call SectionMTransLateIRDesc(IRDesc, bUCase)
      Call PAYEAddXMLChild(XML_doc, node_Class1A, "", "Desc", IRDesc, bUCase)
      bUCase = False
      If (Len(Description) > 0) Then
        Call PAYEAddXMLChild(XML_doc, node_Class1A, "", "Other", Description, , , L_MAX_LONG_STANDARD_FIELD_LENGTH)
      End If
      Call Standard_XML_Out(XML_doc, node_Class1A, "", True, value, True, MadeGood, True, benefit)
    End If

    BenArr(1) = BC_NON_CLASS_1A_M
    If (SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc)) Then
      Set node_NonClass1A = PAYEAddXMLChild(XML_doc, node, "", "NonClass1A", "")
      
      Call SectionMTransLateIRDesc(IRDesc, bUCase)

      Call PAYEAddXMLChild(XML_doc, node_NonClass1A, "", "Desc", IRDesc)
      If (Len(Description) > 0) Then
        Call PAYEAddXMLChild(XML_doc, node_NonClass1A, "", "Other", Description, , , L_MAX_LONG_STANDARD_FIELD_LENGTH)
      End If
      Call Standard_XML_Out(XML_doc, node_NonClass1A, "", True, value, True, MadeGood, True, benefit)
    End If

    BenArr(1) = BC_INCOME_TAX_PAID_NOT_DEDUCTED_M
    If (SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc)) Then
      Call PAYEAddXMLChild(XML_doc, node, "", "TaxPaid", GetXMLAmount(benefit))
    End If
  End If
 
err_End:
  Call xReturn("SectionM")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionM", "SectionM", "Error in Type M Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function

Private Function SectionN(XML_doc As DOMDocument60, er As IBenefitClass, ee As Employee, benefits_node As IXMLDOMNode) As Boolean
  Dim BenArr() As BEN_CLASS
  Dim value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDesc As Variant
  Dim node As IXMLDOMNode, node_genbus As IXMLDOMNode, node_ent As IXMLDOMNode
  Dim node_travel As IXMLDOMNode, node_HomeTel As IXMLDOMNode, node_NonQualRes As IXMLDOMNode
  Dim node_other As IXMLDOMNode
  Dim sText As String, sType As String, sValueDesc As String
  On Error GoTo err_Err

  Call xSet("SectionN")
  
  ReDim BenArr(1 To 7)
  BenArr(1) = BC_TRAVEL_AND_SUBSISTENCE_N
  BenArr(2) = BC_ENTERTAINMENT_N
  BenArr(3) = BC_GENERAL_EXPENSES_BUSINESS_N 'CAD 2017 - remove, and from menu, depends on P11D form?
  BenArr(4) = BC_PHONE_HOME_N
  BenArr(5) = BC_NON_QUALIFYING_RELOCATION_N
  BenArr(6) = BC_OOTHER_N
  BenArr(7) = BC_CHAUFFEUR_OTHERO_N

  If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
    SectionN = True
    Set node = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", "ExpPaid", "", "Type", p11d32.Rates.HMITSectionToHMITCode(HMIT_N))

    'travel and subsistance
    ReDim BenArr(1 To 1)
    BenArr(1) = BC_TRAVEL_AND_SUBSISTENCE_N
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then

      Set node_travel = PAYEAddXMLChild(XML_doc, node, "", "TravAndSub", "")
      Call SectionN_XMLOut(XML_doc, node_travel, "", value, MadeGood, benefit)
    End If
    'entertainment
    BenArr(1) = BC_ENTERTAINMENT_N
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
      Set node_ent = PAYEAddXMLChildWithAttribute(XML_doc, node, "", "Ent", "", "TradingOrgInd", PAYEBoolYesNo(er.value(employer_CT_db)))
      'Call PAYEAddXMLChild(XML_doc, node_ent, "", "TradingOrdInd", er.value(employer_CT))
      Call SectionN_XMLOut(XML_doc, node_ent, "", value, MadeGood, benefit)
    End If
    'business travel
    BenArr(1) = BC_GENERAL_EXPENSES_BUSINESS_N
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
      Set node_genbus = PAYEAddXMLChild(XML_doc, node, "", "General", "")
      Call SectionN_XMLOut(XML_doc, node_genbus, "", value, MadeGood, benefit)
    End If
    'home telephone
    BenArr(1) = BC_PHONE_HOME_N
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
      Set node_HomeTel = PAYEAddXMLChild(XML_doc, node, "", "HomeTel", "")
      Call SectionN_XMLOut(XML_doc, node_HomeTel, "", value, MadeGood, benefit)
    End If
    'non qualifying relocation
    BenArr(1) = BC_NON_QUALIFYING_RELOCATION_N
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
      Set node_NonQualRes = PAYEAddXMLChild(XML_doc, node, "", "NonQualRel", "")
      Call SectionN_XMLOut(XML_doc, node_NonQualRes, "", value, MadeGood, benefit)
    End If
    'non other and chauffeur
    ReDim BenArr(1 To 2)
    BenArr(1) = BC_OOTHER_N
    BenArr(2) = BC_CHAUFFEUR_OTHERO_N
    If SumBenefit(Description, value, MadeGood, benefit, er, ee, ee.benefits, BenArr, IRDesc) Then
      Set node_other = PAYEAddXMLChild(XML_doc, node, "", "Other", "")
      Call PAYEAddXMLChild(XML_doc, node_other, "", "Desc", IRDesc)
      Call SectionN_XMLOut(XML_doc, node_other, "", value, MadeGood, benefit)
    End If
  End If
    
err_End:
  Call xReturn("SectionN")
  Exit Function
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SectionN", "SectionN", "Error in Type O or L Benefits in PAYEOnline.")
  Resume err_End
  Resume
End Function


Private Sub SectionN_XMLOut(XML_doc As DOMDocument60, top_node As IXMLDOMNode, root As String, value As Variant, MadeGood As Variant, benefit As Variant)

      Call PAYEAddXMLChild(XML_doc, top_node, root, S_PAYE_ONLINE_COST_OR_AMOUNT_FORGONE, GetXMLAmount(value))
      Call PAYEAddXMLChild(XML_doc, top_node, root, "MadeGood", GetXMLAmount(MadeGood))
      Call PAYEAddXMLChild(XML_doc, top_node, root, S_PAYE_ONLINE_TAXABLE_PAYMENT_OR_RELEVANT_AMOUNT, GetXMLAmount(benefit))

End Sub

Private Sub Standard_XML_Out(XML_doc As DOMDocument60, top_node As IXMLDOMNode, root As String, ValueRequired As Boolean, value As Variant, MadeGoodRequired As Boolean, MadeGood As Variant, BenefitRequired As Boolean, benefit As Variant)
    If ValueRequired Then
      Call PAYEAddXMLChild(XML_doc, top_node, root, S_PAYE_ONLINE_COST_OR_AMOUNT_FORGONE, GetXMLAmount(value))
    End If
    If MadeGoodRequired Then
      Call PAYEAddXMLChild(XML_doc, top_node, root, "MadeGood", GetXMLAmount(MadeGood))
    End If
    If BenefitRequired Then
      Call PAYEAddXMLChild(XML_doc, top_node, root, S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(benefit))
    End If

End Sub


Private Sub SortBenArray(ByRef Arr() As BEN_CLASS)

    Dim i As Integer, j As Integer, iStart As Integer, iEnd As Integer
    Dim iLeft As Integer, iRight As Integer
    Dim temp As BEN_CLASS
    
    iStart = LBound(Arr)
    iEnd = UBound(Arr)
    
    For i = iStart To iEnd
        If i = iEnd Then GoTo SortBenArray_End
        For j = i + 1 To iEnd
            'IK 02/04/2004 - Need to sort array in ascending order but Ben_Class Enum is not in correct order
            iLeft = Asc(p11d32.Rates.BenClassTo(Arr(j), BCT_HMIT_SECTION_STRING))
            iRight = Asc(p11d32.Rates.BenClassTo(Arr(i), BCT_HMIT_SECTION_STRING))
            If iLeft < iRight Then
                temp = Arr(i)
                Arr(i) = Arr(j)
                Arr(j) = temp
            End If
        Next
    Next
    
SortBenArray_End:

End Sub

Private Function SumBenefitIRDesc(ByRef Description As Variant, ByRef sIRDesc As Variant, ByRef value As Variant, ByRef MadeGood As Variant, ByRef benefit As Variant, ey As IBenefitClass, ee As Employee, benefits As ObjectList, BenArr() As BEN_CLASS, Optional IRDesc As Integer) As Boolean

'  ' This function sums the benefit for one IRDesc and returns the Description as a string too.
'  Dim ben As IBenefitClass
'  Dim i As Long, j As Long
'  Dim lError As Long
'  Dim bc As BEN_CLASS
'  Dim tempDescription As String
'  Dim temp As Integer
'
'  On Error GoTo SumBenefitIRDesc_Err
'  Call xSet("SumBenefitIRDesc")
'
'  value = 0
'  MadeGood = 0
'  Benefit = 0
'  Description = ""
'
'  For i = 1 To benefits.Count
'    Set ben = benefits(i)
'    If Not (ben Is Nothing) Then
'      bc = ben.BenefitClass
'      If BenClassInArray(BenArr, bc) Then
'
'        If ben.value(ITEM_IR_DESC) = IRDesc Then
'          If CheckBen(ey, ee, ben) Then
'            j = j + 1
'            sIRDesc = GetIRDescString(bc, IRDesc)
'            If LCase(IRDesc) = "other" And Len(ben.value(ITEM_DESC)) > 0 Then
'              Description = ben.value(ITEM_DESC)
'            End If
'            Description = ben.value(ITEM_DESC)
'            value = value + ben.value(ITEM_VALUE)
'            MadeGood = MadeGood + ben.value(ITEM_MADEGOOD_NET)
'            Benefit = Benefit + ben.value(ITEM_BENEFIT)
'          End If
'        End If
'      End If
'     Set ben = Nothing
'    End If
'  Next i
'
'  If IRDesc > 0 Then
'    If j > 1 Then
'      sIRDesc = "multiple"
'      Description = ""
'    Else
'      Call GetIRDescString(BenArr(1), IRDesc, sIRDesc)
'    End If
'  End If
'
'  SumBenefitIRDesc = CBool(j)
'
'SumBenefitIRDesc_End:
'  Call xReturn("SumBenefitIRDesc")
'  Exit Function
'SumBenefitIRDesc_Err:
'  If Not ben Is Nothing Then
'    Call ErrorMessage(ERR_ERROR, Err, "SumBenefitIRDesc", FilterMessageTitle(ey.Name, "", ben.Name, ""), "Unable to sum benefits.")
'  Else
'    Call ErrorMessage(ERR_ERROR, Err, "SumBenefitIRDesc", FilterMessageTitle(ey.Name, , "Unknown"), "Unable to sum benefits.")
'  End If
'  SumBenefitIRDesc = False
'  Description = S_ERROR
'  value = S_ERROR
'  MadeGood = S_ERROR
'  Benefit = S_ERROR
'  Resume SumBenefitIRDesc_End
'  Resume
End Function


'returns employee count node
Private Sub CreatePAYEEmployerHead(node_p11d_record_count As IXMLDOMNode, node_p46_record_count As IXMLDOMNode, node_declarations As IXMLDOMNode, XML_doc As DOMDocument60, XMLHeaderNode As IXMLDOMNode, Employers As ObjectList)
  Dim ey As Employer, ee As Employee
  Dim ben As IBenefitClass
  Dim i As Long, j As Long
  Dim node_employer As IXMLDOMNode
  Dim benefit1A As Double, add1A As Double, deduct1A As Double, dNIC As Double, dBenefitsBeforeAddtionsDeductions As Double
  Dim benefitsPotentiallyFor1A As Double
  Dim bAdjustments As Boolean
  Dim p46CarCount As Long
  Dim benEyFirst As IBenefitClass
  Dim node_temp As IXMLDOMNode, node_p11db As IXMLDOMNode, node_class_contributions As IXMLDOMNode, node_adjustments_due As IXMLDOMNode
  Dim node_adjustments_not_due As IXMLDOMNode, node_adjustments As IXMLDOMNode
  Dim sAdditionDescription As String
  Dim sEmployerName As String
  
  
  On Error GoTo err_Err
  
  Call xSet("CreateEmployeeXML")
  
  If (Employers.Count = 0) Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "CreatePAYEEmployerHead", "No employers selected to submit")
  End If
  
  Set benEyFirst = Employers(1)
    
  For i = 1 To Employers.Count
    Set ey = Employers(i)
    Call p11d32.LoadEmployer(ey, False, False)
    Set ben = ey
    Call ben.Calculate
    
    benefitsPotentiallyFor1A = benefitsPotentiallyFor1A + ben.value(employer_TotalBenefitsPotentiallySubjectToClass1A)
    add1A = add1A + ben.value(employer_NIC_AdjustmentAdd)
    deduct1A = deduct1A + ben.value(employer_NIC_AdjustmentDeduct)
    bAdjustments = bAdjustments Or deduct1A Or add1A
    
    dNIC = dNIC + ben.value(ITEM_NIC_CLASS1A_BENEFIT)
    benefit1A = benefit1A + ben.value(ITEM_BENEFIT_SUBJECT_TO_CLASS1A)
  Next
   
  Set node_employer = PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "Employer", "")
  sEmployerName = Trim$(benEyFirst.value(employer_Name_db))
  If Len(sEmployerName) > 35 Then sEmployerName = Left$(sEmployerName, 35)
  sEmployerName = MarkUpField(sEmployerName)
  Call PAYEAddXMLChild(XML_doc, node_employer, "", "Name", sEmployerName)
  If p11d32.PAYEonline.Type_Selected = POT_P46 Then
    Set node_temp = PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "Declarations", "")
    Set node_declarations = PAYEAddXMLChild(XML_doc, node_temp, "", "P11Dincluded", "") 'IK 31/03/04
    Call PAYEAddXMLChild(XML_doc, node_temp, "", "P46CarDeclaration", PAYE_Yes)
  ElseIf p11d32.PAYEonline.Type_Selected = POT_P11D Or p11d32.PAYEonline.Type_Selected = POT_P11DB_ONLY Then 'P11Db section
    'P11D Declaration
    Set node_declarations = PAYEAddXMLChild(XML_doc, PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "Declarations", ""), "", "P11Dincluded", "")
    
    Set node_p11db = PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "P11Db", "")
    Set node_class_contributions = PAYEAddXMLChildWithAttribute(XML_doc, node_p11db, "", "Class1AcontributionsDue", "", "NICsRate", CStr(p11d32.Rates.value(carNICRate) * 100) & "0")
    
    If bAdjustments Then
      Call PAYEAddXMLChildWithAttribute(XML_doc, node_class_contributions, "", "TotalBenefit", GetXMLAmount(benefitsPotentiallyFor1A), "AdjustmentRequired", "yes")
      Set node_adjustments = PAYEAddXMLChild(XML_doc, node_class_contributions, "", "Adjustments", "")
        Call PAYEAddXMLChild(XML_doc, node_adjustments, "", "TotalBenefit", GetXMLAmount(benefitsPotentiallyFor1A))
        Set node_adjustments_due = PAYEAddXMLChild(XML_doc, node_adjustments, "", "AmountDue", "")
          Call PAYEAddXMLChild(XML_doc, node_adjustments_due, "", "Description", MarkUpField(P11DbAdditionsDescription(benEyFirst, CLng(add1A)), "NONE"), , , L_MAX_LONG_STANDARD_FIELD_LENGTH)
          Call PAYEAddXMLChild(XML_doc, node_adjustments_due, "", "Adjustment", GetXMLAmount(add1A))
        Set node_adjustments_not_due = PAYEAddXMLChild(XML_doc, node_adjustments, "", "AmountNotDue", "")
          Call PAYEAddXMLChild(XML_doc, node_adjustments_not_due, "", "Description", MarkUpField(P11DbDeductionsDescription(benEyFirst, CLng(deduct1A)), "NONE"), , , L_MAX_LONG_STANDARD_FIELD_LENGTH)
          Call PAYEAddXMLChild(XML_doc, node_adjustments_not_due, "", "Adjustment", GetXMLAmount(deduct1A))
      
        Call PAYEAddXMLChild(XML_doc, node_adjustments, "", "Total", GetXMLAmount(benefit1A))
        Call PAYEAddXMLChild(XML_doc, node_adjustments, "", "Payable", GetXMLAmount(dNIC, True))
    Else
      Call PAYEAddXMLChild(XML_doc, node_class_contributions, "", "TotalBenefit", GetXMLAmount(benefitsPotentiallyFor1A))
      Call PAYEAddXMLChild(XML_doc, node_class_contributions, "", "NICpayable", GetXMLAmount(dNIC, True))
    End If
  End If

      
  'Call PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "P9DrecordCount", "0") 'cad 2017 removed
  Set node_p11d_record_count = PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "P11DrecordCount", "0")
  Set node_p46_record_count = PAYEAddXMLChild(XML_doc, XMLHeaderNode, "", "P46CarRecordCount", "0")


err_End:
  Call xReturn("CreatePAYEEmplyerHead")
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CreatePAYEEmplyerHead", "CreatePAYEEmplyerHead", "Error creating Head Employee XML in PAYEOnline.")
  Resume err_End
  Resume
End Sub
'returns if invalid NI

Private Function NIValid(ee As IBenefitClass)
   NIValid = ValidateNI(ee.value(ee_NINumber_db), False) = STANDARD
End Function
Public Sub CreateEmployeeXML(ByVal XML_doc As DOMDocument60, benefits_node As IXMLDOMNode, ben As IBenefitClass, ByVal PayeOnlineType As Long)
    Dim sTxt As String
    Dim ee As Employee
    Dim v As Variant
    Dim lNum As Long
    Dim node_temp As IXMLDOMNode
    Dim node_name As IXMLDOMNode
    Dim sNodeName As String
    Dim node_ee_Details As IXMLDOMNode
    On Error GoTo err_Err
  
    Call xSet("CreateEmployeeXML")

    'Type sensitive entries to go here
    Select Case PayeOnlineType
        Case POT_P46
            sNodeName = "EmployeeDetails"
        Case POT_P11D
            sNodeName = "Employee"
    End Select

    'Employee Details
    If ben.value(ee_Director_db) And PayeOnlineType = POT_P11D Then
      Set node_ee_Details = PAYEAddXMLChildWithAttribute(XML_doc, benefits_node, "", sNodeName, "", "DirInd", PAYE_Yes)
    Else
      Set node_ee_Details = PAYEAddXMLChild(XML_doc, benefits_node, "", sNodeName, "")
    End If
    
    Set node_name = PAYEAddXMLChild(XML_doc, node_ee_Details, "", "Name", "")

    '1. Title
        sTxt = ben.value(ee_Title_db)
        If PAYEFieldValid(sTxt, PAYE_TITLE) And Len(sTxt) > 0 Then
            Call PAYEAddXMLChild(XML_doc, node_name, "", "Ttl", sTxt, True)
        End If

    '2. FORENAME
        sTxt = MarkUpField(TrimMaxLength(ben.value(ee_Firstname_db), L_MAX_LONG_FIELD_LENGTH)) ' MarkUpField(Trim$(ben.value(ee_Firstname_db)))
        lNum = InStr(1, sTxt, " ")
        If lNum Then sTxt = Mid(sTxt, 1, lNum - 1)
        If PAYEFieldValid(sTxt, PAYE_FORENAME) Then
            Call PAYEAddXMLChild(XML_doc, node_name, "", "Fore", sTxt, True, , L_MAX_LONG_STANDARD_FIELD_LENGTH)
        Else 'IK 02/04/2004
            Call PAYEAddXMLChild(XML_doc, node_name, "", "Fore", S_UNKNOWN, True)
        End If


      '2nd. FORENAME
      sTxt = MarkUpField(TrimMaxLength(ben.value(ee_Initials_db), L_MAX_LONG_FIELD_LENGTH))
      lNum = InStr(1, sTxt, " ")
      If lNum Then sTxt = Mid(sTxt, 1, lNum - 1)
      If Len(sTxt) > 0 Then
        If PAYEFieldValid(sTxt, PAYE_FORENAME) Then
          Call PAYEAddXMLChild(XML_doc, node_name, "", "Fore", sTxt, True, , L_MAX_LONG_FIELD_LENGTH)
        End If
      End If

      
    'SURNAME
        sTxt = MarkUpField(TrimMaxLength(ben.value(ee_Surname_db), L_MAX_LONG_FIELD_LENGTH))
        lNum = InStr(1, sTxt, " ")
        If lNum Then sTxt = Mid(sTxt, 1, lNum - 1)
        If PAYEFieldValid(sTxt, PAYE_SURNAME) Then
          Call PAYEAddXMLChild(XML_doc, node_name, "", "Sur", MarkUpField(sTxt), True, , L_MAX_LONG_FIELD_LENGTH)
        End If
    'Personnel Number
        If PayeOnlineType = POT_P11D Then
          Call PAYEAddXMLChild(XML_doc, node_ee_Details, "", "WksNo", MarkUpFieldPersonnelNumber(ben.value(ee_PersonnelNumber_db)), True)
        End If
        'NI NUMBER
        Set ee = ben

        If Not NIValid(ee) Then
          sTxt = ""
        Else
          sTxt = UCASE$(TrimEx(ben.value(ee_NINumber_db)))
        End If
        
        If (NIValid(ee)) And (Len(sTxt) = 8) Then
          sTxt = sTxt & " "
        End If
        
        'only include if we have one
        If (Len(Trim$(sTxt)) > 0) Then
          Call PAYEAddXMLChild(XML_doc, node_ee_Details, "", "NINO", sTxt, True, False)
        End If

        If Not NIValid(ee) Then
          'DOB
          v = ben.value(ee_DOB_db)
          If (v = UNDATED) Then
            v = ""
          End If
          sTxt = DateString2(v, "", "YYYY-MM-DD")
          If (Len(sTxt) = 0) Then
            sTxt = DateString2(TryConvertDateDMY("01/01/1901"), "", "YYYY-MM-DD")
          End If
          Call PAYEAddXMLChild(XML_doc, node_ee_Details, "", "BirthDate", sTxt, True, False)
          'GENDER
          sTxt = ""
          If (ben.value(ee_Gender_db) = S_GENDER_MALE) Then
            sTxt = "male"
          ElseIf (ben.value(ee_Gender_db) = S_GENDER_FEMALE) Then
            sTxt = "female"
          End If
                  
          Call PAYEAddXMLChild(XML_doc, node_ee_Details, "", "Gender", sTxt, False, False)
        End If
  

err_End:
  Call xReturn("CreateEmployeeXML")
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CreateEmployeeXML", "CreateEmployeeXML", "Error creating Employee XML in PAYEOnline.")
  Resume err_End
  Resume
End Sub
Private Function MarkUpField(ByVal Field As String, Optional ByVal IfZeroLengthUseThis As String = "") As String
  Dim s As String
  Dim i As Long, iLen As Long, iAsc As Long
  Dim s1 As String, s2 As String
  
  s = Field
    
  
  's = Replace(s, ">", "&gt;")
  's = Replace(s, "<", "&lt;")
  's = Replace(s, "&", "&amp;")

  's = Replace(s, """", "&quot;")
  's = Replace(s, "'", "&apos;")

  s1 = Replace(s, ",", "")
  s1 = Replace(s1, ".", "")
  s1 = Replace(s1, "=", "")
  s1 = Replace(s1, "é", "e")
  s1 = Replace(s1, "ó", "o")
  s1 = Replace(s1, "+", " ")
  
  iLen = Len(s1)
  
  s = ""
  For i = 1 To iLen
    s2 = Mid$(s1, i, 1)
    iAsc = Asc(s2)
    If iAsc > 128 Then
       ' "&#" & iAsc & ";" 'igore
    Else
      s = s & s2
    End If
  Next
  
  If Len(IfZeroLengthUseThis) > 0 Then
    If Len(Trim$(s)) = 0 Then
      s = IfZeroLengthUseThis
    End If
  End If
    
  MarkUpField = s

End Function

Private Function MarkUpFieldPersonnelNumber(ByVal Field As String) As String
  Dim s As String
  
  s = MarkUpField(Field)
  s = Replace(s, "(", "")
  s = Replace(s, ")", "")
    
  MarkUpFieldPersonnelNumber = s

End Function
Public Function TrimP46MakeAndModel(make_model As String) As String
    Dim sTxt As String
    Dim i As Long
    Dim s As String, ch As String
    On Error GoTo TrimP46MakeAndModel_ERR
    
    
    sTxt = Trim$(make_model)
    
    If Len(sTxt) > L_MAX_LONG_STANDARD_FIELD_LENGTH Then sTxt = Mid(sTxt, 1, L_MAX_LONG_STANDARD_FIELD_LENGTH)
    s = ""
    For i = 1 To Len(sTxt)
      ch = Mid$(sTxt, i, 1)
      If IsNumeric(ch) Then
        s = s & ch
      ElseIf IsAlphaStr(ch) Then
        s = s & ch
      ElseIf ch = " " Then
        s = s & ch
      End If
    Next
    'remove any only allow alpha,numeric
    TrimP46MakeAndModel = MarkUpField(s)
    
TrimP46MakeAndModel_END:
  Call xReturn("TrimP46MakeAndModel")
  Exit Function
TrimP46MakeAndModel_ERR:
  Resume TrimP46MakeAndModel_END
  Resume
  
End Function

Private Function CarCO2Is0(ben As IBenefitClass) As Boolean
  If ben.value(car_p46NoApprovedCO2Figure_db) = False Then
    CarCO2Is0 = False
  Else
    If (ben.value(car_p46CarbonDioxide_db) = 0) Then
      CarCO2Is0 = False
    Else
      CarCO2Is0 = True
    End If
  End If

End Function
Private Function CarRegisteredBefore1998(ben As IBenefitClass) As Boolean
  CarRegisteredBefore1998 = (ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE))
End Function

Public Sub CreateXMLCar(XML_doc As DOMDocument60, node_cars As IXMLDOMNode, ben As IBenefitClass, ByVal PAYEtype As Long, Optional bIndicator1 As Boolean = False, Optional bIndicator2 As Boolean = False, Optional bIndicator3 As Boolean = False, Optional bIndicator5 As Boolean = False)
    Dim car As CompanyCar
    Dim root As String, Path As String
    Dim temp As String, sTxt As String, att As String
    Dim sMakeAndModel As String, sEngineSize As String
    Dim lRegistrationYear As Long
    Dim lTotalAccessories   As Long
    Dim benfuel As IBenefitClass
    Dim CompanyCars As ObjectList
    Dim lNumCars As Long
    Dim bFuelAvailable As Boolean
    Dim iIndicatorCount As Long
    Dim benEE As IBenefitClass
    Dim bFuelWithdrawnPresent As Boolean
    Dim sFuelTypeLetterP46 As String
    Dim bP46Before98 As Boolean, bP46NoApproved  As Boolean, bP46CO2Emissions As Boolean
    Dim bP46PayeAnyIndicators1to4 As Boolean
    
    Dim top_node As IXMLDOMNode, head_node As IXMLDOMNode
    Dim node_submissionReplacedMultiple As IXMLDOMNode, node_submissionReplaced As IXMLDOMNode, node_submission As IXMLDOMNode
    Dim node_car As IXMLDOMNode, node_cardetails As IXMLDOMNode, node_submissionWithdrawn As IXMLDOMNode
    Dim node_fuel As IXMLDOMNode, node_money As IXMLDOMNode, node_co2 As IXMLDOMNode
    'Set ben = Me
    sEngineSize = CStr(IIf(ben.value(car_enginesize_db) <> 0, ben.value(car_enginesize_db), 0))
    sMakeAndModel = TrimP46MakeAndModel(ben.value(car_Make_db) + " " + ben.value(car_Model_db))
    Set benEE = ben.Parent
        
    
    lTotalAccessories = ben.value(car_AccessoriesOriginal_db) + ben.value(car_AccessoriesNew_db) - ben.value(car_CheapAccessories_db)
   
    bP46PayeAnyIndicators1to4 = False
   
    Select Case PAYEtype
    Case POT_P46
        
        'SUBMISSION REASON
        Set node_submission = PAYEAddXMLChild(XML_doc, node_cars, "", "SubmissionReason", "")
        
        'INDICATOR 1
        If bIndicator1 Then
          iIndicatorCount = iIndicatorCount + 1
          Call PAYEAddXMLChild(XML_doc, node_submission, "", "ProvidedCar", PAYE_Yes)
          bP46PayeAnyIndicators1to4 = True
        End If
        'INDICATOR 2
        If bIndicator2 Then
          iIndicatorCount = iIndicatorCount + 1
          
            Set node_submissionReplaced = PAYEAddXMLChild(XML_doc, node_submission, "", "ReplacedCar", "")
            Call PAYEAddXMLChild(XML_doc, node_submissionReplaced, "", "Replaced", PAYE_Yes)
            bP46PayeAnyIndicators1to4 = True
            'checking for multiple cars
            If ben.value(car_p46ReplacementForMultipleCarsMakeAndModel) Then
                Call PAYEAddXMLChild(XML_doc, node_submissionReplaced, "", "MultipleCarIndicator", PAYE_Yes)
                Set node_submissionReplacedMultiple = PAYEAddXMLChild(XML_doc, node_submissionReplaced, "", "MultipleCars", "")
                sMakeAndModel = TrimP46MakeAndModel(ben.value(car_CarReplacedMake_db) + " " + ben.value(car_CarReplacedModel_db))
                Call PAYEAddXMLChild(XML_doc, node_submissionReplacedMultiple, "", _
                                         "MakeAndModel", sMakeAndModel)
                att = CStr(GetEngineSizeCategory(ben.value(car_CarReplacedEngineSize_db)))  'setting attribute for engine category
                Call PAYEAddXMLChildWithAttribute(XML_doc, node_submissionReplacedMultiple, "", "EngineSize", ben.value(car_CarReplacedEngineSize_db), "Category", att)
            Else
                Call PAYEAddXMLChild(XML_doc, node_submissionReplaced, "", "MultipleCarIndicator", PAYE_No)
            End If
        End If
        
        'INDICATOR 3
        If bIndicator3 Then
          iIndicatorCount = iIndicatorCount + 1
          Call PAYEAddXMLChild(XML_doc, node_submission, "", "SecondCar", PAYE_Yes)
          bP46PayeAnyIndicators1to4 = True
        End If
        
        'we don't know when an employee has become a director or >8500 thus we never fill this in
        'only fill in IfRunningINIDE and employee is director.
        'INDICATOR 4 if not 5
        If (Not bIndicator5) And benEE.value(ee_Director_db) And IsRunningInIDE Then
          Call PAYEAddXMLChild(XML_doc, node_submission, "", "EarningsOrDirector", PAYE_Yes)
          bP46PayeAnyIndicators1to4 = True
        End If
        
        'INDICATOR 5
        If bIndicator5 Then
          iIndicatorCount = iIndicatorCount + 1
           If ben.value(car_P46CarProvidedReplaced) Then
            temp = ""
           End If
            
          Set node_submissionWithdrawn = PAYEAddXMLChild(XML_doc, node_submission, "", "CarWithdrawn", "")
          Call PAYEAddXMLChild(XML_doc, node_submissionWithdrawn, "", "Withdrawn", PAYE_Yes)
          Call PAYEAddXMLChild(XML_doc, node_submissionWithdrawn, "", "DateWithdrawn", _
                      DateString2(ben.value(Car_AvailableTo_db), "", "YYYY-MM-DD"))
          sMakeAndModel = TrimP46MakeAndModel(ben.value(car_Make_db) + " " + ben.value(car_Model_db))
          Call PAYEAddXMLChild(XML_doc, node_submissionWithdrawn, "", "MakeAndModel", sMakeAndModel)
          
          Call PAYEAddXMLChildWithAttribute(XML_doc, node_submissionWithdrawn, "", "EngineSize", ben.value(car_enginesize_db), "Category", GetEngineSizeCategory(ben.value(car_enginesize_db)))
          'dont need other car information in this case therefore exit
          GoTo XML_Skip
        End If

        If iIndicatorCount > 1 Then
          If IsRunningInIDE() Then Call ECASE("Invalid P46 car " & ben.Name)
        End If
        
        'CAR DETAILS
        Set node_cardetails = PAYEAddXMLChild(XML_doc, node_cars, "", "CarDetails", "")
        sMakeAndModel = TrimP46MakeAndModel(ben.value(car_Make_db) + " " + ben.value(car_Model_db))
        Call PAYEAddXMLChild(XML_doc, node_cardetails, "", "MakeAndModel", sMakeAndModel)
        'setting attribute for engine category
        att = GetEngineSizeCategory(ben.value(car_enginesize_db))
        Call PAYEAddXMLChildWithAttribute(XML_doc, node_cardetails, "", "EngineSize", sEngineSize, "Category", att)
        Call PAYEAddXMLChild(XML_doc, node_cardetails, "", "DateFirstRegistered", _
                DateString2(ben.value(car_Registrationdate_db), "", "YYYY-MM-DD"))
        Call PAYEAddXMLChild(XML_doc, node_cardetails, "", "FuelType", ben.value(car_p46FuelTypeString), True)
        'CO2
        
        Set node_co2 = PAYEAddXMLChild(XML_doc, node_cars, "", "CO2Emissions", "")

        
        
        'was
        ' If (ben.value(car_p46CarbonDioxide_db) > 0) Then
        '        Call PAYEAddXMLChild(XML_doc, node_co2, "", "Emissions", CStr(ben.value(car_p46CarbonDioxide_db)))
        '    ElseIf ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
        '        Call PAYEAddXMLChild(XML_doc, node_co2, "", "Before1998", PAYE_Yes)
        '    Else
        '        Call PAYEAddXMLChild(XML_doc, node_co2, "", "NoApproved", PAYE_Yes)
        '    End If
        'End If

        'now
        bP46Before98 = CarCO2Is0(ben) And CarRegisteredBefore1998(ben)
        bP46NoApproved = ((Not CarRegisteredBefore1998(ben)) And CarCO2Is0(ben)) Or (ben.value(car_p46FuelType_db) = CCFT_ELECTRIC)
        
        'emissions
        If CarRegisteredBefore1998(ben) Or bP46NoApproved Or (ben.value(car_p46FuelType_db) = CCFT_ELECTRIC) Then
        Else
          bP46CO2Emissions = True
          Call PAYEAddXMLChild(XML_doc, node_co2, "", "Emissions", CStr(ben.value(car_p46CarbonDioxide_db)))
        End If

        If CarRegisteredBefore1998(ben) And (Not bP46CO2Emissions) Then
          Call PAYEAddXMLChild(XML_doc, node_co2, "", "Before1998", PAYE_Yes)
        End If
                        
        'no approved
        If bP46NoApproved Then
           Call PAYEAddXMLChild(XML_doc, node_co2, "", "NoApproved", PAYE_Yes)
        End If
        

        'MONETARY DETAILS
        Set node_money = PAYEAddXMLChild(XML_doc, node_cars, "", "MonetaryDetails", "")
        Call PAYEAddXMLChild(XML_doc, node_money, "", "CarPrice", GetXMLAmount(ben.value(car_ListPrice_db)))
        If ben.value(car_Accessories) > 0 Then Call PAYEAddXMLChild(XML_doc, node_money, "", "AccessoriesPrice", GetXMLAmount(CStr(lTotalAccessories)))
        Call PAYEAddXMLChild(XML_doc, node_money, "", "DateFirstAvailable", DateString2(ben.value(Car_AvailableFrom_db), "", "YYYY-MM-DD"))
        
        
        If (p11d32.AppYear > 2017) Then
        
          'CAD CashForgone, new 2017/18, we don't appear to do the indicator 1-4 checks for other values, doing here but not altering others
          If (bP46PayeAnyIndicators1to4) And (ben.value(car_OPRA_Ammount_Foregone_db) > 0) And (ben.value(car_OPRA_Ammount_Foregone_Used_For_Value)) Then
            Call PAYEAddXMLChild(XML_doc, node_money, "", "CashForegone", GetXMLAmount(ben.value(ITEM_OPRA_AMOUNT_FOREGONE)))
          End If
        End If
        
        
        Call PAYEAddXMLChild(XML_doc, node_money, "", "CapitalContributions", GetXMLAmount(ben.value(car_CapitalContributionRestricted)))
        If ben.value(car_ActualAmountMadeGood) > 0 Then
            'RK Why /is this present on the paper form?
            Select Case ben.value(car_p46PaymentFrequency_db)
                Case P46_PAYMENT_FREQUENCY.P46PF_ANNUALLY
                    temp = "Y"
                Case P46_PAYMENT_FREQUENCY.P46PF_QUARTERLY
                    temp = "Q"
                Case P46_PAYMENT_FREQUENCY.P46PF_MONTHLY
                    temp = "M"
                Case P46_PAYMENT_FREQUENCY.P46PF_WEEKLY
                    temp = "W"
                Case Else
                    temp = "Y"
            End Select
            Call PAYEAddXMLChildWithAttribute(XML_doc, node_money, "", "PrivateUsePayment", GetXMLAmount(ben.value(car_MadeGood_db)), "Interval", CStr(temp))
        End If

        'FUEL DETAILS
        
        
        
        If (ben.value(car_privatefuel_db) = True) And (ben.value(car_p46FuelType_db) <> CCFT_ELECTRIC) Then
            Set node_fuel = PAYEAddXMLChild(XML_doc, node_cars, "", "Fuel", "")
            Call PAYEAddXMLChild(XML_doc, node_fuel, "", "PrivateUse", PAYE_Yes)
            If ben.value(car_requiredmakegood_db) And ben.value(car_actualmadegood_db) Then
               Call PAYEAddXMLChild(XML_doc, node_fuel, "", "FuelPaidByEmployee", PAYE_Yes)
            Else
               Call PAYEAddXMLChild(XML_doc, node_fuel, "", "FuelPaidByEmployee", PAYE_No)
            End If
        ElseIf (ben.value(car_p46FuelType_db) <> CCFT_ELECTRIC) Then
          Set node_fuel = PAYEAddXMLChild(XML_doc, node_cars, "", "Fuel", "")
          Call PAYEAddXMLChild(XML_doc, node_fuel, "", "PrivateUse", PAYE_No)
        End If


        GoTo IBenefitClass_XML_END

    Case POT_P11D
        Set node_car = PAYEAddXMLChild(XML_doc, node_cars, "", "Car", "")

        'make and model
        Call PAYEAddXMLChild(XML_doc, node_car, "", "Make", sMakeAndModel)

        'registration date
        Call PAYEAddXMLChild(XML_doc, node_car, "", "Registered", DateString2(ben.value(car_Registrationdate_db), "", "YYYY-MM-DD"))

        'Available From date - check if first available after tax year start
        If ben.value(Car_AvailableFrom_db) <> UNDATED And ben.value(Car_AvailableFrom_db) > p11d32.Rates.value(TaxYearStart) Or ben.value(car_ForceP46_db) Then
            Call PAYEAddXMLChild(XML_doc, node_car, "", "AvailFrom", DateString2(ben.value(Car_AvailableFrom_db), "", "YYYY-MM-DD"))
        End If

        'Available To date - check if last available before tax year end
        If ben.value(Car_AvailableTo_db) < p11d32.Rates.value(TaxYearEnd) Or ben.value(car_ForceP46_db) Then
            Call PAYEAddXMLChild(XML_doc, node_car, "", "AvailTo", DateString2(ben.value(Car_AvailableTo_db), "", "YYYY-MM-DD"))
        End If
        'Engine size - not needed if fuel type electric
        'If ben.value(car_p46FuelType_db) <> CCFT_ELECTRIC Then  'removed in 2016
          If (ben.value(car_Registrationdate_db) < p11d32.Rates.value(carNOCO2CUTOFFDATE)) Or (ben.value(car_p46NoApprovedCO2Figure_db)) Then
            Call PAYEAddXMLChild(XML_doc, node_car, "", "CC", sEngineSize)
          End If
        'End If

        'Fuel Type
        
        If ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE) Then
            Call PAYEAddXMLChild(XML_doc, node_car, "", "Fuel", ben.value(car_p46FuelTypeString), True)
        End If

        'CO2 Emissions
        

        'post 06/04/2016 and greater

        If (p11d32.AppYear >= 2015) Then
          If (ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE)) Then
            If ben.value(car_p46FuelType_db) = CCFT_ELECTRIC Then 'electric always write zero
              Call PAYEAddXMLChild(XML_doc, node_car, "", "CO2", "0")
            ElseIf Not ben.value(car_p46NoApprovedCO2Figure_db) Then 'approved CO2
              Call PAYEAddXMLChild(XML_doc, node_car, "", "CO2", CStr(ben.value(car_p46CarbonDioxide_db)))
            Else
              Call PAYEAddXMLChild(XML_doc, node_car, "", "NoAppCO2Fig", PAYE_Yes)
            End If
          Else
            Call PAYEAddXMLChild(XML_doc, node_car, "", "NoAppCO2Fig", PAYE_Yes)
          End If
        Else
          If (ben.value(car_p46CarbonDioxide_db) > 0) And (ben.value(car_Registrationdate_db) >= p11d32.Rates.value(carNOCO2CUTOFFDATE)) And (ben.value(car_p46FuelType_db) <> CCFT_ELECTRIC) Then
              Call PAYEAddXMLChild(XML_doc, node_car, "", "CO2", CStr(ben.value(car_p46CarbonDioxide_db)))
          Else
              Call PAYEAddXMLChild(XML_doc, node_car, "", "NoAppCO2Fig", PAYE_Yes)
          End If
        End If
        'List Price
        Call PAYEAddXMLChild(XML_doc, node_car, "", "List", GetXMLAmount(ben.value(car_ListPrice_db)))

        'Accessories
        Call PAYEAddXMLChild(XML_doc, node_car, "", "Accs", GetXMLAmount(ben.value(car_Accessories)))

        'Cap Contribution
        Call PAYEAddXMLChild(XML_doc, node_car, "", "CapCont", GetXMLAmount(ben.value(car_CapitalContributionRestricted)))

        'Private Use payment
        Call PAYEAddXMLChild(XML_doc, node_car, "", "PrivUsePmt", GetXMLAmount(ben.value(ITEM_MADEGOOD_NET)))

        'feul withdrawn
        If ben.value(Car_FuelAvailableTo_calc) <> UNDATED Then
          If ben.value(car_FuelBenefit) > 0 Then
            bFuelWithdrawnPresent = True
            If ben.value(car_fuelreinstated_calc) Then
                Call PAYEAddXMLChildWithAttribute(XML_doc, node_car, "", "FuelWithdrawn", _
                DateString2(ben.value(Car_FuelAvailableTo_calc), "", "YYYY-MM-DD"), "Reinstated", PAYE_Yes)
                          
            Else
              Call PAYEAddXMLChild(XML_doc, node_car, "", "FuelWithdrawn", _
              DateString2(ben.value(Car_FuelAvailableTo_calc), "", "YYYY-MM-DD"))
            End If
          End If
        End If
        'Cash Equiv
        Call PAYEAddXMLChild(XML_doc, node_car, "", S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(ben.value(ITEM_BENEFIT)))
        'Fuel Cash Equiv
        If (p11d32.AppYear >= 2015) Then
          bFuelAvailable = (ben.value(car_FuelBenefit) > 0 And ((ben.value(car_privatefuel_db) = True))) Or bFuelWithdrawnPresent
          bFuelAvailable = bFuelAvailable Or (ben.value(car_FuelOPRA_Ammount_Foregone_Used_For_Value) = True)
          If bFuelAvailable Then
 '            IK 01/04/2004 - Not picking up fuel values
            Call PAYEAddXMLChild(XML_doc, node_car, "", "Fuel" & S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(ben.value(car_FuelBenefit)))
          End If
        Else
          If (ben.value(car_FuelBenefit) > 0 And ((ben.value(car_privatefuel_db) = True) And (ben.value(car_p46FuelType_db) <> CCFT_ELECTRIC))) Or bFuelWithdrawnPresent Then
 '            IK 01/04/2004 - Not picking up fuel values
            Call PAYEAddXMLChild(XML_doc, node_car, "", "Fuel" & S_PAYE_ONLINE_CASH_EQUIV_OR_RELEVANT_AMOUNT, GetXMLAmount(ben.value(car_FuelBenefit)))
          End If
        End If
  
    End Select

XML_Skip:

IBenefitClass_XML_END:
  Call xReturn("IBenefitClass_XML")
  Exit Sub
IBenefitClass_XML_ERR:
  Resume IBenefitClass_XML_END
  Resume

End Sub

Private Function GetEngineSizeCategory(size As Long) As Long
    On Error GoTo GetEngineSizeCategory_ERR
    
      If (size <= 1400) And (size > 0) Then
        GetEngineSizeCategory = 1
      ElseIf (size > 1400) And (size <= 2000) Then
        GetEngineSizeCategory = 2
      ElseIf size > 2000 Then
         GetEngineSizeCategory = 3
      Else
        GetEngineSizeCategory = 4
      End If
    
    
GetEngineSizeCategory_END:
  Call xReturn("GetEngineSizeCategory")
  Exit Function
GetEngineSizeCategory_ERR:
  Resume GetEngineSizeCategory_END
  Resume
  
End Function


Public Property Get Namespace()
  Namespace = NameSpaceByYear()
End Property

'Public Property Get NamespaceCurrentYear()
'  NamespaceCurrentYear = Me.NameSpaceByYear(False)
'End Property

Private Function NameSpaceByYear() As String 'open the xsd and extract the namespace thingy - this is such a mess
  Dim Text As String
  Dim search As String
  Dim p0 As Long, p1 As Long
  Dim nspace As String
  Dim envelopeVersionCurrent As String
  Dim envelopeVersionPrevious As String
  Dim i As Long
  Dim xsdPath As String
  
  
  If ((Me.Type_Selected = POT_P46) And (Not p11d32.JuneRelease)) Then
    xsdPath = p11d32.SystemPAYEOnlineEXBPriorXSDPath
  Else
    xsdPath = p11d32.SystemPAYEOnlineEXBXSDPath
  End If
  
  Text = TextFileLoad(xsdPath)
  search = " xmlns:exb="""
  p0 = InStr(1, Text, search, vbTextCompare)
  If (p0 > 0) Then
    p1 = InStr(p0 + Len(search), Text, """", vbTextCompare)
    p0 = p0 + Len(search)
    nspace = Mid$(Text, p0, p1 - p0)
    nspace = Trim(nspace)
  End If
  
  NameSpaceByYear = nspace
  
  'nspace = m_FullNamesSpace
  
  
  
    'Charles,
    'Thanks for sending in your 2013-14 P11D/P11D(b) recognition xml. This will be reviewed over the next day or two and a full response issued on completion.
    'With regards to the P46(Car) final quarter submission; you are using an incorrect namespace. For submissions up until the 5th of April 2014,
    'you should still be using http://www.govtalk.gov.uk/taxation/EXB/12-13/1. The namespace of http://www.govtalk.gov.uk/taxation/EXB/13-14/1 does not
    'take effect (in the "live" service)
    'until the 6th of April 2014 and will relate to 13-14 P11D/P11D(b), P9D returns and 2014-15 P46(Car) submissions.
    'Remember, one schema covers all of these messages and therefore needs to be the same...even though the P46(Car) will be for 14-15.
    'I hope this clarifies things?
    'Regards
    'Dave Hargreaves | Software Developers Support Team | IMS | HM Revenue & Customs | Tel: 03000 518308
    
    'see readme.txt in C:\projects\P11D\P11DXXXX\System\PAYEOnline
    
'email july 2017
'The current Schema and Business Validation Rules, published in the technical pack at
'https://www.gov.uk/government/publications/paye-internet-submissions-expenses-and-benefits-schema , are for use with the 2016-17 P11D/P11D(b)
'and 2017-18 P46(Car) returns.
'IT's always been this way for Expenses & Benefits. We have one over-arching schema which covers all the relevant message types
'but the P46(Car) is always one year ahead of the P11D. This is due to what is, essentially, an in-year form being tied to an end of year form.
'So, the 16-17 P11D is actually submitted in the 17-18 tax year…therefore, it follows that the associated P46(Car) is also for the 17-18 tax.
'Slightly confusing but again, it's been like this since the inception of the online Expenses & Benefits service.

  
    
    'for 2017/18 this year is 2 so was 2016/17, might change though in a new year
    'envelopeVersionCurrent = p11d32.AppYearShort & "-" & p11d32.AppYearShortNextYear & "/2"
    'envelopeVersionPrevious = p11d32.AppYearShortLastYear & "-" & p11d32.AppYearShort & "/2"
    
    
    'nspace = Replace$(nspace, envelopeVersionCurrent, envelopeVersionPrevious)
  
  
  'End If
  
  'NameSpaceByYear = m_FullNamesSpace
  
End Function



Private Function PAyeOnlineValidateXML(xml As String, xsdPath As String) As String
  Dim p0 As Long, p1 As Long
  Dim qs As QString
  Dim s As String
  
  Set qs = New QString
  Call qs.Append("<?xml version=""1.0"" encoding=""UTF-8"" ?>")
  xsdPath = Replace$(xsdPath, "\", "/")
  
  Call qs.Append("<IRenvelope xmlns=""" & Me.Namespace & """ xsi:schemaLocation=""" & Me.Namespace & " file:///" & xsdPath & """ xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">")
  
  p0 = InStr(1, xml, "<IRenvelope", vbBinaryCompare)
  p0 = InStr(p0, xml, ">", vbBinaryCompare)
  p1 = InStr(1, xml, "</IRenvelope", vbBinaryCompare)
  
  s = Mid$(xml, p0 + 1, p1 - (p0 + 1))
  Call qs.Append(s)
  Call qs.Append("</IRenvelope>")
    
  s = qs.bstr
  s = Replace(s, "xmlns=""http://www.govtalk.gov.uk/CM/envelope""", "", 1, -1, vbTextCompare)
  s = PrettyFormatXML(s)
  's = Replace(s, "1973-12-25", "1973-13-25", 1, -1, vbTextCompare)
  
  PAyeOnlineValidateXML = s
  
  
  
End Function


Private Sub ValidateXMLToRevenueLocal(ByVal xml As String, ByVal xsdPath As String)
  Dim ValidateXML As String
  Dim validator As PayeOnlineXMLValidator
  Dim validateForm As F_PayeOnlineValidate
On Error GoTo err_Err

  ValidateXML = PAyeOnlineValidateXML(xml, xsdPath)
  If (IsRunningInIDE()) Then
    Call TextFileSave("c:\payeonline_validate.xml", ValidateXML)
  End If
    
  Set validator = New PayeOnlineXMLValidator
  Call validator.Validate(ValidateXML, xsdPath)
  'Set ts = fso.CreateTextFile("c:\test2.xml", True)
  If (validator.Errors.Count > 0) Then
    Call ClearCursor
    Call F_PayeOnlineValidate.Init(validator, ValidateXML)
    Call SetCursor(vbHourglass)
    If (Not F_PayeOnlineValidate.ContinueSubmission) Then
      Call Err.Raise(ERR_PAYE_ONLINE_CANCEL, "ValidateXMLToRevenueLocal", "Submission cancelled")
    End If
    'validateForm
    'Call PayeOnlineXMLVlaidatorControl1.Init(validator, sPrettyFileText)
  End If
  
  'create a document that is valid XML
  
  'Call XMLErrorPayeOnline("Data sent to HMRC", qs.bstr)
   

err_End:
  Unload F_PayeOnlineValidate
  Exit Sub
err_Err:
  Call ClearAllCursors
  Unload F_PayeOnlineValidate
  If (Err.Number = ERR_PAYE_ONLINE_CANCEL) Then
    Call Err.Raise(ERR_INVALID, "ValidateXMLToRevenue", "Submission cancelled")
  End If
  
  If MsgBox("Error during local schema validation" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Do you still wish to try and submit to the Revenue?", vbYesNo, "Local schema validation") = vbNo Then
    Call Err.Raise(ERR_INVALID, "ValidateXMLToRevenue", "Submission cancelled")
  End If
  Resume err_End
  Resume
End Sub
Private Function AltovaXMLObject() As Object
  Dim altovaXML As Object
  
  On Error GoTo err_Err:

  Set altovaXML = CreateObject("AltovaXML.Application")

  Set AltovaXMLObject = altovaXML
err_Err:
  
  Exit Function
End Function
Private Function AltovalXMLErrorDescription(sXML As String)
  Dim altovaXML As Object
  
  Set altovaXML = AltovaXMLObject()
  If (altovaXML Is Nothing) Then
    AltovalXMLErrorDescription = "Install Altova XML Objects for further information or use XML Spy (better as gives line numbers) and remove all other nodes up to ExpensesAndBenefits from and validate against local schema."
    Exit Function
  End If

  altovaXML.XMLValidator.InputXMLFromText = sXML
  If (Not altovaXML.XMLValidator.isValid) Then
    AltovalXMLErrorDescription = altovaXML.XMLValidator.LastErrorMessage
  Else
    AltovalXMLErrorDescription = "Valid according to Altova"
  End If
  


End Function
Private Sub XMLErrorPayeOnline(ByVal sDOCName As String, sXML As String)
  Dim xml As MSXML2.DOMDocument60
  Dim vLines As Variant
  Dim sDescription As String
  Dim sLine As String
  Dim XMLError As MSXML2.IXMLDOMParseError
  Dim p0 As Long
  Dim iLen As Long
  Dim altovaXML As Object
    
  Set xml = New MSXML2.DOMDocument60
  xml.Async = False
  xml.validateOnParse = False
  xml.resolveExternals = True
  Call xml.loadXML(sXML)

  Set XMLError = xml.Validate
  If XMLError.ErrorCode <> 0 Then
    sDescription = "Error in XML document: '" & sDOCName & "'" & vbCrLf
    sDescription = sDescription & XMLError.reason & vbCrLf
    sDescription = sDescription & AltovalXMLErrorDescription(sXML) & vbCrLf
    
    If (XMLError.line > 0) Then
    
      sDescription = sDescription & " at line " & (XMLError.line) & ", position " & XMLError.linepos & " (see *)" & vbCrLf & vbCrLf
      p0 = XMLError.linepos - 30
      iLen = 30
      If p0 < 1 Then
        iLen = 30 + p0
        p0 = 1
      End If
      If (Len(sXML) > 0) Then
        vLines = Split(sXML, vbCrLf)
        sLine = vLines(XMLError.line - 1)
        sDescription = sDescription & Mid$(sLine, p0, iLen) & "*" & Mid$(sLine, XMLError.linepos, 30)
      End If
    End If
    
    Call Err.Raise(ERR_XMLNUMBERTOOBIG, "Submit_EfilerCom", sDescription)
  End If


    
End Sub

