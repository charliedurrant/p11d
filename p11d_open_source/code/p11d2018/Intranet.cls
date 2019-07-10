VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Intranet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Private Const S_WINDOWS_AUTHENTICATION As String = "L_AUTHENTICATION_WINDOWS"
Private Const S_XML_AUTHENTICATION As String = "L_AUTHENTICATION_XML"
Private Const S_AUTHENTICATION_OTHER As String = "L_AUTHENTICATION_OTHER"

Private Const S_TAG_EMPLOYER As String = "EMPLOYER"
Private Const S_TAG_EMPLOYEE As String = "EMPLOYEE"
Private Const S_TAG_XML As String = "<?xml version=""1.0""?>"

Public AuthenticationType As INTRANET_AUTHENTICATION_TYPE
Public OutputType As INTRANET_OUTPUT_TYPE
Public LoginUserNameSource As INTRANET_LOGIN_USERNAME_SOURCE
Public BannerTitle As String
Public UserInfoHTML As String
Public CaseSensitiveOnFullAuthentication As Boolean
Public BannerForeColor As Long
Public BannerBackColor As Long


Private m_OutputDirectory As String

Private m_TS As TextStream

Public Function UpdateListViewItem(li As IListItem, ben As IBenefitClass, Optional ByVal BenefitIndex As Long = 0) As Boolean
  
  On Error GoTo UpdateListViewItem_ERR

  Call xSet("UpdateListViewItem")

  Dim emp As Employer

  If li Is Nothing Then Call Err.Raise(ERR_ERROR, "UpdateListViewItem", "The list item is nothing")
  If ben Is Nothing Then Call Err.Raise(ERR_ERROR, "UpdateListViewItem", "The beneift is nothing")

  If BenefitIndex > 0 Then li.Tag = BenefitIndex

  Set emp = ben
  
  If Not emp.IntranetValid(False) Then
    li.SmallIcon = IMG_EXCLAMATION
  Else
    li.SmallIcon = IMG_OK
  End If

  li.Text = ben.Name
  li.SubItems(1) = ben.value(employer_EmployeesCount)
  
  UpdateListViewItem = True

UpdateListViewItem_END:
  Call xReturn("UpdateListViewItem")
  Exit Function

UpdateListViewItem_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateListViewItem", "Update List View Item", "Error Updating List Item")
  Resume UpdateListViewItem_END
  Resume
End Function
Private Function XMLHeader()
  XMLHeader = S_TAG_XML & vbCrLf
End Function

Private Function XMLAttrib(ByVal sAttribname As String, ByVal sAttribValue As String)
  XMLAttrib = " " & sAttribname & "=""" & sAttribValue & """"
End Function
Private Function XMLEmployerOpen(Employer As IBenefitClass, sroot As String) As String
  
  XMLEmployerOpen = XMLTagOpen(S_TAG_EMPLOYER, "name", Employer.value(ITEM_DESC), "virtualroot", sroot)
  
End Function
Private Function XMLTagOpen(ByVal sTagName As String, ParamArray vAttributes())
  Dim sAttribs As String
  Dim s As String
  Dim v As Variant
  Dim i As Long

  s = "<" & sTagName
  v = vAttributes()
  If IsArrayEx(v) Then
    For i = LBound(vAttributes) To UBound(vAttributes) Step 2
      s = s & XMLAttrib(v(i), v(i + 1))
    Next
  End If

  XMLTagOpen = s & ">" & vbCrLf

End Function
Private Function XMLTagClose(ByVal sTagName As String)
  XMLTagClose = "</" & sTagName & ">" & vbCrLf
End Function
Private Function XMLEmployee(Employee As IBenefitClass, ByVal sReportName As String, ByVal rep As Reporter) As String
  Dim sUsername As String
  Dim sFileName As String
  
  sFileName = p11d32.ReportPrint.ExportFileName(sReportName, Employee, rep, False)
  Select Case Me.LoginUserNameSource
    Case 0
      sUsername = Employee.value(ee_Username_db)
    Case 1
      sUsername = Employee.value(ee_PersonnelNumber_db)
    Case 2
      sUsername = Employee.value(ee_Email_db)
    Case 3
      sUsername = Employee.value(ee_FullName)
  End Select
  If Not p11d32.Intranet.CaseSensitiveOnFullAuthentication Then
    sUsername = UCASE$(sUsername)
  End If
  XMLEmployee = XMLTagOpen(S_TAG_EMPLOYEE, "name", Employee.value(ee_FullName), "personnelnumber", Employee.value(ee_PersonnelNumber_db), "filename", sFileName, "username", sUsername, "password", Employee.value(ee_Password_db)) & XMLTagClose(S_TAG_EMPLOYEE)
End Function
Private Function ClearFolder() As Boolean
  Dim fs As FileSystemObject
  Dim folderintranet As folder, folder As folder, File As File
  
  Dim s As String
  
  On Error GoTo ClearFolder_ERR
  
  
  s = Me.OutputDirectory
  If InStr(1, "/\", Right(s, 1), vbBinaryCompare) > 0 Then s = Left$(s, Len(s) - 1)
  Set fs = New FileSystemObject
  Set folderintranet = fs.GetFolder(s)
    
  If folderintranet.Files.Count = 0 And folderintranet.SubFolders.Count = 0 Then
    ClearFolder = True
    GoTo ClearFolder_END
  End If
  
  
  If MsgBox("The following process will remove the contents of:" & vbCrLf & vbCrLf & Me.OutputDirectory & vbCrLf & vbCrLf & "Do you wish to continue?", vbYesNo) = vbNo Then
    GoTo ClearFolder_END
  End If
  
  
  For Each folder In folderintranet.SubFolders
    Call fs.DeleteFolder(folder.Path, True)
  Next
  For Each File In folderintranet.Files
    Call fs.DeleteFile(File.Path, True)
  Next
  
  ClearFolder = True
  
ClearFolder_END:
  Exit Function
ClearFolder_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "ClearFolder"), Err.Description)
  Resume
End Function
Private Function PrintHTMFiles(ByVal sExportDirectory As String, ByVal employees As ObjectList) As String
  Dim CurrentExportAllowUserNameHTML As Boolean
  Dim CurrentPrintedEmployees As Boolean
  Dim CurrentExportDirectory As String
  Dim CurrentDestination As REPORT_DESTINATION
  
  On Error GoTo PrintHTMFiles_ERR
  
  Set p11d32.ReportPrint.SelectedEmployees = employees
  CurrentExportAllowUserNameHTML = p11d32.ReportPrint.ExportAllowUserNameHTML
  CurrentPrintedEmployees = p11d32.ReportPrint.PrintedEmployees
  CurrentExportDirectory = p11d32.ReportPrint.ExportDirectory
  CurrentDestination = p11d32.ReportPrint.Destination
  
  p11d32.ReportPrint.ExportAllowUserNameHTML = False
  p11d32.ReportPrint.PrintedEmployees = False
  p11d32.ReportPrint.ExportDirectory = FullPath(sExportDirectory)
  p11d32.ReportPrint.Destination = REPD_FILE_HTML
  'RC TTP#461 - Make independent of print dialog
  
  If p11d32.Intranet.OutputType = INTRANET_OUTPUT_TYPE.IOT_P11D Then
    Call p11d32.ReportPrint.DoStandardReport(RPT_HMIT, PRINT_REPORT)
    PrintHTMFiles = p11d32.ReportPrint.Name(RPT_HMIT)
  ElseIf p11d32.Intranet.OutputType = IOT_P11D_WORKING_PAPERS Then
    Call p11d32.ReportPrint.DoStandardReport(RPT_HMIT_PLUS_WORKING_PAPERS, PRINT_REPORT)
    PrintHTMFiles = p11d32.ReportPrint.Name(RPT_HMIT_PLUS_WORKING_PAPERS)
  ElseIf p11d32.Intranet.OutputType = IOT_EMPLOYEE_LETTER Then
    Call p11d32.ReportPrint.DoStandardReport(RPT_EMPLOYEE_LETTER, PRINT_REPORT)
    PrintHTMFiles = p11d32.ReportPrint.Name(RPT_EMPLOYEE_LETTER)
  ElseIf p11d32.Intranet.OutputType = IOT_WORKING_PAPERS Then
    Call p11d32.ReportPrint.DoStandardReport(RPT_WORKINGPAPERS, PRINT_REPORT)
    PrintHTMFiles = p11d32.ReportPrint.Name(RPT_WORKINGPAPERS)
  Else
   Call ECASE("Invalid report")
  End If
  
  
  
  
PrintHTMFiles_END:
  Set p11d32.ReportPrint.SelectedEmployees = Nothing
  p11d32.ReportPrint.ExportAllowUserNameHTML = CurrentExportAllowUserNameHTML
  p11d32.ReportPrint.PrintedEmployees = CurrentPrintedEmployees
  p11d32.ReportPrint.ExportDirectory = CurrentExportDirectory
  p11d32.ReportPrint.Destination = CurrentDestination
  
  Exit Function
PrintHTMFiles_ERR:
  If Not p11d32 Is Nothing Then
    If Not p11d32.ReportPrint Is Nothing Then
      p11d32.ReportPrint.ExportNoSpaces = False
    End If
  End If
  Call Err.Raise(Err.Number, ErrorSource(Err, "PrintHTMFiles"), Err.Description)
  Resume
End Function
Private Sub InvalidReportError()
  Dim s As String
  
  On Error GoTo InvalidReportError_ERR
   
  s = "Invalid report selection," & vbCrLf & vbCrLf
  s = s & "the only valid reports available are:" & vbCrLf & vbCrLf
  s = s & p11d32.ReportPrint.Name(RPT_HMIT) & vbCrLf
  s = s & p11d32.ReportPrint.Name(RPT_HMIT_PLUS_WORKING_PAPERS) & vbCrLf
  s = s & vbCrLf & "Please go to the print dialogue to amend"
  Call Err.Raise(ERR_REPORT_INVALID, "InvalidReportError", s)
  
InvalidReportError_END:
  Exit Sub
InvalidReportError_ERR:
  Call Err.Raise(ERR_REPORT_INVALID, "InvalidReportError", s)
  Resume
End Sub
Private Sub CheckReportSelection()
  Dim def As Long
  Dim rt As RPT_TYPE
  
  On Error GoTo CheckReportSelection_ERR
  
  
  rt = p11d32.ReportPrint.ReportType(p11d32.ReportPrint.DefaultReportIndex)
  If rt < 0 Then Call Err.Raise(ERR_REPORT_INVALID, ErrorSource(Err, "ReportPrintPrapare"), Err.Description)
  
  If rt <> RPTT_STANDARD Then
    Call InvalidReportError
  Else
    def = p11d32.ReportPrint.DefaultReportIndex
    If def <> RPT_HMIT And def <> RPT_HMIT_PLUS_WORKING_PAPERS Then
      Call InvalidReportError
    End If
  End If
  
CheckReportSelection_END:
  Exit Sub
CheckReportSelection_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CheckReportSelection"), Err.Description)
  Resume
End Sub

Public Sub XMLFile(ByVal prgEmployee As TCSProgressBar, ByVal lblFile As Label)
  Dim rt As RPT_TYPE
  Dim i As Long, j As Long
  Dim ey As Employer, ee As Employee, employees As ObjectList
  Dim ben As IBenefitClass, benEmployee As IBenefitClass
  Dim Employers As ObjectList
  Dim sReportName As String
  Dim sroot As String
  Dim rep As Reporter
  Dim f As folder
  Dim fs As FileSystemObject
  Dim sEmployerPath As String, s As String
  Dim ts As TextStream
  Dim sFileText As String
  
  On Error GoTo XMLFile_ERR

  Call xSet("XMLFile")
  'RC TTP#461 - Make independent of print dialog
  'Call CheckReportSelection
  If Not ClearFolder Then GoTo XMLFile_END
  Set rep = ReporterNew()
  prgEmployee.Visible = True
  prgEmployee.value = prgEmployee.Min
  prgEmployee.Indicator = None
  prgEmployee.Caption = ""
  
  p11d32.ReportPrint.ExportNoSpaces = True
  Call OpenFile(lblFile)
  
  Set fs = New FileSystemObject
  
  fs.CopyFile p11d32.SystemIntranetPath & "*.asp", OutputDirectory
  
  s = p11d32.SystemIntranetPath & "P11D.asp"
  Set ts = fs.OpenTextFile(s, ForReading)
  sFileText = ts.ReadAll
  sFileText = Replace(sFileText, "{BANNER_TITLE_HTML}", Me.BannerTitle, , vbTextCompare)
  sFileText = Replace(sFileText, "{USER_INFORMATION_HTML}", Me.UserInfoHTML, , vbTextCompare)
  
  sFileText = Replace(sFileText, "{BANNER_FORE_COLOR}", HTMLColor(Me.BannerForeColor), , vbTextCompare)
  sFileText = Replace(sFileText, "{BANNER_BACK_COLOR}", HTMLColor(Me.BannerBackColor), , vbTextCompare)
  s = OutputDirectory & "P11D.asp"
  If FileExists(s) Then
    Call RemoveReadOnlyFile(s)
  End If
  Set ts = fs.OpenTextFile(s, ForWriting, True)
  Call ts.Write(sFileText)
  Call ts.Close
  
  
  fs.CopyFile p11d32.SystemIntranetPath & "*.inc", OutputDirectory
  fs.CopyFile p11d32.SystemIntranetPath & "*.asa", OutputDirectory
  fs.CopyFolder p11d32.SystemIntranetPath & "images", OutputDirectory
  
  
  
  s = OutputDirectory & "p11d.inc"
  Set ts = fs.OpenTextFile(s, ForReading)
  sFileText = ts.ReadAll
  ts.Close
  
  'remove the read only attribute
  Call RemoveReadOnlyFile(s)
  
  Set ts = fs.OpenTextFile(s, ForWriting)
  sFileText = Replace(sFileText, "{XML_FILE_NAME_NO_PATH}", AppExeName & ".xml", , , vbTextCompare)
  If p11d32.Intranet.AuthenticationType = IAT_WINDOWS Then
    sFileText = Replace(sFileText, "{AUTHENTICATION_TYPE}", S_WINDOWS_AUTHENTICATION, , , vbTextCompare)
  ElseIf p11d32.Intranet.AuthenticationType = IAT_OTHER Then
    sFileText = Replace(sFileText, "{AUTHENTICATION_TYPE}", S_AUTHENTICATION_OTHER, , , vbTextCompare)
  Else
    sFileText = Replace(sFileText, "{AUTHENTICATION_TYPE}", S_XML_AUTHENTICATION, , , vbTextCompare)
  End If
  
  If (Not p11d32.Intranet.CaseSensitiveOnFullAuthentication) And (p11d32.Intranet.AuthenticationType = IAT_FULL) Then
    sFileText = Replace(sFileText, "{CASE_SENSITIVE}", "false", , , vbTextCompare)
  Else
    sFileText = Replace(sFileText, "{CASE_SENSITIVE}", "true", , , vbTextCompare)
  End If
  
  Call ts.Write(sFileText)
  Call ts.Close
  
  Call m_TS.Write(XMLHeader)
  Call EmployersSelected(Employers)
  Call m_TS.Write(XMLTagOpen("EMPLOYERS"))
         
  For i = 1 To Employers.Count
    Set ey = Employers(i)
    Call p11d32.LoadEmployer(ey, False)
    Set ben = ey
    If ben.value(employer_EmployeesCount) = 0 Then GoTo NEXT_EMPLOYER
    sroot = GetEmployerFilename(ben.value(ITEM_DESC))
    
    Call m_TS.Write(XMLEmployerOpen(ey, sroot))
    
    sEmployerPath = Me.OutputDirectory & sroot
    Call fs.CreateFolder(sEmployerPath)
    
    Set employees = p11d32.CurrentEmployer.employees
    sReportName = PrintHTMFiles(sEmployerPath, employees)
    
    prgEmployee.Max = ey.employees.Count
    prgEmployee.value = 0
    prgEmployee.Indicator = ValueOfMax
    
    For j = 1 To ey.employees.Count
      Set ee = ey.employees(j)
      Set benEmployee = ee
      Call prgEmployee.StepCaption("Analysing employee: " & ee.FullName)
      If ee Is Nothing Then GoTo NEXT_EMPLOYEE
      Call m_TS.Write(XMLEmployee(ee, sReportName, rep))
NEXT_EMPLOYEE:
    Next
    
NEXT_EMPLOYER:
    Call m_TS.Write(XMLTagClose(S_TAG_EMPLOYER))
    Call p11d32.LoadEmployer(Nothing, False)
  Next
  
  Call m_TS.Write(XMLTagClose("EMPLOYERS"))
  Call CloseFile
  prgEmployee.Caption = "Completed creating XML file."
  prgEmployee.Indicator = None

XMLFile_END:
  p11d32.ReportPrint.ExportNoSpaces = False
  Call xReturn("XMLFile")
  Exit Sub
XMLFile_ERR:
  Call CloseFile
  Call ErrorMessage(ERR_ERROR, Err, "XMLFile", "XMLFile", "Unable to create XML file.")
  Resume XMLFile_END
  Resume
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
      Set li = F_Intranet.lvIntranetEmployers.listitems.Add()
      Call UpdateListViewItem(li, ben, i)
    End If
  Next

EmployersToListView_END:
  Call xReturn("EmployersToListView")
  Exit Sub
EmployersToListView_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployersToListView", "Employers To List View", "Error placing the current employers to the Intranet list view.")
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
      If ben.value(employer_IntranetSelected) Then
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

Private Function OpenFile(ByVal lblFile As Label) As Boolean
  Dim fs As FileSystemObject
  Dim sFileName As String
  
  On Error GoTo OpenFile_ERR

  Set fs = New FileSystemObject
  sFileName = PathAndFile()
  Set m_TS = fs.CreateTextFile(sFileName, True)
'
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

Private Sub CloseFile()
  If Not m_TS Is Nothing Then
    m_TS.Close
    Set m_TS = Nothing
  End If
End Sub


Private Property Get PathAndFile() As String
  Dim File As String
  ' used AppExeName since AppName included a "/" which was causing an error message 'so
  File = Replace(AppExeName, " ", "") & ".xml"
  PathAndFile = OutputDirectory & File
End Property
Public Sub Start()
 
  On Error GoTo Start_ERR
  Call xSet("Start")

  If Not FileExists(OutputDirectory, True) Then
    Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "F_MM", "The Intranet directory does not exist = " & p11d32.Intranet.OutputDirectory & ", unable to create export.")
  End If
  F_Intranet.fb.Directory = FullPath(OutputDirectory)
  
  Call EmployersToListView

  F_Intranet.Caption = "Intranet " & p11d32.AppYear
  Call p11d32.Help.ShowForm(F_Intranet, vbModal)
  


Start_END:
  Set p11d32.CurrentEmployer = Nothing
  Set F_Intranet = Nothing
  Call xReturn("Start")
  Exit Sub
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error starting Intranet.")
  Resume Start_END
End Sub

'cad this shows lack of understanding
Private Function GetEmployerFilename(eyName As String) As String
  Dim s As String, v As String, d As String
  Dim P As String, l As String, Y As String
  Dim k As String, r As String
  
  s = Replace(eyName, " ", "")
  v = Replace(s, "\", "")
  d = Replace(v, ":", "")
  P = Replace(d, ",", "")
  l = Replace(P, ".", "")
  Y = Replace(l, ";", "")
  k = Replace(Y, "'", "")
  r = Replace(k, "/", "")
  
  GetEmployerFilename = r
  
End Function

Public Property Get OutputDirectory() As String
  OutputDirectory = m_OutputDirectory
End Property

Public Property Let OutputDirectory(sNewValue As String)
  m_OutputDirectory = sNewValue
End Property


