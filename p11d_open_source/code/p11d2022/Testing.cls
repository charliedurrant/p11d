VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Testing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IErrorFilter

Public LastGoogleTestSheetId As String
Public GoogleAccessToken As String
Public GoogleAccessTokenExpiresAt As Date

Public YearStart As Long
Public ImportSheets As ObjectList
Private m_specs As Dictionary
Private m_testingTests As ObjectList
Private m_gsheet As GoogleSheet
Private m_Errors As StringList
Private m_importSheet As String
Private Sub Class_Initialize()
  GoogleAccessTokenExpiresAt = UNDATED
  GoogleAccessToken = ""
End Sub
Public Property Get GoggleAccessTokenExpired() As Boolean
  GoggleAccessTokenExpired = True
  If GoogleAccessTokenExpiresAt = UNDATED Then
    Exit Property
  End If
  If (Now() >= GoogleAccessToken) Then
    Exit Property
  End If
  GoggleAccessTokenExpired = False
End Property
Private Function IErrorFilter_FilterErrorMessage(ByVal Username As String, ByVal DateTime As Date, ByVal ErrorNumber As Long, ByVal ErrorName As String, ByVal ErrorText As String, ByVal SourceFunction As String) As Boolean
  Call m_Errors.Add(ErrorText)
End Function

Private Function IErrorFilter_PrintAvailable() As Boolean

End Function

Private Function IErrorFilter_PrintErrors() As Boolean

End Function

Private Function IErrorFilter_ViewAvailable() As Boolean

End Function

Private Function IErrorFilter_ViewErrors() As Boolean

End Function

Public Sub Init(ByVal gSheet As GoogleSheet)
  Dim rs As ADODB.Recordset
  Dim testingTest As testingTest
  Dim row As Long
  Dim value As String
  Dim Item As String
  Dim csv As csvParser
  
On Error GoTo Err_Err
  Set m_Errors = New StringList
  Set m_gsheet = gSheet
  Set m_testingTests = New ObjectList
  Set m_specs = ReadSpecs(m_gsheet)
  
  Set csv = gSheet.SheetAsParsedCsv("settings")
  
  For row = 0 To csv.RowCount - 1
    value = Trim$(csv.value(row, "value"))
    Item = UCASE$(Trim$(csv.value(row, "item")))
    
    Select Case Item
     Case "YEAR_START"
        Me.YearStart = CLng(value)
        If Me.YearStart <> p11d32.AppYear Then
         Call Err.Raise(ERR_TESTING, "Testing.Init", "YEAR_START in the google sheet should be " & p11d32.AppYear)
        End If
     Case "IMPORT_SHEET"
       
       Set testingTest = New testingTest
       testingTest.Sheet = value
       
       If Not m_specs.Exists(testingTest.Sheet) Then
         Call Err.Raise(ERR_TESTING, "Init", "No spec exist fot the import sheet " & testingTest.Sheet)
       End If
       
       Set testingTest.Spec = m_specs(testingTest.Sheet)
       Call m_testingTests.Add(testingTest)
    End Select
    
  Next
  
Err_End:
  Exit Sub
Err_Err:
  Call Err.Raise(ERR_ERROR, "Testing.Init", Err.Description)
  Resume
End Sub
Private Function ReadSpecs(ByVal gSheet As GoogleSheet) As Dictionary
  Dim specs As Dictionary
  Dim Spec As TestingSpec
  Dim csv As csvParser
  Dim row As Long
  
  Set specs = New Dictionary
  Set csv = gSheet.SheetAsParsedCsv("specs")
  
  For row = 0 To csv.RowCount - 1
    Set Spec = New TestingSpec
    Spec.Sheet = csv.value(row, "sheet")
    Spec.ImportSpec = MarkupGoogleSheetData(csv.value(row, "import_spec"))
    Spec.ReportSpec = MarkupGoogleSheetData(csv.value(row, "report_spec"))
    Spec.ReportResultsSheet = MarkupGoogleSheetData(csv.value(row, "report_results_sheet"))
    Call specs.Add(Spec.Sheet, Spec)
  Next
  
  Set ReadSpecs = specs

End Function
Private Function MarkupGoogleSheetData(ByRef data As String) As String
  Dim ret As String
  ret = Replace(data, Chr(13), "")
  ret = Replace(ret, Chr(10), Chr(13) & Chr(10))
  MarkupGoogleSheetData = ret
End Function

Private Function CreateEmployer() As Employer
  Dim empr As Employer
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  Dim employerIndex As Long
  Call p11d32.EmployerScreen
  
  Set empr = New Employer
  
  Call empr.CreateEmployer(FullPath(p11d32.WorkingDirectory) & GuidNewEmployer & p11d32.AppYearShort & p11d32.Rates.FileExtensionCurrent)
  Call empr.OpenEmployer(False)
  Set ben = empr
  ben.value(employer_Name_db) = "Testing:" & Format$(Now, "MM/dd/yyyy hh::mm:ss")
  ben.value(employer_Payeref_db) = "123/E1234"
  
  Call ben.WriteDB
  
  employerIndex = p11d32.Employers.Add(empr)
  Set ibf = CurrentForm
  Call ibf.BenefitToListView(ben, employerIndex)
  Call SelectEmployer(employerIndex)
  Set CreateEmployer = empr
End Function

Private Sub SelectEmployer(employerIndex As Long)
  Dim ibf As IBenefitForm2
  Set ibf = CurrentForm
  
  Call SelectBenefitByBenefitIndex(ibf, employerIndex)
End Sub
Private Function Title(titleStr As String) As String
  Dim line As String
  line = "----------------------------------------------------------------------------------------------------------------------------"
  
  Title = vbCrLf & line & vbCrLf & titleStr & vbCrLf & line & vbCrLf
  
End Function
Public Function Run(ByRef results As String) As Boolean
  Dim sb As QString
  Dim oldEcmployeeeScope As EMPLOYEE_SELECTION
  Dim empr As Employer
  Dim i As Long
  Dim testingTest As testingTest
  Dim importErrors As Boolean
  
  Set sb = New QString
  Set empr = Nothing
  Run = True
On Error GoTo Err_Err
    
  oldEcmployeeeScope = p11d32.ReportPrint.EmployeeSelection
  p11d32.ReportPrint.EmployeeSelection = ES_ALL
  
  Set empr = CreateEmployer()
  sb.Append (Title("Test run: " & Format(Now, "dd/MM/yyyy hh:mm:ss")))
  
  'import the data
  If ImportData(empr, sb) Then 'errors
    GoTo Err_End
  End If
  
  Call p11d32.LoadEmployer(empr, True, True, False)
  
  Call sb.Append(Title("Export reports and upload to the google sheet"))
  'run reports and upload
  For i = 1 To m_testingTests.Count
    Set testingTest = m_testingTests(i)
    If testingTest.Spec.ReportEnabled Then
        
        Dim reportPathAndFile As String
        Dim reportDataPathAndFile As String
        
        reportPathAndFile = GetTempFilename
        reportDataPathAndFile = GetTempFilename
        Call TextFileSave(reportPathAndFile, testingTest.Spec.ReportSpec)
        
        Call p11d32.ReportPrint.StartReportWizard(reportPathAndFile, REPORTW_PREPARE_EXPORT, reportDataPathAndFile)
        
        Dim csvData As String
        
        csvData = TextFileLoad(reportDataPathAndFile)
        
        Call xKill(reportPathAndFile)
        Call xKill(reportDataPathAndFile)
        
        Call m_gsheet.UploadCsvToSheet(testingTest.Spec.ReportResultsSheet, csvData)
        Call sb.Append("Test(" & testingTest.Sheet & "): report uploaded to " & testingTest.Spec.ReportResultsSheet & vbCrLf)
    End If
  Next
  
  Dim resultsCsv As csvParser
  Dim resultsCsvStr As String
  
  Dim failures As Long
  Dim failuresStr As String
  
  Set resultsCsv = m_gsheet.SheetAsParsedCsv2("results", resultsCsvStr)
  sb.Append (Title("Results"))
  sb.Append (resultsCsvStr)
  
  Dim anyFailedTests As Boolean
  For i = 0 To resultsCsv.RowCount - 1
    failuresStr = Trim$(resultsCsv.value(i, "failures"))
    If Len(failuresStr) > 0 Then
      failures = CLng(failuresStr)
      If failures > 0 Then
        anyFailedTests = True
      End If
    End If
  Next
  
  If Not anyFailedTests Then
    Run = False
  End If
  
Err_End:
  p11d32.ReportPrint.EmployeeSelection = oldEcmployeeeScope
  Call p11d32.StartScreen
  
  Dim emprList As Employer
  If (Not empr Is Nothing) Then
    For i = 1 To p11d32.Employers.Count
      Set emprList = p11d32.Employers.Item(i)
      If (emprList.Name = empr.Name) Then
        Call SelectEmployer(i)
        Exit For
      End If
    Next
  End If
  
  results = LTrim$(sb.bstr)
  Exit Function
Err_Err:
  
  Call sb.Append(Err.Description)
  Call ErrorMessage(ERR_TESTING, Err, "Run", "Testing", Err.Description)
  Resume Err_End
  Resume
End Function
Private Function ImportData(empr As Employer, sb As QString) As Boolean
  Dim i As Long, j As Long
  Dim testingTest As testingTest
  Dim importFilePath As String
  Dim specFilePath As String
  
  sb.Append (Title("Imports"))
  
  ImportData = False
  For i = 1 To m_testingTests.Count
    Set testingTest = m_testingTests(i)
    importFilePath = GetTempFilename
    Call m_gsheet.SheetAsCsvToFile(testingTest.Sheet, importFilePath)
    specFilePath = GetTempFilename
    Call TextFileSave(specFilePath, testingTest.Spec.ImportSpec)
    Call ImportFile(empr, importFilePath, specFilePath, testingTest.ImportQueryId, testingTest.Sheet)
    sb.Append (testingTest.Sheet & " errors: " & CStr(m_Errors.Count) & vbCrLf)
    If m_Errors.Count > 0 Then
      ImportData = True
      Call sb.Append(vbCrLf)
      For j = 1 To m_Errors.Count
        sb.Append (m_Errors.Item(j) & vbCrLf)
      Next
      sb.Append (vbCrLf)
    End If
    Call m_Errors.RemoveAll
  Next
  
End Function
Private Sub ImportFile(ByVal empr As Employer, ByVal dataPathAndFile As String, ByVal specPathAndFile As String, importType As QUERY_NAMES_IMP, Sheet As String)
  Dim sql As SQLQUERIES_IMP
  Dim rs As Recordset
  Dim ic As ImportClass
  Dim ErrorCount As Long
  
On Error GoTo Err_Err:

  Set ic = New ImportClass
  Set ic.Workspace = p11d32.P11DWS
  Set ic.ErrorFilter = Me
  
  Set ATC2CORE.ErrorFilter = Me
  
  m_importSheet = Sheet
  ic.RightTrimFields = True
  
  Call ic.InitImport(dataPathAndFile, IMPORT_DELIMITED, specPathAndFile)
    
  Set sql = New SQLQUERIES_IMP
  Set rs = empr.db.OpenRecordset(sql.Queries(importType))
  'cad import
  
  ErrorCount = ic.ImportFile(rs)
  Call rs.Close
  Set rs = Nothing
  If (ErrorCount < 0) Then
    Call m_Errors.Add("Fatal error, no lines imported")
  ElseIf ErrorCount = 0 Then
    Dim linesOk As Long
    Dim sqli As New SQLQUERIES_IMP
    Dim importName As String
    importName = sqli.Name(importType)
    Set p11d32.Importing.DbP11dFile = empr.db
    Call p11d32.Importing.PostProcessEx(ErrorCount, linesOk, importName, importType)
  End If
    
Err_End:
  Set p11d32.Importing.DbP11dFile = Nothing
  Set ATC2CORE.ErrorFilter = Nothing
  Set ic = Nothing
  Exit Sub
Err_Err:
  Set p11d32.Importing.DbP11dFile = Nothing
  Set ATC2CORE.ErrorFilter = Nothing
  Call m_Errors.Add(Err.Description)
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub
Private Sub RunImport(dataPathAndFile As String, specPathAndFile As String)
  
End Sub
Private Function RunImports() As String
    
End Function


