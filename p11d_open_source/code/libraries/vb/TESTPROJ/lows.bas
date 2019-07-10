Attribute VB_Name = "lows"
Option Explicit

Public Sub TestByVal(ByVal s As String)
  Dim s2 As String
  
  s2 = Mid$(s, 5, 1)
End Sub

Public Sub TestByRef(s As String)
  Dim s2 As String
  
  s2 = Mid$(s, 5, 1)
End Sub


Public Sub TestError()
  Dim s As String
  On Error GoTo TestError_err
    
  Err.Raise ERR_TESTERROR, "TestError", "This is the correct error"
TestError_end:
  Exit Sub
  
TestError_err:
  Call ErrorMessagePush(Err)
  Call TestErrorSub
  Call ErrorMessagePop(ERR_ERROR, Err, "", "Test Error handling", "Test Error handling" & vbCrLf)
  Resume TestError_end
End Sub

Private Sub TestErrorSub()
  On Error GoTo TestErrorSub_err
  
  Err.Raise ERR_TESTERROR, "TestErrorSub", "Wrong error!!"
  
TestErrorSub_end:
  Exit Sub
  
TestErrorSub_err:
  Call ErrorMessagePush(Err)
  Call TestErrorSubSub
  Call ErrorMessagePop(ERR_ERROR, Err, "TestErrorSub", "Test Error handling", "Test Error handling" & vbCrLf)
  Resume TestErrorSub_end
End Sub

Private Sub TestErrorSubSub()
  On Error GoTo TestErrorSubSub_err
  
  Err.Raise ERR_TESTERROR, "TestErrorSubSub", "first Wrong error!!"
  
TestErrorSubSub_end:
  Exit Sub
  
TestErrorSubSub_err:
  Call ErrorMessage(ERR_ERROR, Err, "TestErrorSubSub", "Test Error handling", "Test Error handling" & vbCrLf)
  Resume TestErrorSubSub_end
End Sub

Public Sub SQLTestDebug()
  Dim sqlTest As SQLDebug
 
  Set sqlTest = New SQLDebug
  Call sqlTest.Show(AppPath & "\TEST.MDB", False)
End Sub
Public Function ADOOracleConnectString(ByVal DataSource As String, Optional ByVal UserID As String, Optional ByVal Password As String) As String
  ADOOracleConnectString = "Provider=MSDAORA.1;Persist Security Info=True;Password=" & Password & ";User ID=" & UserID & ";Data Source=" & DataSource
End Function
Public Sub SQLADOTestDebug(ByVal TARGET As DATABASE_TARGET)
  Dim sqlTest As SQLDebugADO
  Dim s As String
  
  Set sqlTest = New SQLDebugADO
  
  Select Case TARGET
    Case DATABASE_TARGET.DB_TARGET_JET
      s = ADOAccess4ConnectString(AppPath & "\TEST.MDB")
    Case DATABASE_TARGET.DB_TARGET_ORACLE
      s = ADOOracleConnectString("idb", "sa_idb", "ukcentral")
    Case DATABASE_TARGET.DB_TARGET_SQLSERVER
      s = ADOSQLConnectString("londs3103", "sa", "ukcentral", "WebStatistics")
  End Select
  
  Call sqlTest.Show(s, False, TARGET)
End Sub


Public Sub ImportTest()
  Dim db As Database, rs As Recordset
  Dim ic As ImportClass, nErr As Long
  Dim sFile As String, sImportDir As String
  
  On Error GoTo ImportTest_err
  sImportDir = FullPath(AppPath)
  Set db = InitDB(gwsMain, sImportDir & "test.mdb", "Test Database Name")
  Set rs = db.OpenRecordset("SELECT * from Contacts")
  Set ic = New ImportClass
  sFile = FileOpenDlg("Choose file to import", "Comma Separated Values (*.csv)|*.CSV|Text Files (*.txt)|*.txt|All Files|*.*", sImportDir)
  If ic.InitImport(sFile, IMPORT_DELIMITED, sImportDir & "TEST.CFG") Then
    'ic.HeaderCount = 1
    nErr = ic.ImportFile(rs, frmMain.Status1.prg)
    Call frmMain.Status1.ClearCaptions
    Set rs = Nothing
  End If
  Exit Sub
  
ImportTest_err:
  Call ErrorMessage(ERR_ERROR, Err, "ImportTest", "Import Test", "Test Import" & vbCrLf)
End Sub

Public Sub ImportWizardTest()
  Dim db As Database, rs As Recordset
  Dim ic As ImportClass, icw As ImportWizard, nErr As Long
  Dim sFile As String, sImportDir As String
  
  On Error GoTo ImportTest_err
  sImportDir = FullPath(AppPath)
  Set db = InitDB(gwsMain, sImportDir & "test.mdb", "Test Database Name")
  Set rs = db.OpenRecordset("SELECT * from Contacts")
  Set ic = New ImportClass
  Set icw = ic.ImportWizard
  Call icw.AddRS(rs, "ImportContacts", , , NO_UPDATES, , "Import Contacts")
  Call icw.StartWizard
  Call ic.KillImporter
  Exit Sub
  
ImportTest_err:
  Call ErrorMessage(ERR_ERROR, Err, "ImportTest", "Import Test", "Test Import" & vbCrLf)
End Sub

