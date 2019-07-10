VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBONewFixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbUpgrade As Boolean
Private mbCancel As Boolean
Private mlFixLevel As Long
Private mlPostFixLevel As Long
Private mlDbVersion As Long
Private mlDBSubVersion As Long
Private mlTemplateFixLevel As Long
Private mlTemplateDbVersion As Long
Private mlTemplatePostFixLevel As Long
Private mlTemplateDbSubVersion As Long
Private dbTemplate As Database

Private Sub Class_Terminate()
  Call Kill
End Sub

Public Function NewTemplateFile(ByVal SelectedDB As Database, strFileName As String, iNotify As IBaseNotify) As Database
  
  On Error GoTo NewTemplateFile_Err
  
  If Not SelectedDB Is Nothing Then
    If Not iNotify Is Nothing Then Set gNotify = iNotify
    Set NewTemplateFile = NewTemplate(SelectedDB, strFileName, iNotify)
  End If
  
NewTemplateFile_End:
  Exit Function

NewTemplateFile_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NewTemplateFile", "Error checking version", "Each time an Abacus+ file is opened, the file is compared against the template file to ensure that it is fully up to date. The two files will then be synchronised. An error has occurred whilst running this process.")
  Resume NewTemplateFile_End
  Resume
End Function

Private Function NewTemplate(ByVal SelectedDB As Database, strFileName As String, iNotify As IBaseNotify) As Database
  Dim NewDB As Database, CopySelectedDB As Database
  Dim strNewFileName As String, strCopyFileName As String, strTemplateFile As String
  Dim strTempFile As String, strExt As String, strDir As String
  Dim fs As FileSystemObject
  Dim drvFile As Drive
  Dim drvLocal As Drive
  Dim bUpgrade As Boolean
  
  On Error GoTo NewTemplate_Err
  
  If Not FileExists(strFileName) Then GoTo NewTemplate_End
  mbUpgrade = False
  mbCancel = False
  Call getVersions(SelectedDB, mlFixLevel, mlDbVersion, mlPostFixLevel, mlDBSubVersion)
  
  ' Find out what template file is or re-establish it
  Call xKill(GetStatic("NewFileTemplateDir") & "\template.mdb")
  strTemplateFile = GetDbTemplateFile(SelectedDB)
  If Not FileExists(strTemplateFile) Then
    strTemplateFile = EstablishTemplateFile(SelectedDB, Right$(strTemplateFile, 12))
    If StrComp(strTemplateFile, "", vbTextCompare) = 0 Then
      Err.Raise ERR_APPLY_FIXES, "NewTemplate", "The template file '" & strTemplateFile & "' could not be found."
    End If
  End If
 
  ' Initialise template database and get version infomation
  Set dbTemplate = InitDB(gwsMain, strTemplateFile, "Template File")
  If dbTemplate Is Nothing Then Err.Raise ERR_APPLY_FIXES, "NewTemplate", "The template file '" & strTemplateFile & "' could not be opened."
  Call getVersions(dbTemplate, mlTemplateFixLevel, mlTemplateDbVersion, mlTemplatePostFixLevel, mlTemplateDbSubVersion)
  
  ' Check versions
  If mlTemplateDbVersion < mlDbVersion Then
    Call Err.Raise(ERR_APPLY_FIXES, "NewTemplate", "The file which you are trying to open has been created using a later version of the template. This means that it will not work in your version of the software, which needs to be upgraded.")
  End If
  If mlTemplateFixLevel < mlFixLevel Then
    Call ErrorMessage(ERR_ERRORSILENT, Err, "NewTemplate", "Fix Level clash", "Fix level is higher in the file (" & mlFixLevel & ") than in the template (" & mlTemplateFixLevel & ") in database '" & SelectedDB.Name & "'.")
  End If
  If mlTemplatePostFixLevel < mlPostFixLevel Then
    Call ErrorMessage(ERR_ERRORSILENT, Err, "NewTemplate", "Post Fix Level clash", "Fix level is higher in the file (" & mlPostFixLevel & ") than in the template (" & mlTemplatePostFixLevel & ") in database '" & SelectedDB.Name & "'.")
  End If
  
  ' Check whether user wants to run fixes and update of database
  If mlTemplateDbVersion > mlDbVersion Then
    If MultiDialog("Open Database", "The file which you are trying to open requires updating. This may take several minutes depending upon the size of the file and the speed of your computer.  Are you sure you want to continue?", "Continue", "Cancel") = 2 Then
      mbCancel = True
      GoTo NewTemplate_End
    Else
      mbUpgrade = True
    End If
  ElseIf mlTemplateDbSubVersion > mlDBSubVersion Then
    If MultiDialog("Open Database", "The file which you are trying to open requires updating. This may take several minutes depending upon the size of the file and the speed of your computer.  Are you sure you want to continue?", "Continue", "Cancel") = 2 Then
      mbCancel = True
      GoTo NewTemplate_End
    Else
      mbUpgrade = True
    End If
  End If
  
  ' Apply fixes to selected database
  Call ApplyIndividualFixes(SelectedDB, dbTemplate, mlFixLevel, mlPostFixLevel, mlDbVersion, mlDBSubVersion)
  Call DisplayMessageKill
  
  If mbUpgrade Then
' Not needed if commented out below because of kernel32.dll error

'    ' Find out what is the local drive and what drive the file is on
'    Set fs = New FileSystemObject
'    Set drvLocal = fs.GetDrive(fs.GetDriveName(GetTempDirectory))
'    Set drvFile = fs.GetDrive(fs.GetDriveName(strFileName))
    
    ' Kill template database object in order to make copy and then copy
    Call dbTemplate.Close
    Set dbTemplate = Nothing
    Call SplitPath(strFileName, strDir, strTempFile, strExt)
    
    Call ProvideFeedback(0, 1, "Copying template")
    If FileCopyEx(strTemplateFile, strDir & strTempFile & "_New" & strExt) Then
      strNewFileName = strDir & strTempFile & "_New" & strExt
    End If
    Set NewDB = InitDB(gwsMain, strNewFileName, "New Abacus+ file")
    If Not CopyData(NewDB, SelectedDB, strFileName) Then Exit Function
    Call SelectedDB.Close
    Set SelectedDB = Nothing
  End If
    
' Had to comment out as produced kernel32.dll gpf on userappshutdown - no idea why!
    
'    ' The file is not saved locally
'    If StrComp(drvLocal.DriveLetter, drvFile.DriveLetter, vbTextCompare) <> 0 Then
'      ' Insufficient space to copy locally
'      If FileLen(strFileName) * 2.1 > drvLocal.FreeSpace Then
'        ' Exit if users does not wish to continue
'        If MultiDialog("Open Database", "The file which you are trying to open is " & FormatNumber(FileLen(strFileName) / 1024 ^ 2, 2) & " Mb.  There is only " & FormatNumber(drvLocal.FreeSpace / 1024 ^ 2, 2) & " Mb available on your local drive.  Files are copied locally to increase the speed of the upgrade operation and free space of at least twice the file size is required.  Are you sure you want to continue opening the file in the current location, which may be slow?", "Continue", "Cancel") = 2 Then
'          mbCancel = True
'          GoTo NewTemplate_End
'        ' Copy template database to file's present network location and transfer
'        ' data from original database to this new template copy
'        Else
'          Call ProvideFeedback(0, 1, "Copying template")
'          If FileCopyEx(strTemplateFile, strDir & strTempFile & "_New" & strExt) Then
'            strNewFileName = strDir & strTempFile & "_New" & strExt
'          End If
'          Set NewDB = InitDB(gwsMain, strNewFileName, "New Abacus+ file")
'          If Not CopyData(NewDB, SelectedDB, strFileName) Then Exit Function
'          Call SelectedDB.Close
'          Set SelectedDB = Nothing
'        End If
'      ' Sufficient space to copy locally
'      Else
'        ' Copy template database to temporary local location, copy original file to temporary
'        ' local location and transfer data between these two copies
'        Call ProvideFeedback(0, 1, "Copying template")
'        If FileCopyEx(strTemplateFile, GetTempDirectory & strTempFile & "_New" & strExt) Then
'          strNewFileName = GetTempDirectory & strTempFile & "_New" & strExt
'        End If
'        Set NewDB = InitDB(gwsMain, strNewFileName, "New Abacus+ file")
'        Call SelectedDB.Close
'        Set SelectedDB = Nothing
'        Call ProvideFeedback(0, 1, "Copying file")
'        If FileCopyEx(strFileName, GetTempDirectory & strTempFile & strExt) Then
'          strCopyFileName = GetTempDirectory & strTempFile & strExt
'        End If
'        Set CopySelectedDB = InitDB(gwsMain, strCopyFileName, "Copy of Original Abacus+ file")
'        If Not CopyData(NewDB, CopySelectedDB, strCopyFileName) Then Exit Function
'      End If
'    ' The file is saved locally
'    Else
'      ' Copy template database to file's present local location and transfer
'      ' data from original database to this new template copy
'      Call ProvideFeedback(0, 1, "Copying template")
'      If FileCopyEx(strTemplateFile, strDir & strTempFile & "_New" & strExt) Then
'        strNewFileName = strDir & strTempFile & "_New" & strExt
'      End If
'      Set NewDB = InitDB(gwsMain, strNewFileName, "New Abacus+ file")
'      If Not CopyData(NewDB, SelectedDB, strFileName) Then Exit Function
'      Call SelectedDB.Close
'      Set SelectedDB = Nothing
'    End If
      
  If mbUpgrade Then
    ' Apply post fixes to new database
    Call SyncMultiSchedules(NewDB, dbTemplate)
  Else
    Call ApplyIndividualPostFixes(SelectedDB, dbTemplate, mlFixLevel, mlPostFixLevel, mlDbVersion, mlDBSubVersion)
  End If
      
  If mbUpgrade Then
    ' Delete original file, copy the new file to original location and delete copies
    Call xKill(strFileName)
    If Not NewDB Is Nothing Then
      Call NewDB.Close
    End If
    Set NewDB = Nothing
    Call FileCopyEx(strNewFileName, strFileName)
    Call xKill(strNewFileName)
    If Not CopySelectedDB Is Nothing Then
      Call CopySelectedDB.Close
    End If
    Set CopySelectedDB = Nothing
    If Len(strCopyFileName) > 0 Then Call xKill(strCopyFileName)
    Set NewTemplate = InitDB(gwsMain, strFileName, "Updated Abacus+ File")
  Else
    Set NewTemplate = SelectedDB
  End If
    
NewTemplate_End:
  Call ClearCursor
  Call DisplayMessageKill
  If Not NewDB Is Nothing Then
    Call NewDB.Close
  End If
  Set NewDB = Nothing
  If Not CopySelectedDB Is Nothing Then
    Call CopySelectedDB.Close
  End If
  Set CopySelectedDB = Nothing
  Set fs = Nothing
  Set drvFile = Nothing
  Set drvLocal = Nothing
  Exit Function
  
NewTemplate_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NewTemplate", "Error updating database", "An error occurred updating the database")
  Resume NewTemplate_End
  Resume
End Function

Private Function CopyData(NewDB As Database, SelectedDB As Database, strFileName As String) As Boolean
  Dim NewTd As TableDef, SelectedTd As TableDef
  Dim fld As Field
  Dim sql As String, strFldList As String
  Dim j As Long

  On Error GoTo CopyData_Err
  
  ' Copy data from selected database to new database
  CopyData = False
  j = 1
  For Each NewTd In NewDB.TableDefs
    If InCollection(SelectedDB.TableDefs, NewTd.Name) Then
      If Not IsSysTable(NewTd) And StrComp(Left$(NewTd.Name, 4), "sys_", vbTextCompare) <> 0 Then
        Call ProvideFeedback(j, NewDB.TableDefs.Count, "Updating table " & NewTd.Name)
        Select Case NewTd.Name
          Case "Menu"
            Call CopyTable2(gwsMain, NewDB, "Menu_Selected", SelectedDB, "Menu", TBL_COPY_OVERWRITE, PROP_ALL)
            sql = "INSERT INTO Menu_Selected SELECT * FROM Menu IN '" & strFileName & "'"
            Call NewDB.Execute(sql, dbFailOnError)
            sql = ""
            sql = "UPDATE Menu INNER JOIN Menu_Selected ON Menu.ObjectName=Menu_Selected.ObjectName SET Menu.Complete=Menu_Selected.Complete, "
            sql = sql & "Menu.CompleteUser=Menu_Selected.CompleteUser, Menu.CompleteDate=Menu_Selected.CompleteDate, "
            sql = sql & "Menu.Review=Menu_Selected.Review, Menu.ReviewUser=Menu_Selected.ReviewUser, "
            sql = sql & "Menu.ReviewDate=Menu_Selected.ReviewDate, Menu.TaxpackSchedule=Menu_Selected.TaxpackSchedule"
          Case "Validation"
            Call CopyTable2(gwsMain, NewDB, "Validation_Selected", SelectedDB, "Validation", TBL_COPY_OVERWRITE, PROP_ALL)
            sql = "INSERT INTO Validation_Selected SELECT * FROM Validation IN '" & strFileName & "'"
            Call NewDB.Execute(sql, dbFailOnError)
            sql = ""
            sql = "UPDATE Validation INNER JOIN Validation_Selected ON Validation.SupportObjectName=Validation_Selected.SupportObjectName "
            sql = sql & "SET Validation.Cleared=Validation_Selected.Cleared, "
            sql = sql & "Validation.ClearedUser=Validation_Selected.ClearedUser, Validation.ClearedDate=Validation_Selected.ClearedDate"
          Case "TaxQuestions"
            Call CopyTable2(gwsMain, NewDB, "TaxQuestions_Selected", SelectedDB, "TaxQuestions", TBL_COPY_OVERWRITE, PROP_ALL)
            sql = "INSERT INTO TaxQuestions_Selected SELECT * FROM TaxQuestions IN '" & strFileName & "'"
            Call NewDB.Execute(sql, dbFailOnError)
            sql = ""
            sql = "UPDATE TaxQuestions INNER JOIN TaxQuestions_Selected ON TaxQuestions.QuestionCode=TaxQuestions_Selected.QuestionCode "
            sql = sql & "SET TaxQuestions.Answer=TaxQuestions_Selected.Answer"
          Case "TaxQuestionsDivisions"
            Call CopyTable2(gwsMain, NewDB, "TaxQuestionsDivisions_Selected", SelectedDB, "TaxQuestionsDivisions", TBL_COPY_OVERWRITE, PROP_ALL)
            sql = "INSERT INTO TaxQuestionsDivisions SELECT * FROM TaxQuestionsDivisions IN '" & strFileName & "'"
            Call NewDB.Execute(sql, dbFailOnError)
            sql = ""
          Case "ScheduleSetup"
            Call CopyTable2(gwsMain, NewDB, "ScheduleSetup_Selected", SelectedDB, "ScheduleSetup", TBL_COPY_OVERWRITE, PROP_ALL)
            sql = "INSERT INTO ScheduleSetup_Selected SELECT * FROM ScheduleSetup IN '" & strFileName & "'"
            Call NewDB.Execute(sql, dbFailOnError)
            sql = ""
            sql = "UPDATE ScheduleSetup INNER JOIN ScheduleSetup_Selected ON (ScheduleSetup.ObjectName=ScheduleSetup_Selected.ObjectName "
            sql = sql & "AND ScheduleSetup.ColumnField=ScheduleSetup_Selected.ColumnField) SET ScheduleSetup.Displayed=ScheduleSetup_Selected.Displayed"
          Case "TaxTypes"
            ' Insert any user defined tax treatments from selected database
            Call SyncTaxTypes(NewDB, SelectedDB)
          Case "WorkflowHelp", "ProcessHelp", "DistributeExcel"
            ' Do nothing
          Case Else
            If StrComp(NewTd.Name, "Divisions", vbTextCompare) = 0 Then
              Call NewDB.Execute("DELETE * FROM Divisions", dbFailOnError)
            End If
            Set SelectedTd = SelectedDB.TableDefs(NewTd.Name)
            For Each fld In NewTd.Fields
              If InCollection(SelectedTd.Fields, fld.Name) Then
                strFldList = strFldList & "[" & fld.Name & "]" & ", "
              End If
            Next
            strFldList = Left$(strFldList, Len(strFldList) - 2)
            sql = "INSERT INTO " & NewTd.Name & " (" & strFldList & ") SELECT " & strFldList & " FROM " & NewTd.Name & " IN '" & strFileName & "'"
        End Select
        If Len(sql) <> 0 Then
          Call NewDB.Execute(sql, dbFailOnError)
        End If
        strFldList = ""
        sql = ""
      Else
        If StrComp(NewTd.Name, "sys_Control", vbTextCompare) = 0 Then
          Call NewDB.Execute("DELETE * FROM sys_Control", dbFailOnError)
          sql = "INSERT INTO sys_Control SELECT * FROM sys_Control IN '" & strFileName & "'"
          Call NewDB.Execute(sql, dbFailOnError)
          mlDBSubVersion = setDBSubVersion(NewDB, mlTemplateDbSubVersion)
          mlDbVersion = setDBVersion(NewDB, mlTemplateDbVersion)
          mlPostFixLevel = setPostFixLevel(NewDB, mlTemplatePostFixLevel)
        End If
      End If
    End If
    j = j + 1
  Next
  ' Remove all tables copied in
  Call NewDB.TableDefs.Delete("Menu_Selected")
  Call NewDB.TableDefs.Delete("Validation_Selected")
  Call NewDB.TableDefs.Delete("TaxQuestions_Selected")
  Call NewDB.TableDefs.Delete("TaxQuestionsDivisions_Selected")
  Call NewDB.TableDefs.Delete("ScheduleSetup_Selected")
  CopyData = True

CopyData_End:
  Call ClearCursor
  Set fld = Nothing
  Set NewTd = Nothing
  Set SelectedTd = Nothing
  Exit Function
  
CopyData_Err:
  CopyData = False
  Call ErrorMessage(ERR_ERROR, Err, "CopyData ", "Error copying data", "An error occurred copying data between files.")
  Resume CopyData_End
  Resume
End Function

Private Sub getVersions(db As Database, lFixLevel As Long, lVersion As Long, lPostFixLevel As Long, lDBSubVersion As Long)
  Dim rs As Recordset
  
  On Error GoTo getVersions_err
  
  Set rs = db.OpenRecordset("SELECT * FROM sys_Control", dbOpenDynaset, dbFailOnError)
  rs.FindFirst ("Flag=" & StrSQL("TPFixLevel"))
  If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "initVars", "The fix level for the file cannot be found."
  lFixLevel = rs!State
  rs.FindFirst ("Flag=" & StrSQL("TPDBVersion"))
  If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "initVars", "The internal version for the file cannot be found."
  lVersion = rs!State
  rs.FindFirst ("Flag=" & StrSQL("TPPostFixLevel"))
  If rs.NoMatch Then
    rs.AddNew
      rs!Flag = "TPPostFixLevel"
      rs!State = 0
    rs.Update
    lPostFixLevel = 0
  Else
    lPostFixLevel = rs!State
  End If
  rs.FindFirst ("Flag=" & StrSQL("TPDBSubVersion"))
  If rs.NoMatch Then
    rs.AddNew
      rs!Flag = "TPDBSubVersion"
      rs!State = 0
    rs.Update
    lDBSubVersion = 0
  Else
    lDBSubVersion = rs!State
  End If
  
getVersions_end:
  Set rs = Nothing
  Exit Sub
  
getVersions_err:
  Err.Raise Err.Number, ErrorSource(Err, "getVersions"), "Error getting the database version number." & vbCrLf & Err.Description
  Resume
End Sub

Public Property Get Upgrade() As Boolean
  Upgrade = mbUpgrade
End Property

Public Property Get Cancel() As Boolean
  Cancel = mbCancel
End Property

Public Function Kill()
  On Error Resume Next
  
  Set gNotify = Nothing
  If Not dbTemplate Is Nothing Then
    dbTemplate.Close
  End If
  Set dbTemplate = Nothing
  Call DoDBEvents(REFRESH_CACHE)
End Function

