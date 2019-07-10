VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBOFixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

Private mlFixLevel As Long
Private mlPostFixLevel As Long
Private mlDbVersion As Long
Private mlDBSubVersion As Long
Private mOverride As Boolean
Private mlTemplateFixLevel As Long
Private mlTemplateDbVersion As Long
Private mlTemplatePostFixLevel As Long
Private mlTemplateDbSubVersion As Long
Private dbTemplate As Database

Private Sub Class_Terminate()
  Call Kill
End Sub

Public Function ApplyDAFixes(db As Database, iNotify As IBaseNotify) As Boolean
  Dim sTemplateFile As String
  Dim rsVersion As Recordset, rsFixLevel As Recordset
  Dim errmsg As String
  Dim sync As dbSync
  Dim flt As syncFilter
  Dim sTempFile As String, sExt As String
  Dim bExclusive As Boolean
  
  On Error GoTo ApplyDAFixes_err
  
  ApplyDAFixes = True
  Call getVersions(db, mlFixLevel, mlDbVersion, mlPostFixLevel, mlDBSubVersion)
  Call xKill(GetStatic("NewFileTemplateDir") & "\template.mdb")
  sTemplateFile = GetDbTemplateFile(db)
  If Not FileExists(sTemplateFile) Then
    sTemplateFile = EstablishTemplateFile(db, Right$(sTemplateFile, 12))
    If StrComp(sTemplateFile, "", vbTextCompare) = 0 Then
      Err.Raise ERR_APPLY_FIXES, "ApplyDAFixes", "The template file '" & sTemplateFile & "' could not be found."
    End If
  End If
  Call SplitPath(sTemplateFile, , sTempFile, sExt)
  If FileCopyEx(sTemplateFile, GetTempDirectory & sTempFile & sExt) Then
    sTemplateFile = GetTempDirectory & sTempFile & sExt
    bExclusive = True
  End If
  Set dbTemplate = InitDB(gwsMain, sTemplateFile, "Tax Pack template file", , bExclusive, True)
  If dbTemplate Is Nothing Then Err.Raise ERR_APPLY_FIXES, "ApplyDAFixes", "The template file '" & sTemplateFile & "' could not be opened."
  Call getVersions(dbTemplate, mlTemplateFixLevel, mlTemplateDbVersion, mlTemplatePostFixLevel, mlTemplateDbSubVersion)

  ' Check whether file is of later version than template
  If mlTemplateDbVersion < mlDbVersion Then
    Call Err.Raise(ERR_APPLY_FIXES, "ApplyDAFixes", "The file which you are trying to open has been created using a later version of the template. This means that it will not work in your version of the software, which needs to be upgraded.")
  End If
  If mlTemplateFixLevel < mlFixLevel Then
    Call ErrorMessage(ERR_ERRORSILENT, Err, "ApplyDAFixes", "Fix Level clash", "Fix level is higher in the file (" & mlFixLevel & ") than in the template (" & mlTemplateFixLevel & ") in database '" & db.Name & "'.")
  End If
  If mlTemplatePostFixLevel < mlPostFixLevel Then
    Call ErrorMessage(ERR_ERRORSILENT, Err, "ApplyDAFixes", "Post Fix Level clash", "Fix level is higher in the file (" & mlPostFixLevel & ") than in the template (" & mlTemplatePostFixLevel & ") in database '" & db.Name & "'.")
  End If
  
  ' Check whether user wants to run fixes and update of database
  If mlTemplateDbVersion > mlDbVersion Then
    If MultiDialog("Open Database", "The file which you are trying to open requires updating.  This could take between 5 and 20 minutes dependent upon the size of the file and the speed of your computer.  Are you sure you want to continue?", "Continue", "Cancel") = 2 Then
      ApplyDAFixes = False
      GoTo ApplyDAFixes_end
    End If
  ElseIf mlTemplateDbSubVersion > mlDBSubVersion Then
    If MultiDialog("Open Database", "The file which you are trying to open requires updating.  This could take between 5 and 20 minutes dependent upon the size of the file and the speed of your computer.  Are you sure you want to continue?", "Continue", "Cancel") = 2 Then
      ApplyDAFixes = False
      GoTo ApplyDAFixes_end
    End If
  End If
  
  ' Apply pre fixes
  Call ApplyIndividualFixes(db, dbTemplate, mlFixLevel, mlPostFixLevel, mlDbVersion, mlDBSubVersion)

  ' Synchronise database
  If mlTemplateDbVersion > mlDbVersion Then
    Set sync = New dbSync
    sync.FieldPropertyChecks = "AllowZeroLength;Attributes;DefaultValue;Required;Size;Type;ValidationRule;ValidationText;OrdinalPosition"
    Call sync.AddFilter("sys_*", INCLUDE_FILTER, DELETE_QUERIES + DELETE_TABLES + OVERWRITE_DATA + OVERWRITE_QUERY_STRUCTURE + OVERWRITE_TABLE_STRUCTURE, "sys_Control")
    Call sync.AddFilter("sys_*", EXCLUDE_FILTER, DELETE_QUERIES + DELETE_TABLES + OVERWRITE_QUERY_STRUCTURE + OVERWRITE_TABLE_STRUCTURE)
    Call sync.AddFilter("sys_control", INCLUDE_FILTER, DELETE_QUERIES + DELETE_TABLES + OVERWRITE_QUERY_STRUCTURE + OVERWRITE_TABLE_STRUCTURE)
    Call sync.SynchroniseDatabases(dbTemplate, db, iNotify)
    mlDbVersion = setDBVersion(db, mlTemplateDbVersion)
    mlDBSubVersion = setDBSubVersion(db, mlTemplateDbSubVersion)
  ElseIf mlTemplateDbSubVersion > mlDBSubVersion Then
    Set sync = New dbSync
    Call sync.AddFilter("*", INCLUDE_FILTER, DELETE_QUERIES + OVERWRITE_QUERY_STRUCTURE)
    Call sync.SynchroniseDatabases(dbTemplate, db, iNotify)
    mlDBSubVersion = setDBSubVersion(db, mlTemplateDbSubVersion)
  End If
  
  ' Apply post fixes
  Call ApplyIndividualPostFixes(db, dbTemplate, mlFixLevel, mlPostFixLevel, mlDbVersion, mlDBSubVersion)

ApplyDAFixes_end:
  Call DisplayMessageKill
  Set dbTemplate = Nothing
  Set rsFixLevel = Nothing
  Set rsVersion = Nothing
  Exit Function

ApplyDAFixes_err:
  ApplyDAFixes = (False Or mOverride)
  Call ErrorMessage(ERR_ERROR, Err, "ApplyDAFixes", "Error checking version", "Each time a TBO file is opened, the file is compared against the template file to ensure that it is fully up to date. The two files will then be synchronised. An error has occurred whilst running this process.")
  Resume ApplyDAFixes_end
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
 
Private Sub UpdateSystemTables(db As Database, dbTemplate As Database)
  Dim td As TableDef
  
  On Error GoTo UpdateSystemTables_err
  
  For Each td In dbTemplate.TableDefs
    If StrComp(Left$(td.Name, 4), "sys_", vbTextCompare) = 0 Then
      If TablePresent(db.TableDefs, td.Name) Then
        If Not SystemTableException(td.Name) Then
          If Not CopyTable2(gwsMain, db, td.Name, dbTemplate, td.Name, TBL_COPY_OVERWRITE, PROP_ALL, "SELECT * FROM " & td.Name, True, False) Then Err.Raise ERR_COPY_TEMPLATE_FILE, "UpdateSystemTables", "An error occurred updating the system table '" & td.Name & "'."
        End If
      End If
    End If
  Next td
  
UpdateSystemTables_end:
  Exit Sub
  
UpdateSystemTables_err:
  Err.Raise Err.Number, ErrorSource(Err, "UpdateSystemTables"), "An error occurred updating the system tables."
End Sub

Private Function SystemTableException(sTableName As String) As Boolean
  
  If (StrComp(sTableName, "sys_Control", vbTextCompare) = 0) _
  Then
    SystemTableException = True
  End If
End Function

Private Sub UpdateQueryDefs(db As Database, dbTemplate As Database)
  Dim qd As QueryDef
  
  On Error GoTo UpdateQueryDefs_err
  
  For Each qd In dbTemplate.QueryDefs
    If QueryPresent(db.QueryDefs, qd.Name) Then
      If qd.LastUpdated > db.QueryDefs(qd.Name).LastUpdated Then
        Call AddNewQuery(db, qd, True)
      End If
    Else
      Call AddNewQuery(db, qd, True)
    End If
  Next qd
  
UpdateQueryDefs_end:
  Exit Sub
  
UpdateQueryDefs_err:
  Err.Raise Err.Number, ErrorSource(Err, "UpdateQueryDefs"), "An error occurred updating the system queries."
End Sub

Public Property Let ContinueOnFixError(NewVal As Boolean)
  mOverride = NewVal
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

