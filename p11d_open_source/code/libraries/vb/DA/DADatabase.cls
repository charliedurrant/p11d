VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DADatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private mFileName As String
Private mDbOpen As Boolean
Private mProgID As String
Private mstrMode As String
Private mPQRS As ParsedQueriesRS
Private mCalcMethod As CALC_METHOD
Private mbUpdateCompleted As Boolean

Public Enum RS_TYPE
  QUERY_HEADER_RS = 1
  QUERY_DETAIL_RS
  QUERY_RELATIONS_RS
  QUERY_CACHE
End Enum

Public Enum CALC_METHOD
  AUTO_RECALC
  MANUAL_RECALC
  MANUAL_RECALC_WARN
End Enum

Private mrsControl As recordset
Private mrsCalcNo As recordset

Private mBatch As Collection
Private mTouch As Collection

Dim intlockcount As Long
Dim mbMode As Boolean             ' enhance with applies for multi packs
Dim mForceMode As Boolean

Dim strbatchexecute() As String   ' remove as replaced with collection of batchitems

Private db As database
Dim idString As String

Dim objqueries As New Collection
Dim objRecordsets As New Collection
Dim intver As Long
Private mLocks As Collection   ' collection of locks on the connection by other applications

Public Property Get name() As String
  name = mFileName
End Property

Public Property Let name(sFile As String)
  mFileName = sFile
End Property

Public Property Get Locks() As Collection
  Set Locks = mLocks
End Property

Public Property Set Locks(NewLocks As Collection)
  Set mLocks = NewLocks
End Property

Public Property Get AccessMode() As Boolean
  AccessMode = mbMode
End Property

Friend Property Get ParsedQueries() As ParsedQueriesRS
  Set ParsedQueries = mPQRS
End Property

Public Property Let AccessMode(NewMode As Boolean)
  
  On Error GoTo AccessMode_err
  Call xSet("AccessMode")
  
  If mDbOpen Then
    Call ForceCloseDatabase
  End If
  If NewMode Then
    If IsDatabaseOpen(mFileName) Then logerr "The database " & mFileName & " is in use by another user."
  End If
  If mDbOpen Then
    mForceMode = True
    OpenDatabase mFileName, NewMode
    mForceMode = False
  End If

AccessMode_end:
  mbMode = NewMode
  Call xReturn("AccessMode")
  Exit Property

AccessMode_err:
  Err.Raise ERR_ACCESS_MODE, Err.Source, Err.Description
End Property

Friend Function OpenDB(Filename As String, idString As String, Optional dbExclusive As Boolean = False, Optional bTBONewFixes As Boolean = True) As Boolean
  Dim qd As QueryDef
  Dim td As TableDef
  Dim q As Query
  Dim dummy As Long
  Dim rsQuery As recordset
  Dim i As Long
  Dim j As Long
  Const k = 0.01
  Dim TBOFix As TBOFixes
  Dim TBONewFixes As TBONewFixes
  
  Call xSet("OpenDb")
  
  On Error GoTo Opendb_err:
  If mDbOpen Then Err.Raise ERR_DBOPEN, "OpenDatabase", "The database '" & Filename & "' is already open."
  
  If bTBONewFixes Then
    If Not OpenDatabase(Filename, mbMode) Then Exit Function
    Set TBONewFixes = New TBONewFixes
    Set db = TBONewFixes.NewTemplateFile(db, Filename, gNotify)
    If db Is Nothing Then
      If Not TBONewFixes.Upgrade And Not TBONewFixes.Cancel Then
        ' Do nothing since upgrade not required
      Else
        OpenDB = False
        GoTo Opendb_end
      End If
    Else
      If TBONewFixes.Upgrade Then mbUpdateCompleted = True
      If Not OpenDatabase(Filename, mbMode) Then Exit Function
    End If
  Else
    If Not OpenDatabase(Filename, mbMode) Then Exit Function
    Call ShowProgress("Checking " & Filename, "Opening Database", 2 * k)
    Set TBOFix = New TBOFixes
    If Not TBOFix.ApplyDAFixes(db, gNotify) Then
      OpenDB = False
      GoTo Opendb_end
    End If
  End If
  
  mFileName = Filename
  gOpenDatabases.Add Me, Filename
  Call inclockcount
  mDbOpen = True
  Set mrsControl = db.OpenRecordset("SELECT * FROM sys_Control")
  Set mrsCalcNo = mrsControl.Clone
  mrsCalcNo.FindFirst "Flag=" & strSql("CALCNO")
  If mrsCalcNo.NoMatch Then
    mrsCalcNo.addnew
    mrsCalcNo!flag = "CALCNO"
    mrsCalcNo!State = 0
    mrsCalcNo.update
    mrsCalcNo.FindFirst "Flag=" & strSql("CALCNO")
  End If
  Call ShowProgress("Checking " & Filename, "Opening Database", 2 * k)
  
  intver = Me.flag("id", dummy, mProgID)
  
  'need to upgrade version
  ' setup tables

  'zzzz identify tables for loading through query rather than tabledefs
  Set rsQuery = db.OpenRecordset("select * from sys_Cache_Tables_for_Loading", dbOpenSnapshot)
  i = 1
  If Not rsQuery.eof Then
      j = rsQuery.RecordCount
      While Not rsQuery.eof
        Call ShowProgress("Loading table " & CStr(i) & " of " & CStr(rsQuery.RecordCount), "Opening Database", k * 10 / j)
        i = i + 1
        Set q = New Query
        Call q.setuptable(Me, rsQuery!name)
        rsQuery.movenext
      Wend
  End If
  
  ' zzzz remove redundant entries from cache (sys_query_header, sys_query_relations) use sys_... delete query
  Call db.execute("sys_Cache_Delete_Queries_Header")
  Call db.execute("sys_Cache_Delete_Queries_Relations")
  Call db.execute("sys_Cache_Delete_Queries_Relationships")

  ' zzzz identify cache for loading queries
  Set rsQuery = db.OpenRecordset("select * from sys_Cache_Queries_for_Loading", dbOpenSnapshot)
  i = 1
  If Not rsQuery.eof Then
    j = rsQuery.RecordCount
    While Not rsQuery.eof
      Call ShowProgress("Loading query " & CStr(i) & " of " & CStr(rsQuery.RecordCount) & vbCr & rsQuery.fields("Name"), "Opening Database", k * 65 / j)
      i = i + 1
        ' load select and non select queries
        Set q = New Query
        Call q.LoadQueryName(Me, rsQuery.fields("Name"))
      rsQuery.movenext
    Wend
  End If
  
  ' need to identify when and why we cache select queries
  
  ' identify outstanding queries which require setup and then caching
  ' this should not be run on live files (only on the template/development)
  
  ' Updating queries
  Set rsQuery = db.OpenRecordset("select * from sys_Cache_Queries_for_Updating", dbOpenSnapshot)
  i = 1
  If Not rsQuery.eof Then
    j = rsQuery.RecordCount
    While Not rsQuery.eof
      Call ShowProgress("Updating query " & CStr(i) & " of " & CStr(rsQuery.RecordCount) & vbCr & rsQuery.fields("Name"), "Opening Database", k * 10 / j)
      i = i + 1
      ' a query
      Set qd = db.QueryDefs(rsQuery.fields("Name"))
      If InCollection(objqueries, rsQuery.fields("Name")) Then
        Set q = objqueries(rsQuery.fields("Name"))
        q.Kill
        'objqueries.Remove rsQuery.fields("Name")
      End If
      ' only update non selects
      Set q = New Query
      q.Transient = False
      q.setupquery Me, qd
      rsQuery.movenext
    Wend
  End If
  
  Call ShowProgress("Establishing relations", "Opening Database", 2 * k)
  For Each q In objqueries
    q.linkchildren
  Next q
  Call ShowProgress("Establishing relations", "Opening Database", 2 * k)

  For Each q In objqueries
    q.linkparents
  Next q
  
  dummy = Me.flag("MODE", dummy, mstrMode)
'  If Not mbMode Then
'    AccessMode = False
'  End If
  If Not TBOFix Is Nothing Then
    Call TBOFix.Kill
    Set TBOFix = Nothing
  End If
  If Not TBONewFixes Is Nothing Then
    Call TBONewFixes.Kill
    Set TBONewFixes = Nothing
  End If
  OpenDB = True
  
Opendb_end:
  Set rsQuery = Nothing
  Set qd = Nothing
  Set td = Nothing
  Set q = Nothing
  Set TBOFix = Nothing
  Set TBONewFixes = Nothing
  Call DisplayMessageKill
  Call ResetProgress
  Call xReturn("OpenDb")
  Exit Function
Opendb_err:
  OpenDB = False
  Call ErrorMessage(ERR_ERROR, Err, "tcsda.OpenDb", "Opening Database", "An error has occurred opening " & Filename)
  Resume Opendb_end
  Resume
  Err.Raise ERR_OPEN_DB, "OpenDB", "An error occurred opening the database '" & Filename & "'" & vbCrLf & Err.Description
End Function

Public Property Get UpdateCompleted() As Boolean
  UpdateCompleted = mbUpdateCompleted
End Property

Friend Function Recordsets() As Collection
  Set Recordsets = objRecordsets
End Function

Public Function CloseDbForced(sAppName As String, bForceClose As Boolean)
  Dim s As String
  
  On Error GoTo CloseDbForced_Err
  Call xSet("CloseDbForced")
  CloseDbForced = True
  Debug.Print "File close started for file '" & mFileName & "' by " & sAppName
  If bForceClose Then
    Call ForceCloseDb
    GoTo CloseDbForced_End
  End If
  If mDbOpen Or Not db Is Nothing Then
    If Not mLocks Is Nothing Then
      If InStrCollection(mLocks, sAppName) Then
        mLocks.Remove sAppName
        Debug.Print "File close continued for file '" & mFileName & "' by " & sAppName
        If mLocks.Count = 0 Then
          ForceCloseDb
        End If
      ElseIf mLocks.Count = 0 Then
        ForceCloseDb
'      Else
'        s = mLocks(1)
'        Err.Raise ERR_NO_CONNECTION, "Close Database", "The application " & s & " is currently accessing the database and so it can not be closed." & vbCr & vbCr & "When " & s & " releases the database, then it will close."
      End If
    Else
      ForceCloseDb
    End If
  End If

CloseDbForced_End:
  Call xReturn("CloseDbForced")
  Exit Function

CloseDbForced_Err:
  CloseDbForced = False
  Call ErrorMessage(ERR_ERROR, Err, "CloseDbForced", "Error in CloseDbForced", "Undefined error.")
  Resume CloseDbForced_End
  Resume
End Function

Public Function CloseDb(sAppName As String) As Boolean
  Call CloseDbForced(sAppName, False)
End Function

Friend Sub ForceCloseDb()
  Dim q As Query
  
  On Error GoTo err_ForceCloseDb
  Call ShowProgress("Closing File", "Closing Database")
  If mDbOpen Then
    ForceCloseDatabase
    For Each q In objqueries
      q.Kill
    Next q
    If InCollection(gOpenDatabases, mFileName) Then gOpenDatabases.Remove mFileName
    mDbOpen = False
  End If
  If Not mPQRS Is Nothing Then Call mPQRS.Kill
  Set mPQRS = Nothing
  Set mrsControl = Nothing
  Set mrsCalcNo = Nothing
  Set mBatch = Nothing
  Set objqueries = Nothing
  Set objRecordsets = Nothing
  If Not db Is Nothing Then
    Call db.Close
    Set db = Nothing
  End If
  Call DoDBEvents(REFRESH_CACHE)
  Debug.Print "File closed (" & mFileName & ")"
  Call DisplayMessageKill
  Exit Sub
  
err_ForceCloseDb:
  logerr "Error in ForceCloseDb "
  Resume Next
  Resume
End Sub

Private Sub Class_Initialize()
  mDbOpen = False
  ReDim strbatchexecute(0)
  Set mLocks = New Collection
  Set mPQRS = New ParsedQueriesRS
End Sub

Private Sub Class_Terminate()
  'Me.CloseDb
  Me.ForceCloseDb
  Set mLocks = Nothing
End Sub

Public Function database() As database
  If mDbOpen Then
    Set database = db
  Else
    Set database = Nothing
  End If
End Function

Public Function QueryforVal(sqlstring As String) As String
Dim dars As DARecordset

  On Error GoTo QueryforVal_Err
  QueryforVal = "0"
  Set dars = Me.OpenRecordset(sqlstring)
  #If DEBUGVER Then
    logmessage "DADatabase.QueryforVal for SQL: " & sqlstring
  #End If
  If Not dars Is Nothing Then
    If Not (dars.eof And dars.bof) Then QueryforVal = CStr("" & dars.recordset.fields(0).Value)
  End If
  
QueryforVal_End:
  If Not dars Is Nothing Then dars.closedars
  Set dars = Nothing
  Exit Function

QueryforVal_Err:
  Call ErrorMessage(ERR_ERROR, Err, "QueryforVal", "Error in QueryforVal", "Undefined error.")
  Resume QueryforVal_End
  Resume
End Function

Public Function OpenRecordset(sqlstring As String) As DARecordset
  Dim q As Query
  Dim qd As QueryDef
  Dim retry As Long
      
  On Error GoTo OpenRecordset_Err
  
  If InCollection(objqueries, sqlstring) Then
    #If DEBUGVER Then
      logmessage "DADatabase.OpenRecordset in ObjQueries for SQL: " & sqlstring
    #End If
    Set q = objqueries(sqlstring)
  Else
    If InCollection(db.QueryDefs, sqlstring) Then
      #If DEBUGVER Then
        logmessage "DADatabase.OpenRecordset pre CreateQuery in QueryDefs for SQL: " & sqlstring
      #End If
      Set qd = db.QueryDefs(sqlstring)
      Set q = CreateQuery(qd, "", False)
      qd.Close
    Else
      #If DEBUGVER Then
        logmessage "DADatabase.OpenRecordset pre CreateQuery not in QueryDefs for SQL: " & sqlstring
      #End If
      Set q = CreateQuery(Nothing, sqlstring, True)
    End If
  End If
  Set OpenRecordset = q.createdars
  
OpenRecordset_End:
  Exit Function

OpenRecordset_Err:
  Select Case Err.Number
  Case 3012
    retry = retry + 1
    If retry < 10 Then
      Resume
    Else
      logerr "Error in openrecordset."
      Resume OpenRecordset_End
    End If
  Case Else
    logerr "Error in openrecordset."
    Resume OpenRecordset_End
  
  End Select
Resume
End Function

Public Property Get flag(flagname As String, calcno As Long, Optional flagtext As Variant) As Long
  On Error GoTo err_getflag
  If mrsControl Is Nothing Then Err.Raise ERR_DBOPEN, "TCSDA.FLAG", "No Control Recordset for identifying database settings"
    
  mrsControl.FindFirst "Flag='" + UCase$(flagname) + "'"
  If mrsControl.NoMatch Then
    flag = 0
    calcno = 0
    Exit Property
  End If
  flag = xLng(mrsControl!State)
  calcno = xLng(mrsControl!calcno)
  If Not IsMissing(flagtext) Then
    flagtext = IsNullEx(mrsControl!text, "")
  End If
end_getflag:
  Exit Property
  
err_getflag:
  logerr "Error in get flag "
  flag = 0
  calcno = 0
  Resume end_getflag
  Resume
End Property

Public Property Let flag(flagname As String, calcno As Long, Optional flagtext As Variant, val As Long)
  On Error GoTo err_letflag


  mrsCalcNo.edit
  calcno = mrsCalcNo!State + 1
  mrsCalcNo!State = calcno
  
  mrsControl.FindFirst "Flag='" + UCase$(flagname) + "'"
  If mrsControl.NoMatch Then
    mrsControl.addnew
    mrsControl!flag = flagname
    mrsControl!State = val
    mrsControl!calcno = calcno
    If Not IsMissing(flagtext) Then mrsControl!text = flagtext
  Else
    mrsControl.edit
    mrsControl!State = val
    mrsControl!calcno = calcno
    If Not IsMissing(flagtext) Then mrsControl!text = flagtext
  End If
  mrsControl.update
  mrsCalcNo.update
  
  Exit Property
  
err_letflag:
  logerr "Error in letflag " + Err.Description
Rem need to handle locking contention
End Property

Friend Sub inclockcount()
  intlockcount = intlockcount + 1
End Sub

Friend Function queries() As Collection
  Set queries = objqueries
End Function

Public Sub BatchExecute(sqlstring As String, Optional SetCalcFlags As Boolean = True)
  Dim i As Long
  Dim c As calc
  Dim q As Query
  Dim qd As QueryDef
  Dim retry As Long
  Static j As Long
  Static k As Long
  Dim bKillPop As Boolean
  Dim b As BatchItem
    
  ' batch execute will set boocalced for children to false, if SetCalcFlags is true
  ' if SetCalcFlags is false then it will expect user to handle touching
  ' touching can be done through batchexecute
  ' it will not calc parents and will not derive parents
  ' zzzz batch execute should use transactions to allow undo all if fail (apart from multiuser worries)
  ' touching should perhaps be done during the execute rather than as preparing

  On Error GoTo Batchexecute_ERR
  Call xSet("BatchExecute")
  
  If mBatch Is Nothing Then Set mBatch = New Collection
  
  ' Execute all the batch items
  If StrComp("EXECUTE", Trim$(sqlstring), vbTextCompare) = 0 Then
    bKillPop = True
    For j = 1 To mBatch.Count
      Set b = mBatch("b" & CStr(j))
      Call ShowProgress("Executing " & IIf(mBatch.Count > 1, CStr(j) & " of " & CStr(mBatch.Count) & " queries", " query"), "Executing")
      ' If calc of batch item is true and therefore queries have been created
      If b.calc Then
        ' Calc each batch item
        If InCollection(queries, b.sql) Then
          Set q = queries(b.sql)
          q.Calc_calced(Nothing) = False
          Set c = q
          c.calc
          Set c = Nothing
          q.Kill
          Set q = Nothing
          #If DEBUGVER Then
            logmessage "DADatabase.BatchExecute in queries for calced batch " & j & " and SQL: " & b.sql
          #End If
        ' Shouldn't arise - raise error
        Else
          Err.Raise ERR_SQL_PARSE, "BATCHEXECUTE", "Failed to locate " & vbCr & strbatchexecute(i) & vbCr & " in executables queries.  Will execute anyway, but will fail to set calc flags."
          #If DEBUGVER Then
            logmessage "DADatabase.BatchExecute not in queries for calced batch " & j & " and SQL: " & b.sql
          #End If
          db.execute b.sql
        End If
      ' If calc of batch item is false then directly execute query without calcing
      Else
        #If DEBUGVER Then
          logmessage "DADatabase.BatchExecute for non-calced batch " & j & " and SQL: " & b.sql
        #End If
        db.execute b.sql
      End If
    Next j
    j = 0
    Set mBatch = Nothing
    Set mBatch = New Collection
    Call ShowProgress("Setting calc status", "Executing")
    If mTouch Is Nothing Then Set mTouch = New Collection
    For Each b In mTouch
      Call touch(Trim$(Mid$(b.sql, 6)))
    Next b
    Set mTouch = Nothing
    Set mTouch = New Collection
    k = 0
  ElseIf InStr(1, sqlstring, "TOUCH ", vbTextCompare) = 1 Then
    If mTouch Is Nothing Then Set mTouch = New Collection
    k = k + 1
    If Not InCollection(mTouch, sqlstring) Then
      Set b = New BatchItem
      b.sql = sqlstring
      Call mTouch.Add(b, b.sql)
    End If
  ' Create batch items and if set calc flags then create queries
  Else
    j = j + 1
    Call ShowProgress("Preparing executable " & CStr(j), "Executing")
    Set b = New BatchItem
    b.calc = SetCalcFlags
    b.sql = sqlstring
    Call mBatch.Add(b, "b" & CStr(j))
    If SetCalcFlags Then
      If Not InCollection(queries, sqlstring) Then
        Set c = CreateQuery(Nothing, sqlstring, True, False)
      End If
    End If
  End If
  
Batchexecute_END:
  If bKillPop Then Call DisplayMessageKill
  Set q = Nothing
  xReturn ("BatchExecute")
  Exit Sub
  
Batchexecute_ERR:
  logerr "Error in database batchexecute " & sqlstring
  Resume Next
  Resume
End Sub

Public Property Let Mode(val As String)
  Dim c As calc
  Dim dars As DARecordset
  Dim calcno As Long
 ' zzzz check whether mode has changed
 On Error GoTo Mode_err
  mstrMode = val
  flag("MODE", calcno, mstrMode) = 0
  For Each c In queries
    c.calced(Nothing) = False
  Next c
  For Each dars In Recordsets
    dars.notify
  Next dars
Mode_end:
  Exit Property
Mode_err:
  Call logerr("Error setting Mode to " & val)
  Resume Mode_end
  
End Property

Public Property Get Mode() As String
  Mode = mstrMode
End Property

Public Function touch(TableName As String)
  Dim c As calc
  Dim dars As DARecordset
  On Error GoTo touch_err
  If InCollection(objqueries, TableName) Then
    #If DEBUGVER Then
      logmessage "Touching " & TableName & " in DADatabase.Touch"
    #End If
    Set c = objqueries(TableName)
    c.calced(Nothing) = False
    For Each dars In Recordsets
      dars.notify
    Next dars
  End If
touch_end:
  Set dars = Nothing
  Set c = Nothing
  Exit Function
touch_err:
  Call logerr("Error marking " & TableName & " for calculation.")
  Resume touch_end
End Function

Public Function DirtyAll() As Boolean
  Dim q As Query
  
  On Error GoTo DirtyAll_Err
  
  For Each q In objqueries
    If q.QueryType = dbQTable Then
      Call touch(q.name)
    End If
  Next
  DirtyAll = True
  
DirtyAll_End:
  Exit Function

DirtyAll_Err:
  DirtyAll = False
  Call Err.Raise(ERR_DIRTY_ALL, "DirtyAll", "An error occurred in marking the data as uncalculated.", Err.HelpFile, Err.HelpContext)
End Function

Public Function CleanAll() As Boolean
  Dim q As Query
  
  On Error GoTo CleanAll_Err
  
  For Each q In objqueries
    If q.QueryType = dbQTable Then
      q.Calc_calced(Nothing) = True
    End If
  Next
  CleanAll = True
  
CleanAll_End:
  Exit Function

CleanAll_Err:
  CleanAll = False
  Call Err.Raise(ERR_CLEAN_ALL, "CleanAll", "An error occurred in marking the data as calculated.", Err.HelpFile, Err.HelpContext)
End Function

Public Sub run(spname As String)
  Dim qd As QueryDef
  On Error GoTo run_err
  
  Set qd = db.QueryDefs(spname)
  qd.execute
  
run_end:
  Set qd = Nothing
  Exit Sub
  
run_err:
  Call logerr("Error in run " & spname)
  Resume run_end
End Sub

Public Function daquerydef(qname As String) As Query
Dim qd As QueryDef
  On Error GoTo err_openquerydef
 
  If Not InCollection(objqueries, qname) Then
    If Not InCollection(db.QueryDefs, qname) Then
      Err.Raise ERR_NOT_IN_LIST, "tcsda.daquerydef", "There is a missing expected query, " & qname
    Else
      Set qd = db.QueryDefs(qname)
      Set daquerydef = CreateQuery(qd, "", False)
    End If
  Else
    Set daquerydef = objqueries(qname)
  End If
  
err_openquerydef_End:
  Set qd = Nothing
  Exit Function
  
err_openquerydef:
  logerr "Error in open querydef " & qname
  Resume err_openquerydef_End
  Resume
End Function

Friend Function CreateQuery(Optional qd As QueryDef = Nothing, Optional sql As String, Optional Transient As Boolean = True, Optional HasParents As Boolean = True) As Query
  If qd Is Nothing Then
    If Len(sql) = 0 Then
      Set CreateQuery = Nothing
    Else
      #If DEBUGVER Then
        logmessage "DADatabase.CreateQuery with qd for SQL: " & sql
      #End If
      Set CreateQuery = CreateQuerySQL(sql, Transient, HasParents)
    End If
  Else
    #If DEBUGVER Then
      logmessage "DADatabase.CreateQuery without qd for SQL: " & sql
    #End If
    Set CreateQuery = CreateQueryQD(qd, Transient)
  End If
  
End Function

Private Function CreateQueryQD(ByVal qd As QueryDef, Optional ByVal Transient As Boolean = True) As Query
Dim q As Query
Dim rs As recordset

  On Error GoTo CreateQueryQD_err
  Call xSet("CreateQueryQD")
  
  ' use qd as need to create the object and cache for a uncached querydef
  
  Set q = New Query
  q.Transient = Transient
  Set rs = mPQRS.GetParsedQueryRS(db, QUERY_HEADER_RS)
  rs.FindFirst "Name='" & qd.name & "'"
  If Not rs.NoMatch Then
    If rs.fields("LastUpdated") = qd.LastUpdated Then
      q.LoadQueryName Me, qd.name
    Else
      q.setupquery Me, qd
    End If
  Else
    q.setupquery Me, qd
  End If
  q.linkchildren
  q.linkparents
  q.linkinrelations

CreateQueryQD_End:
  Set CreateQueryQD = q
  Call xReturn("CreateQueryQD")
  Exit Function

CreateQueryQD_err:
  Err.Raise ERR_CREATE_QUERY, ErrorSource(Err, "tcsda.CreateQueryQD"), ""
  Resume CreateQueryQD_End
  Resume
End Function

Private Function CreateQuerySQL(sql As String, Optional Transient As Boolean = True, Optional HasParents As Boolean = True) As Query
Dim q As Query
Dim qKid As Query

  On Error GoTo CreateQuerySQL_err
  Call xSet("CreateQuerySQL")
  
  Set q = New Query
  q.Transient = Transient
  q.setupquery Me, Nothing, sql, HasParents
  q.linkchildren
  q.linkparents
  q.linkinrelations

CreateQuerySQL_End:
  Set CreateQuerySQL = q
'  Set q = Nothing
  Call xReturn("CreateQuerySQL")
  Exit Function

CreateQuerySQL_err:
  Err.Raise ERR_CREATE_SQL, ErrorSource(Err, "tcsda.CreateQueryQD"), ""
  Resume CreateQuerySQL_End
  Resume
End Function
Friend Sub LogRelations(obj As Object, i As REL_TYPE)
Dim c As calc
  On Error Resume Next
  
  Select Case i
  Case LIST_REL
    If Not TypeOf obj Is ListView Then Exit Sub
  Case TREE_TOP_REL
    If Not TypeOf obj Is TreeView Then Exit Sub
  Case Else
    ECASE "Invalid rel_type in LogRelations for DADatabase"
  End Select
  If Not objqueries Is Nothing Then
    For Each c In objqueries
      c.LogRelations obj, i
    Next c
  End If
  If Not objRecordsets Is Nothing Then
    For Each c In objRecordsets
      c.LogRelations obj, i
    Next c
  End If
  Set c = Nothing
  Call DisplayMessageKill
End Sub
Friend Sub LogObjects(frm As Form)

Dim c As calc
Dim q As Query

Dim tlv As ListView
Dim tli As ListItem

  On Error Resume Next
  Set tlv = frm.lv
  If Not objqueries Is Nothing Then
    For Each c In objqueries
      Set q = c
      Set tli = tlv.ListItems.Add
      tli.text = mFileName
      Call tli.ListSubItems.Add(, , "Query")
      Call tli.ListSubItems.Add(, , c.name)
      Call tli.ListSubItems.Add(, , CStr(c.boocalced))
      Call tli.ListSubItems.Add(, , q.Modes)
    Next c
  End If
  
  If Not objRecordsets Is Nothing Then
    For Each c In objRecordsets
      Set tli = tlv.ListItems.Add
      tli.text = mFileName
      Call tli.ListSubItems.Add(, , "Recordset")
      Call tli.ListSubItems.Add(, , c.name)
      Call tli.ListSubItems.Add(, , CStr(c.boocalced))
      Call tli.ListSubItems.Add(, , q.Modes)
    Next c
  End If
  Set c = Nothing
  Set q = Nothing
  Set tlv = Nothing
  Set tli = Nothing
End Sub

Public Function RepairCompact() As Boolean
  
  On Error GoTo RepairCompact_Err
  Call xSet("RepairCompact")
  If mDbOpen Then
    ForceCloseDatabase
  End If
  If Not IsDatabaseOpen(mFileName) Then
    If Not RepairCompactDB(mFileName, MODE_COMPACT + MODE_REPAIR) Then Err.Raise ERR_DB_LOCKED, "RepairCompact", "The repair and compact of the database " & mFileName & " failed."
  Else
    logerr "The database " & mFileName & " is in use by another user."
  End If
  If mDbOpen Then
    OpenDatabase mFileName, mbMode
  End If
  
RepairCompact_End:
  Call xReturn("RepairCompact")
  Exit Function

RepairCompact_Err:
  Call ErrorMessage(ERR_ERROR, Err, "RepairCompact", "Error in RepairCompact", "Undefined error.")
  Resume RepairCompact_End
End Function


Private Function ForceCloseDatabase() As Boolean
  ' closes database without losing queries
  On Error GoTo ForceCloseDatabase_Err
  Call xSet("ForceCloseDatabase")
  If Not mrsControl Is Nothing Then mrsControl.Close
  If Not mrsCalcNo Is Nothing Then mrsCalcNo.Close
  Call mPQRS.Kill
  
  If Not db Is Nothing Then
    Call db.Close
    Set db = Nothing
  End If
  
  Set mrsControl = Nothing
  Set mrsCalcNo = Nothing

ForceCloseDatabase_End:
  Call xReturn("ForceCloseDatabase")
  Exit Function

ForceCloseDatabase_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ForceCloseDatabase", "Error in ForceCloseDatabase", "Undefined error.")
  Resume ForceCloseDatabase_End
  Resume
End Function


Public Function OpenDatabase(Filename As String, Optional bMode As Boolean = False) As Boolean
' opens database without loading queries
' used for repair etc which closes without removing
  On Error GoTo OpenDatabase_Err
  Call xSet("OpenDatabase")
  Call ShowProgress("Opening " & Filename, "Opening Database")
  
  If Not db Is Nothing Then
    Call db.Close
  End If
  Set db = Nothing
  
  'If (Not IsDatabaseOpen(mFileName)) And (Not mForceMode) Then
  '  Set db = InitDB(gwsMain, Filename, "DA Database", , True)
  'Else
    Set db = InitDB(gwsMain, Filename, "DA Database", , bMode)
  'End If
  
  Call DisplayMessageKill
  OpenDatabase = Not db Is Nothing
  
OpenDatabase_End:
  mForceMode = False
  Call xReturn("OpenDatabase")
  Exit Function

OpenDatabase_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenDatabase", "Error in OpenDatabase", "Undefined error.")
  Resume OpenDatabase_End
End Function

Public Function Recalc() As Boolean
  On Error GoTo Recalc_Err:
  
  Recalc = False
  If mrsControl.bof And mrsControl.eof Then GoTo Recalc_End
  
  If mCalcMethod = MANUAL_RECALC_WARN Then
    If MultiDialog("Manual Recalculation", "Are you sure you want to perform a manual recalculation?", "Continue", "Cancel") = 1 Then
      Recalc = RecalcEx
    End If
  Else
    Recalc = RecalcEx
  End If

Recalc_End:
  Exit Function
  
Recalc_Err:
  Call Err.Raise(ERR_RECALC, "Recalc", "An error occurred in the recalc", Err.HelpFile, Err.HelpContext)
  Recalc = False
End Function

Private Function RecalcEx() As Boolean
  Dim q As Query
  Dim c As calc
  
  On Error GoTo RecalcEx_Err
  
  Call SetCursor(vbHourglass)
  For Each q In objqueries
    Set c = q
    c.calc
  Next
  RecalcEx = True
  
RecalcEx_End:
  Call ClearCursor
  Exit Function

RecalcEx_Err:
  Call Err.Raise(ERR_RECALC_EX, "RecalcEx", "An error occurred in RecalcEx", Err.HelpFile, Err.HelpContext)
End Function

Public Property Get CalcMethod() As CALC_METHOD
  CalcMethod = mCalcMethod
End Property

Public Property Let CalcMethod(NewValue As CALC_METHOD)
  Let mCalcMethod = NewValue
  If mCalcMethod = AUTO_RECALC Then
    gCalc = True
  Else
    gCalc = False
  End If
End Property

    
