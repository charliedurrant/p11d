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

Public Enum RS_TYPE
  QUERY_HEADER_RS
  QUERY_DETAIL_RS
  QUERY_RELATIONS_RS
End Enum
Private mrsControl As recordset
Private mrsCalcNo As recordset
Private mrsDetail As recordset
Private mrsRelations As recordset
Private mrsQuery As recordset

Dim intlockcount As Long
Dim mbMode As Boolean             ' enhance with applies for multi packs

Dim strbatchexecute() As String   ' remove as replaced with collection of batchitems

Dim db As database
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
    OpenDatabase mFileName, NewMode
  End If

AccessMode_end:
  mbMode = NewMode
  Call xReturn("AccessMode")
  Exit Property

AccessMode_err:
  Err.Raise Err.Number, Err.Source, Err.Description
End Property

Friend Function OpenDB(Filename As String, idString As String) As Boolean
  Dim qd As QueryDef
  Dim td As TableDef
  Dim q As Query
  Dim dummy As Long
  Dim rsquery As recordset
  Dim i As Long
  Dim j As Long
  Const k = 0.01
  Dim TBOFix As TBOFixes
  
  Call xSet("OpenDb")
  
  On Error GoTo Opendb_err:
  If mDbOpen Then Err.Raise ERR_DBOPEN, "OpenDatabase", "The database '" & Filename & "' is already open."
  If Not OpenDatabase(Filename, mbMode) Then Exit Function
  Call ShowProgress("Checking " & Filename, "Opening Database", 2 * k)
  
  'If Not iFix Is Nothing Then Call iFix.ApplyDAFixes(db)
  Set TBOFix = New TBOFixes
  'Do we want to check return value of ApplyDAFixed??????
  Call TBOFix.ApplyDAFixes(db)
  mFileName = Filename
  gOpenDatabases.Add Me, Filename
  Call inclockcount
  mDbOpen = True
  Set mrsControl = db.OpenRecordset("SELECT * FROM sys_Control")
  Set mrsCalcNo = mrsControl.Clone
  mrsCalcNo.FindFirst "Flag=" & strsql("CALCNO")
  If mrsCalcNo.NoMatch Then
    mrsCalcNo.addnew
    mrsCalcNo!flag = "CALCNO"
    mrsCalcNo!state = 0
    mrsCalcNo.update
    mrsCalcNo.FindFirst "Flag=" & strsql("CALCNO")
  End If
  Call ShowProgress("Checking " & Filename, "Opening Database", 2 * k)
  
  intver = Me.flag("id", dummy, mProgID)
  ' zzzz reinstate If StrComp(idString, mProgID, vbTextCompare) <> 0 Then Err.Raise ERR_DBOPEN, "DBOpen", "The internal version and the current version do not match."
  
  
'need to upgrade version
Call xSet("SetUpTables")
  i = 1
  j = db.TableDefs.Count
  For Each td In db.TableDefs
    Call ShowProgress("Loading table " & CStr(i) & " of " & CStr(db.TableDefs.Count), "Opening Database", k * 10 / j)
    i = i + 1
    If StrComp(Left$(td.name, 4), "MSYS", vbTextCompare) <> 0 And StrComp(Left$(td.name, 4), "SYS_", vbTextCompare) <> 0 Then
      Set q = New Query
      Call q.setuptable(td, Me)
    End If
  Next td
Call xReturn("SetUpTables")
Call xSet("LoadQueries")
  
  Set mrsDetail = db.OpenRecordset("select * from sys_queries_detail")
  Set mrsRelations = db.OpenRecordset("select * from sys_queries_relations")
  Set rsquery = db.OpenRecordset("select * from sys_queries_for_Loading")
  i = 1
  If Not rsquery.eof Then
    rsquery.movelast
    rsquery.movefirst
    j = rsquery.RecordCount
    While Not rsquery.eof
      Call ShowProgress("Loading query " & CStr(i) & " of " & CStr(rsquery.RecordCount) & vbCr & rsquery.fields("Name"), "Opening Database", k * 65 / j)
      i = i + 1
      If rsquery.fields("Type") = 5 Then
        ' load select and non select queries
        Set qd = db.QueryDefs(rsquery.fields("Name"))
        Set q = New Query
        q.Transient = False
        Call q.LoadQuery(Me, qd)   ', ParseQD)
      End If
      rsquery.movenext
    Wend
  End If
Call xReturn("LoadQueries")
Call xSet("SetUpQueries")
  
  Set rsquery = db.OpenRecordset("select * from sys_queries_for_Updating")
  i = 1
  If Not rsquery.eof Then
    rsquery.movelast
    rsquery.movefirst
    j = rsquery.RecordCount
    While Not rsquery.eof
      Call ShowProgress("Updating query " & CStr(i) & " of " & CStr(rsquery.RecordCount) & vbCr & rsquery.fields("Name"), "Opening Database", k * 10 / j)
      i = i + 1
      If rsquery.fields("Type") = 5 Then
        ' a query
        Set qd = db.QueryDefs(rsquery.fields("Name"))
        If qd.Type <> dbQSelect Then
          ' only update non selects
          Set q = New Query
          q.Transient = False
          q.setupquery Me, qd
        End If
      End If
      rsquery.movenext
    Wend
  End If
Call xReturn("SetUpQueries")
  
  
Call ShowProgress("Establishing relations", "Opening Database", 2 * k)
Call xSet("LinkChildren")
  For Each q In objqueries
    q.linkchildren
  Next q
Call xReturn("LinkChildren")
Call ShowProgress("Establishing relations", "Opening Database", 2 * k)
Call xSet("LinkParents")
  For Each q In objqueries
    q.linkparents
  Next q
Call xReturn("LinkParents")
  
  dummy = Me.flag("MODE", dummy, mstrMode)
  
  OpenDB = True
  
Opendb_end:
  Set mrsDetail = Nothing
  Set mrsRelations = Nothing
  Set mrsQuery = Nothing
  Set qd = Nothing
  Call DisplayMessageKill
  Call ResetProgress
  Call xReturn("OpenDb")
  Exit Function
Opendb_err:
  OpenDB = False
  Call ErrorMessage(ERR_ERROR, Err, "tcsda.OpenDb", "Opening Database", "An error has occurred opening " & Filename)
  Resume Opendb_end
  Resume
  Err.Raise Err.Number, "OpenDB", "An error occurred opening the database '" & Filename & "'" & vbCrLf & Err.Description
End Function
Friend Function Recordsets() As Collection
  Set Recordsets = objRecordsets
End Function
Public Function RS_ofType(rs_t As RS_TYPE) As recordset
  On Error GoTo RS_ofType_ERR
  If db Is Nothing Then Exit Function
  xSet ("RS_ofType")
  Select Case rs_t
  Case QUERY_HEADER_RS
    Set RS_ofType = db.OpenRecordset("select * from sys_queries_header")
  Case QUERY_DETAIL_RS
    If mrsDetail Is Nothing Then
      Set RS_ofType = db.OpenRecordset("select * from sys_queries_detail")
    Else
      Set RS_ofType = mrsDetail
    End If
  Case QUERY_RELATIONS_RS
    If mrsRelations Is Nothing Then
      Set RS_ofType = db.OpenRecordset("select * from sys_queries_relations")
    Else
      Set RS_ofType = mrsRelations
    End If
  Case Else
    ECASE "Invalid rs_type - " & CStr(rs_t)
  End Select
RS_ofType_END:
  xReturn ("RS_ofType")
  Exit Function
RS_ofType_ERR:

  Resume RS_ofType_END
End Function
Public Function CloseDb(sappname As String) As Boolean
  Dim s As String
  
  On Error GoTo CloseDb_Err
  Call xSet("CloseDb")
  CloseDb = True
  If mDbOpen Or Not db Is Nothing Then
    If Not mLocks Is Nothing Then
      If InStrCollection(mLocks, sappname) Then
        mLocks.Remove sappname
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

CloseDb_End:
  Call xReturn("CloseDb")
  Exit Function

CloseDb_Err:
  CloseDb = False
  Call ErrorMessage(ERR_ERROR, Err, "CloseDb", "Error in CloseDb", "Undefined error.")
  Resume CloseDb_End
  Resume
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
  Set objqueries = Nothing
  Set objRecordsets = Nothing
  Set db = Nothing
  Call DisplayMessageKill
  Exit Sub
err_ForceCloseDb:
  logerr "Error in ForceCloseDb "
  Resume Next
End Sub

Private Sub Class_Initialize()
  mDbOpen = False
  ReDim strbatchexecute(0)
  Set mLocks = New Collection
  
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
    Set q = objqueries(sqlstring)
  Else
    If InCollection(db.QueryDefs, sqlstring) Then
      Set qd = db.QueryDefs(sqlstring)
      Set q = CreateQuery(qd, "", False)
      qd.Close
    Else
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
  flag = xLng(mrsControl!state)
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
  calcno = mrsCalcNo!state + 1
  mrsCalcNo!state = calcno
  
  mrsControl.FindFirst "Flag='" + UCase$(flagname) + "'"
  If mrsControl.NoMatch Then
    mrsControl.addnew
    mrsControl!flag = flagname
    mrsControl!state = val
    mrsControl!calcno = calcno
    If Not IsMissing(flagtext) Then mrsControl!text = flagtext
  Else
    mrsControl.edit
    mrsControl!state = val
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
  Dim bKillPop As Boolean
  Dim Batch As Collection
  Dim b As BatchItem
    
  ' batch execute will set boocalced for children to false, if SetCalcFlags is true
  ' if SetCalcFlags is flase then it will expect user to handle touching
  ' touching can be done through batchexecute
  ' it will not calc parents and will not derive parents
  ' zzzz batch execute should use transactions to allow undo all if fail (apart from multiuser worries)

  On Error GoTo Batchexecute_ERR
  xSet ("BatchExecute")
  If InStr("EXECUTE", UCase$(sqlstring)) = 1 Then
    bKillPop = True
    
    For Each b In Batch
      Call ShowProgress("Executing " & IIf(UBound(strbatchexecute) > 1, CStr(i) & " of " & CStr(UBound(strbatchexecute)) & " queries", " query"), "Executing")
      If b.calc Then
        If InCollection(queries, b.sql) Then
          Set q = queries(b.sql)
          q.Calc_calced(Nothing) = False
          Set c = q
          c.calc
          Set c = Nothing
          q.Kill
          Set q = Nothing
        Else
          Err.Raise ERR_SQL_PARSE, "BATCHEXECUTE", "Failed to locate " & vbCr & strbatchexecute(i) & vbCr & " in executables queries.  Will execute anyway, but will fail to set calc flags."
          db.execute b.sql
        End If
      Else
        db.execute b.sql
      End If
    Next b
    j = 0
        
    
  
  ElseIf InStr("TOUCH ", UCase$(sqlstring)) = 1 Then
    ' touch the object sqlstring
    Call touch(Trim$(Mid$(sqlstring, 6)))
  Else
    j = j + 1
    Call ShowProgress("Preparing executable " & CStr(j), "Executing")
    Set b = New BatchItem
    b.calc = SetCalcFlags
    b.sql = sqlstring
    If Batch Is Nothing Then Set Batch = New Collection
    Batch.Add b
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

  Select Case Err.Number
  Case 3012
    retry = retry + 1
    If retry < 10 Then
      Resume
    Else
      logerr "Error in database batchexecute " & sqlstring
      Resume Batchexecute_END
    End If
  Case Else
    logerr "Error in database batchexecute " & sqlstring
    Resume Next
    Resume
  End Select
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
      Set CreateQuery = CreateQuerySQL(sql, Transient, HasParents)
    End If
  Else
    Set CreateQuery = CreateQueryQD(qd, Transient)
  End If
  
End Function

Private Function CreateQueryQD(qd As QueryDef, Optional Transient As Boolean = True) As Query
Dim q As Query
Dim rs As recordset

  On Error GoTo CreateQueryQD_err
  Call xSet("CreateQueryQD")
  
  Set q = New Query
  q.Transient = Transient
  Set rs = RS_ofType(QUERY_HEADER_RS)
  rs.FindFirst "Name='" & qd.name & "'"
  If Not rs.NoMatch Then
    If rs.fields("DateUpdate") = qd.LastUpdated Then
      q.LoadQuery Me, qd
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
  Err.Raise Err.Number, ErrorSource(Err, "tcsda.CreateQueryQD"), ""
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
  Err.Raise Err.Number, ErrorSource(Err, "tcsda.CreateQueryQD"), ""
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
  If Not mrsDetail Is Nothing Then mrsDetail.Close
  If Not mrsRelations Is Nothing Then mrsRelations.Close
  If Not mrsQuery Is Nothing Then mrsQuery.Close
  
  If Not db Is Nothing Then db.Close
  Set mrsControl = Nothing
  Set mrsCalcNo = Nothing
  Set mrsDetail = Nothing
  Set mrsRelations = Nothing
  Set mrsQuery = Nothing

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
  
  Set db = InitDB(gwsMain, Filename, "DA Database", , bMode)
  
  Call DisplayMessageKill
  OpenDatabase = Not db Is Nothing
OpenDatabase_End:
  Call xReturn("OpenDatabase")
  Exit Function

OpenDatabase_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenDatabase", "Error in OpenDatabase", "Undefined error.")
  Resume OpenDatabase_End
End Function

