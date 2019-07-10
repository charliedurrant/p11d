VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ILibraryVersion

Public Enum COPYTABLE_MODE
  TBL_COPY = 8
  TBL_COPY_OVERWRITE = 16
  TBL_UPDATE = 32
  TBL_UPDATE_SYNC = 64
End Enum

Public Enum PROPERTIES_FILTER
  [_INVALID_PROP] = 0
  PROP_DEFAULTVALUE = 1
  PROP_ALLOWZEROLENGTH = 2
  PROP_REQUIRED = 4
  PROP_ALL = -1
End Enum

Public Enum TRANSACTION_TYPE
  TRANS_BEGIN = 1
  TRANS_ROLLBACK
  TRANS_COMMIT
End Enum

Public Enum REPAIRCOMPACT_MODE
  MODE_REPAIR = 2
  MODE_COMPACT = 4
End Enum

Public Enum DATABASEUPDATE_MODE
  DB_DELETE
  DB_ADD = TBL_UPDATE
  DB_SYNC = TBL_UPDATE_SYNC
End Enum

Public Enum DB_EVENT_TYPES
  IDLE_ONLY = 0
  FREE_LOCKS = dbFreeLocks
  REFRESH_CACHE = dbRefreshCache
  VB_DOEVENTS = 256
End Enum


Private Sub Class_Initialize()
  Call xSet("Core_DatabaseSetup")
  If mDBInitCount = 0 Then
    mDBInitCount = mDBInitCount + 1
    m_SystemMDWPath = AppPath & "\SYSTEM.MDW"
    If InitialiseDefaultWS Then
      If InitWorkSpace(m_SystemMDWPath, True) Then
        Call RegisterDB(Nothing)
      End If
    End If
    Call RegisterLibrary(Me)
  Else
    mDBInitCount = mDBInitCount + 1
  End If
  Call xReturn("Core_DatabaseSetup")
End Sub

Private Sub Class_Terminate()
  mDBInitCount = mDBInitCount - 1
  If mDBInitCount = 0 Then
    Set m_wsMain = Nothing
  End If
End Sub

Public Function SetWorkSpace(ByVal UserName As String, ByVal Password As String, Optional ByVal ShowErrors As Boolean = False, Optional ByVal SystemMDWPath As String = "") As Boolean
  If Len(Trim$(SystemMDWPath)) = 0 Then SystemMDWPath = m_SystemMDWPath
  SetWorkSpace = InitWorkSpace(SystemMDWPath, ShowErrors, UserName, Password)
End Function

' Setup workspace, if ShowErrors is true than show fatalmessage if fail
Private Function InitWorkSpace(ByVal SystemMDWPath As String, ByVal ShowErrors As Boolean, Optional ByVal UserName As String = "Admin", Optional ByVal Password As String = "") As Boolean
  Call xSet("InitWorkSpace")
  On Error GoTo InitWorkSpace_err
  Set m_wsMain = Nothing
  DBEngine.SystemDB = SystemMDWPath
  DBEngine.DefaultUser = UserName
  DBEngine.DefaultPassword = Password  ' was ukcentral
  Set m_wsMain = DBEngine.Workspaces(0)
  InitWorkSpace = True
  
InitWorkSpace_end:
  Call xReturn("InitWorkSpace")
  Exit Function

InitWorkSpace_err:
  If ShowErrors Then
    Call ErrorMessage(ERR_ERROR, Err, "InitWorkSpace", "ERR_INITWORKSPACE", "Failed to initialise database workspace." & vbCr & "SystemMDW: " & SystemMDWPath)
  End If
  InitWorkSpace = False
  Set m_wsMain = Nothing
  Resume InitWorkSpace_end
End Function

'* Global Workspace created at initialisation for you use for transactions etc..
'* notes:
'* SYSTEM.MDW is assumed to be in the Applications home directory
'* For multiple simultaneous transactions you will need multiple workspaces
'* see IsolateODBCTrans for description
Public Property Get gwsMain() As Workspace
  Set gwsMain = m_wsMain
End Property


'* Copies an entire database structure and data (if required)
'* it can also change the version of the database as it copies
'*
'* ws             workspace to use for copy will normally be gwsMain
'* Dest          name of destination database
'* Src           name of source database
'* CopyData       Flag to copy data
'* [dbVersion]    Version of database to create (defaults = Version 3.0)
'* [ShowErrors]   Show error messages (default = true)
'*
'* return value:
'* true if success, false otherwise
Public Function CopyDB(ws As Workspace, ByVal Dest As String, ByVal Src As String, Optional ByVal CopyData As Boolean, Optional ByVal dbVersion As Long = dbVersion30, Optional ByVal ShowErrors As Boolean = True) As Boolean
  Dim dbDest As Database, dbSrc As Database
  On Error GoTo CopyDB_err
  Call xSet("CopyDB")
  
  'open out db and create
  Set dbDest = ws.CreateDatabase(Dest, dbLangGeneral, dbVersion)
  Set dbSrc = ws.OpenDatabase(Src, True)
  CopyDB = UpdateDatabase(ws, dbSrc, dbDest, DB_ADD, True, CopyData, ShowErrors)

CopyDB_end:
  Call xReturn("CopyDB")
  Exit Function
  
CopyDB_err:
  If ShowErrors Then
    Call ErrorMessage(ERR_ERROR, Err, "CopyDB", "ERR_COPYING_DATABASE", "Error copying database " & Src & " to " & Dest)
  End If
  CopyDB = False
  Resume CopyDB_end
End Function

'* Copies from one table to another in the same database (assumed to be open)
'* Compatibility only use CopyTableEx instead in new code
Public Function CopyTable(ws As Workspace, Db As Database, ByVal DestTable As String, ByVal SrcTable As String, ByVal DeleteDest As Boolean, Optional ByVal SQL_Select As String = "", Optional ByVal ShowErrors As Boolean = True, Optional ByVal UseTrans As Boolean = True) As Boolean
  Dim mode As COPYTABLE_MODE
  mode = TBL_COPY
  If DeleteDest Then mode = TBL_COPY_OVERWRITE
  CopyTable = CopyTableEx(ws, Db, DestTable, Db, SrcTable, mode, SQL_Select, ShowErrors, UseTrans)
End Function

'* Synchronizes two tables, optionally removing fields from destination
'* Compatibility only use CopyTableEx instead in new code
Public Function UpdateTable(ws As Workspace, tdSource As TableDef, tdDest As TableDef, Optional ByVal bSync As Boolean = False) As Boolean
  Call ECASE("UpdateTable obsolete - call CopyTableEx")
End Function

'* Copies from one table to another in the same or different database (assumed to be open)
'* note: This function assumes that the workspace is gwsMain and should the function fail
'* then the transaction is rolled back
'*
'* During the process the cursor is changed to an hourglass
'* To copy all data use "SELECT * from " & sSrc (default = "", copy structure only)
'*
'* ws As Workspace
'* DestDb           Destination database
'* DestTable        Destination table name
'* SrcDb            Source database
'* SrcTable         Source table name
'* CopyMode         CopyMODE
'*                     TBL_COPY   Copies tabledef and data to new table
'*                                only if new table does not exist
'*                     TBL_COPY_OVERWRITE
'*                                Copies tabledef and data to new table,
'*                                deletes dest table if it exists
'*                     TBL_UPDATE Synchronizes two tables
'*                     TBL_UPDATE_SYNC
'*                                Synchronizes two tables removing fields from
'*                                destination
'*
'* [sSelect]        SQL statement used to decide what data to copy from table source
'* [bShowErrors]    Show error messages (default = true)
'* [bUseTrans]      On error roll back updates
'*
'* return value:
'* True if success, false otherwise
Public Function CopyTableEx(ws As Workspace, DestDb As Database, ByVal DestTable As String, SrcDb As Database, ByVal SrcTable As String, CopyMode As COPYTABLE_MODE, Optional ByVal SQL_Select As String = "", Optional ByVal ShowErrors As Boolean = True, Optional ByVal UseTrans As Boolean = True) As Boolean
  CopyTableEx = CopyTable2(ws, DestDb, DestTable, SrcDb, SrcTable, CopyMode, PROP_ALL, SQL_Select, ShowErrors, UseTrans)
End Function

Public Function CopyTable2(ws As Workspace, DestDb As Database, ByVal DestTable As String, SrcDb As Database, ByVal SrcTable As String, ByVal CopyMode As COPYTABLE_MODE, ByVal PropFilter As PROPERTIES_FILTER, Optional ByVal SQL_Select As String = "", Optional ByVal ShowErrors As Boolean = True, Optional ByVal UseTrans As Boolean = True) As Boolean
  Dim tdSrc As TableDef
  Dim tdDest As TableDef
  Dim fd As Field, fld As Field
  Dim bDestExists As Boolean, i As Long
  Dim sql As String
  Call xSet("CopyTable2")

  On Error GoTo CopyTable2_err
  Call SetCursor
  If Not (DestDb.Transactions And SrcDb.Transactions) Then UseTrans = False
  If UseTrans Then ws.BeginTrans

  If StrComp(SrcTable, DestTable, vbTextCompare) = 0 Then
    If DestDb.Name = SrcDb.Name Then Call Err.Raise(ERR_SAME_DEST, "CopyTable - ERR_SAME_DEST", "The source and destination tables are the same.")
  End If

  bDestExists = TablePresent(DestDb.TableDefs, DestTable)
  Set tdSrc = SrcDb.TableDefs(SrcTable)
  
  If bDestExists Then
    Set tdDest = DestDb.TableDefs(DestTable)
    'If (tdSrc.LastUpdated <> tdDest.LastUpdated) And ((tdDest.Attributes And dbAttachedTable) = 0) Then GoTo Copy_data
    
    If (CopyMode And TBL_COPY) <> 0 Then Call Err.Raise(ERR_TABLEEXISTS, "CopyTableEx", "The destination table, " & tdDest.Name & " already exists.")
    If (CopyMode And TBL_COPY_OVERWRITE) <> 0 Then
      Set tdDest = Nothing
      Call DestDb.TableDefs.Delete(DestTable)
      GoTo AddNewTableDef
    End If
    If (CopyMode And TBL_UPDATE_SYNC) <> 0 Then
      i = 0
      Do While i < tdDest.Fields.Count
        Set fd = tdDest.Fields(i)
        If Not FieldPresent(tdSrc.Fields, fd.Name) Then
          If Not KillField(tdDest, fd.Name) Then Call Err.Raise(ERR_KILL_FIELD, "UpdateTable", "Unable to remove the field " & fd.Name)
        Else
          i = i + 1
        End If
      Loop
      If Not CopyIndexes(Nothing, tdDest, True) Then Call Err.Raise(ERR_KILL_FIELD, "UpdateTable", "Unable to remove indices from the table " & DestTable & ".")
    End If
    If ((CopyMode And TBL_UPDATE) <> 0) Or ((CopyMode And TBL_UPDATE_SYNC) <> 0) Then
      If Not CopyTableDefEx(tdDest, tdSrc, CopyMode, PropFilter) Then Call Err.Raise(ERR_COPYTABLEDEF, "CopyTableEx", "Error copying the tabledef " & DestTable & ".")
    End If
  Else
AddNewTableDef:
    Set tdDest = DestDb.CreateTableDef
    If Not CopyTableDefEx(tdDest, tdSrc, CopyMode, PropFilter, True) Then Call Err.Raise(ERR_COPYTABLEDEF, "CopyTableEx", "Error copying the tabledef " & DestTable & ".")
    tdDest.Name = DestTable
    Call DestDb.TableDefs.Append(tdDest)
    Call DestDb.TableDefs.Refresh
  End If
  If UseTrans Then ws.CommitTrans
  
  ' Copy over user defined properties
  If UseTrans Then ws.BeginTrans
  
  Set tdDest = DestDb.TableDefs(DestTable)
  If Not CopyTableDefEx(tdDest, tdSrc, CopyMode, PropFilter, False, True) Then Call Err.Raise(ERR_COPYTABLEDEF, "CopyTableEx", "Error copying the tabledef " & DestTable & ".")
  
  If UseTrans Then ws.CommitTrans
Copy_data:
  If UseTrans Then ws.BeginTrans
    'copy data if sSelect has string and not an attached table
    If (Len(SQL_Select) > 0) And ((tdDest.Attributes And dbAttachedTable) = 0) Then
      If DestDb.Name = SrcDb.Name Then
        sql = "INSERT INTO " & tdDest.Name & " " & SQL_Select
      Else
        sql = "INSERT INTO " & tdDest.Name & " IN " & StrSQL(DestDb.Name) & " " & SQL_Select
      End If
      SrcDb.Execute sql, dbFailOnError
    End If
  If UseTrans Then ws.CommitTrans
  CopyTable2 = True

CopyTable2_end:
  Set tdSrc = Nothing
  Set tdDest = Nothing
  Call DBEngine.Idle(dbFreeLocks)
  Call ClearCursor
  Call xReturn("CopyTableEx")
  Exit Function

CopyTable2_err:
  If ShowErrors Then
    Call ErrorMessage(ERR_ERROR, Err, "CopyTable", "Copy Table", "Error copying the Table " & SrcTable)
  End If
  If UseTrans Then ws.Rollback
  CopyTable2 = False
  Resume CopyTable2_end
  Resume
End Function


'* Synchronizes two databases, optionally removing fields/tables/queries from destination
'*
'* ws               workspace
'* dbSource         Source database
'* dbDest           Destination database
'* mode             Update mode
'*
'* return value:
'* returns false on faliure
Public Function UpdateDatabase(ws As Workspace, SrcDb As Database, DestDb As Database, ByVal Updatemode As DATABASEUPDATE_MODE, Optional ByVal UseTrans As Boolean = True, Optional ByVal CopyData As Boolean = False, Optional ByVal ShowErrors As Boolean = True) As Boolean
  UpdateDatabase = UpdateDatabaseEx(ws, SrcDb, DestDb, Updatemode, PROP_ALL, UseTrans, CopyData, ShowErrors)
End Function

Public Function UpdateDatabaseEx(ws As Workspace, SrcDb As Database, DestDb As Database, ByVal Updatemode As DATABASEUPDATE_MODE, ByVal PropFilter As PROPERTIES_FILTER, Optional ByVal UseTrans As Boolean = True, Optional ByVal CopyData As Boolean = False, Optional ByVal ShowErrors As Boolean = True) As Boolean
  Dim sql As String
  Dim td As TableDef, qd As QueryDef
  Dim i As Long
  
  On Error GoTo UpdateDatabaseEx_err
  Call xSet("UpdateDatabaseEx")
  Call SetCursor
  Call DBEngine.Idle(dbFreeLocks + dbRefreshCache)
  If Not (DestDb.Transactions And SrcDb.Transactions) Then UseTrans = False
  
  If UseTrans Then ws.BeginTrans
  If ((Updatemode And DB_SYNC) <> 0) Or ((Updatemode And DB_DELETE) <> 0) Then
    ' deletes all the unwanted tables and queries in the destination database
    i = 0
    Do While i < DestDb.TableDefs.Count
      Set td = DestDb.TableDefs(i)
      If Not TablePresent(SrcDb.TableDefs, td.Name) Then
        If Not KillTable(DestDb, td.Name) Then Err.Raise ERR_KILL_TABLE, "UpdateDatabase", "Unable to delete the table " & td.Name
      Else
        i = i + 1
      End If
    Loop
    i = 0
    Do While i < DestDb.QueryDefs.Count
      Set qd = DestDb.QueryDefs(i)
      If Not QueryPresent(SrcDb.QueryDefs, qd.Name) Then
        If Not KillQuery(DestDb, qd.Name) Then Err.Raise ERR_KILL_QUERY, "UpdateDatabase", "Unable to delete the query " & qd.Name
      Else
        i = i + 1
      End If
    Loop
  End If
  If UseTrans Then ws.CommitTrans
  
  If ((Updatemode And DB_SYNC) <> 0) Or ((Updatemode And DB_ADD) <> 0) Then
    ' copies over all Tabledefs as required
    For Each td In SrcDb.TableDefs
      If Not IsSysTable(td) Then
        If CopyData Then
          sql = "SELECT * FROM " & td.Name
        Else
          sql = ""
        End If
        Call CopyTable2(ws, DestDb, td.Name, SrcDb, td.Name, Updatemode, PropFilter, sql, True, UseTrans)
      End If
    Next td
    If UseTrans Then ws.BeginTrans
      For Each qd In SrcDb.QueryDefs
        If Not AddNewQuery(DestDb, qd, True) Then Err.Raise ERR_COPY_DB, "UpdateDatabase", "Unable to update the query " & qd.Name
      Next qd
    If UseTrans Then ws.CommitTrans
  End If
  UpdateDatabaseEx = True
  
UpdateDatabaseEx_end:
  Set td = Nothing
  Set qd = Nothing
  Call DBEngine.Idle(dbFreeLocks)
  Call ClearCursor
  Call xReturn("UpdateDatabaseEx")
  Exit Function
  
UpdateDatabaseEx_err:
  UpdateDatabaseEx = False
  If UseTrans Then ws.Rollback
  If ShowErrors Then Call ErrorMessage(ERR_ERROR, Err, "UpdateDatabaseEx", "ERR_UPDATE_DB", "An error has occured updating " & DestDb.Name & " to " & SrcDb.Name & vbCrLf & "This is a serious problem because required tables, fields, queries and indexes may be missing." & vbCrLf & "The database has been returned to its previous state and no data loss has occured.")
  Resume UpdateDatabaseEx_end:
End Function


'* Given a grid and a recordset, this function labels the first row
'* with the field names and then appends each record from a
'* recordset to the grid.
'* The columns all have default to equal width, If you need customised
'* widths then adjust it after the call to this function
'*
'* FlexGridObj    The FlexGrid Object to Fill
'* rs             The recordset to fill the grid with
'*
'* Return values:
'*                none
Public Sub FlexGridFill(FlexGridObj As Object, rs As Recordset)
  Dim grd As MSFlexGrid
  Dim s As String
  Dim i As Long, iNumFields As Long
        
  Call xSet("FlexGridFill")

  If Not (TypeOf FlexGridObj Is MSFlexGrid) Then
    Err.Raise ERR_INVALIDOBJECT, "FlexGridFill", "The object is not a valid type."
    GoTo FlexGridFill_end
  Else
    Set grd = FlexGridObj
  End If

On Error GoTo FlexGridFill_err
  
  grd.Redraw = False
  
  iNumFields = rs.Fields.Count
  grd.Rows = 1
  grd.Cols = iNumFields
  grd.Row = 0
  
  'fill in first row with field names
  For i = 0 To iNumFields - 1
    grd.Col = i
    grd.Text = rs.Fields(i).Name
  Next i
 
  If rs.EOF And rs.BOF Then GoTo FlexGridFill_end:
  rs.MoveFirst
  Do While Not rs.EOF
    s = ""
    For i = 0 To iNumFields - 1
      s = s & rs.Fields(i) & vbTab
    Next i
    grd.AddItem s
    rs.MoveNext
  Loop
  
  grd.Redraw = True
FlexGridFill_end:
  grd.Refresh
  Call xReturn("FlexGridFill")
  Exit Sub
    
FlexGridFill_err:
  Call ErrorMessage(ERR_ERROR, Err, "FlexGridFill", "ERR_UNDEFINED", "Error filling the flexgrid")
  grd.Redraw = True
  Resume FlexGridFill_end
  
End Sub

'* Given a Listbox and a recordset, this function take a number of fields from
'* the recordset and adds them into the listbox
'*
'* ListBoxObj         Listbox to fill
'* rs                 Recordset to use as source
'* sFld               Array of fields to add to the list box, by name
'*
'* return value:
'* none
Public Sub ListBoxFill(ListObj As Object, rsOrg As Recordset, ParamArray FieldNames())
  Dim rs As Recordset
  Dim lmin As Long, lmax As Long
  Dim sText As String
  Dim fldFirst As Field, fldLast As Field
  Dim i As Long
  
  On Error GoTo ListBoxFill_err
  Call xSet("ListBoxFill")
  If Not (TypeOf ListObj Is ListBox) Then
    If Not (TypeOf ListObj Is ComboBox) Then
      Call ECASE("ListBoxFill - ListBoxFillObj not of type ListBox")
      GoTo ListBoxFill_end
    End If
  End If
  
  Set rs = rsOrg.Clone
  lmin = LBound(FieldNames)
  lmax = UBound(FieldNames)
  'bookmark current record pointer
  If (rs.Bookmarkable = False) Or ((lmax - lmin) < 0) Then
    Call ECASE("ListBoxFillEx - Bookmark or Field list error")
  Else
    'Clear current list then add new items from recordset
    ListObj.Clear
    Set fldFirst = rs.Fields(FieldNames(lmin))
    Set fldLast = rs.Fields(FieldNames(lmax))
    If Not (rs.BOF And rs.EOF) Then
      rs.MoveFirst
      Do
        sText = xStrPad(fldFirst.Value, " ", GLOBAL_LBCOLCHARS)
        If lmin < lmax Then
          sText = sText & vbTab
          For i = (lmin + 1) To (lmax - 1)
            sText = sText & xStrPad(rs.Fields(FieldNames(i)).Value, " ", GLOBAL_LBCOLCHARS) & vbTab
          Next i
          sText = sText & xStrPad(fldLast.Value, " ", GLOBAL_LBCOLCHARS)
        End If
        Call ListObj.AddItem(sText)
        rs.MoveNext
      Loop Until rs.EOF
    End If
  End If
  
ListBoxFill_end:
  Call xReturn("ListBoxFill")
  Exit Sub
    
ListBoxFill_err:
  Call ErrorMessage(ERR_ERROR, Err, "ListBoxFill", "Error Filling ListBox", "Error filling the listbox")
  Resume ListBoxFill_end
End Sub

'Returns true or false
'This function searches through RecordSet and finds the first occurrence of sSrc within sFld of the rs.
'bCurrent is whether to search from current record pointer, skipleading is to skip leading characters such as zeroes.
'* This function searches through RecordSet and finds the first occurrence of sSrc within sFld of the rs.
'*
'*
'* var1           desc..
'* [optvar]       desc..
'*
'* return value:
'* True if success, false otherwise
Function SearchRS(rs As Recordset, ByVal FieldName As String, ByVal SearchFor As String, Optional ByVal CurrentRS As Boolean = True) As Boolean
  Dim criteria As String
  Dim bmkLast As String
  Dim ilen As Long, i As Long
    
  On Error GoTo SearchRS_Err
  Call xSet("SearchRS")
  ilen = Len(SearchFor)
  If ilen = 0 Then GoTo SearchRS_end
  
  If Not rs.Bookmarkable Then Call Err.Raise(ERR_NOTBOOKMARKABLE, "SearchRS", "Recordset " & rs.Name & " does not support Bookmarks")
  bmkLast = rs.Bookmark
  
  If CurrentRS Then 'Search from current record pointer
    If StrComp(Left$(rs.Fields(FieldName).Value, ilen), SearchFor, vbTextCompare) <> 0 Then
      criteria = FieldName & " LIKE " & StrSQL(SearchFor & "*")
      rs.FindNext (criteria)
    End If
  Else
    i = 1
    criteria = FieldName & " LIKE " & StrSQL(Left$(SearchFor, i) & "*")
    rs.FindFirst (criteria)
    Do While (Not rs.NoMatch) And (i < ilen)
      If StrComp(Left$(rs.Fields(FieldName).Value, ilen), SearchFor, vbTextCompare) = 0 Then
        Exit Do
      End If
      bmkLast = rs.Bookmark
      i = i + 1
      criteria = FieldName & " LIKE " & StrSQL(Left$(SearchFor, i) & "*")
      rs.FindNext (criteria)
    Loop
  End If
  
  If rs.NoMatch Then
    rs.Bookmark = bmkLast
    SearchRS = False
  Else
    SearchRS = True
  End If
  
SearchRS_end:
  Call xReturn("SearchRS")
  Exit Function
  
SearchRS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SearchRS", "ERR_SEARCHRS", "Error searching recordset")
  SearchRS = False
  Resume SearchRS_end
End Function

'* Open a specified database within the workspace passed.
'* Optionall opens a recordset given the sql string to check for existance of tables / fields etc
'*
'* ws             Workspace to open database in (normally = gwsMain)
'* sDBPath        Full path to database including filename
'* sDBName        Description of database for error messages
'* [sql]          Open and close a specified recordset (error if does not exist)
'* [dbMode]       Open database exclusively
'*
'* return value:
'* open database if success, Nothing otherwise
Public Function InitDB(ws As Workspace, ByVal DatabasePath As String, ByVal FriendlyDatabaseName As String, Optional ByVal sql As String = "", Optional ByVal dbMode As Boolean = False, Optional ByVal ShowErrors As Boolean = True) As Database
  Dim rs As Recordset
  Dim Db As Database
  Dim retrycount As Long
  
  Call xSet("InitDB")
  On Error GoTo InitDB_err
  retrycount = 0
  
reopen:
  retrycount = retrycount + 1
  'apf dbMode is a variant Options
  Set Db = ws.OpenDatabase(DatabasePath, dbMode)
  If Len(sql) Then
    Set rs = Db.OpenRecordset(sql, dbOpenDynaset)
    Set rs = Nothing
  End If
  Set InitDB = Db
  
InitDB_end:
  Call xReturn("InitDB")
  Exit Function
  
InitDB_err:
   If ShowErrors Then
     Select Case Err.Number
      Case 3029 To 3031 'incorrect logins to open db
        Call ErrorMessage(ERR_ERROR, Err, "InitDB", "ERR_DB_LOGIN", "There was an error opening the database " & DatabasePath)
      Case 3049 'corrupt db
        If retrycount < 2 Then 'Tried to reopen db and failed again
          Call ErrorMessage(ERR_ERROR, Err, "InitDB", "ERR_CORRUPT_DB", "The database " & DatabasePath & " is corrupt.")
          DBEngine.RepairDatabase DatabasePath
          Resume reopen
        Else 'Fatal error repair on a db has failed to fix
          Call ErrorMessage(ERR_ERROR, Err, "InitDB", "ERR_REPAIR_FAILED", "The repair on the " & DatabasePath & " has failed")
        End If
      Case 3024  'Workspace is invalid
        Call ErrorMessage(ERR_ERROR, Err, "InitDB", "ERR_DB_WORKSPACE", "There was an error opening the database " & DatabasePath & vbLf & "The workspace is invalid or the database does not exist.")
      Case 3051 ' Database is read-only/not enough permissions to access file
        Call ErrorMessage(ERR_ERROR, Err, "InitDB", "ERR_DBOPEN", "There was an error opening the database " & DatabasePath & vbLf & "The file is marked as read only or you do not have sufficient rights to access the file.")
      Case 3261
        Call ErrorMessage(ERR_ERROR, Err, "InitDB", "ERR_TABLE_LOCKED", "There was an error validating the database " & DatabasePath & vbLf & Err.Description)
      Case Else
        If ErrorMessage(ERR_ERROR + ERR_ALLOWRETRY, Err, "InitDB", "Opening database", "An error occured opening the " & DatabasePath & ".") Then Resume reopen
    End Select
  End If
  Resume InitDB_end
End Function



'* Returns true if the table in the database
'* db               Database to check in
'* sTable           name of table
'*
'* return value:
'* True if present, false otherwise
Public Function IsTablePresent(Db As Database, ByVal TableName As String) As Boolean
  Call xSet("IsTablePresent")
  IsTablePresent = InCollection(Db.TableDefs, TableName)
  Call xReturn("IsTablePresent")
End Function

Public Function TablePresent(cTables As TableDefs, ByVal TableName As String) As Boolean
  Call xSet("TablePresent")
  TablePresent = InCollection(cTables, TableName)
  Call xReturn("TablePresent")
End Function

'* Returns true if the query in the database
'* db               Database to check in
'* sQuery           name of query
'*
'* return value:
'* True if present, false otherwise
Public Function IsQueryPresent(Db As Database, ByVal QueryName As String) As Boolean
  Call xSet("IsQueryPresent")
  IsQueryPresent = InCollection(Db.QueryDefs, QueryName)
  Call xReturn("IsQueryPresent")
End Function

Public Function QueryPresent(cQueries As QueryDefs, ByVal QueryName As String) As Boolean
  Call xSet("QueryPresent")
  QueryPresent = InCollection(cQueries, QueryName)
  Call xReturn("QueryPresent")
End Function

'* Returns true if the field is in the table def
'* td               Table def to check in
'* sField           name of field
'*
'* return value:
'* True if present, false otherwise
Public Function IsFieldPresent(td As TableDef, ByVal FieldName As String) As Boolean
  Call xSet("IsFieldPresent")
  IsFieldPresent = InCollection(td.Fields, FieldName)
  Call xReturn("IsFieldPresent")
End Function

Public Function FieldPresentRDO(cColumns As rdoColumns, ByVal ColumnName As String) As Boolean
  Dim v As Variant
  On Error Resume Next
  Call xSet("FieldPresentRDO")
  Set v = cColumns.Item(ColumnName)
  FieldPresentRDO = Not (v Is Nothing)
  Call xReturn("FieldPresentRDO")
End Function

Public Function FieldPresent(cFields As Fields, ByVal FieldName As String) As Boolean
  Call xSet("FieldPresent")
  FieldPresent = InCollection(cFields, FieldName)
  Call xReturn("FieldPresent")
End Function

'* Returns true if the index is in the table def
'* td               table def to check in
'* sIndex           name of index
'*
'* return value:
'* True if present, false otherwise
Public Function IsIndexPresent(td As TableDef, ByVal IndexName As String) As Boolean
  Call xSet("IsIndexPresent")
  IsIndexPresent = InCollection(td.Indexes, IndexName)
  Call xReturn("IsIndexPresent")
End Function

Public Function IndexPresent(cIndexes As Indexes, ByVal IndexName As String) As Boolean
  Call xSet("IndexPresent")
  IndexPresent = InCollection(cIndexes, IndexName)
  Call xReturn("IndexPresent")
End Function



'* Adds a new field to a Tabledef, No error messages raised
'*
'* td             TableDef object of Table to add field to
'* fieldname      name of field to add
'* fieldtype      type of field to add (choose from list)
'* [fieldsize]    size of field if applicable, set for dbText etc. (default=0)
'*
'* return value:
'* returns true if field added successfully or if field with same name, type
'* and size was already present, false otherwise
' note : FieldAttributeEnum  is part of DAO
Public Function AddNewField(td As TableDef, ByVal FieldName As String, ByVal FieldType As DataTypeEnum, Optional ByVal FieldSize As Integer = 0, Optional ByVal DefaultValue As Variant, Optional ByVal DAO_Attributes As FieldAttributeEnum = 0) As Boolean
  Dim fld As Field
  Dim i As Long

  Call xSet("AddNewField")
  On Error Resume Next
  Set fld = td.Fields(FieldName)
  If Not (fld Is Nothing) Then
    AddNewField = False
    If fld.Type = FieldType Then
      If IsMissing(DefaultValue) Then DefaultValue = ""
      Select Case fld.Type
        Case dbLong, dbNumeric, dbSingle, dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, dbFloat, dbInteger
          If fld.DefaultValue = DefaultValue Then AddNewField = True
        Case dbChar, dbBoolean
          If fld.DefaultValue = DefaultValue Then AddNewField = True
        Case dbBinary, dbText, dbVarBinary
          If (fld.Size = FieldSize) And (fld.DefaultValue = DefaultValue) Then AddNewField = True
        Case dbTime, dbTimeStamp, dbDate
          If fld.DefaultValue = DefaultValue Then AddNewField = True
        'Case dbLongBinary, dbMemo
        'If fld.DefaultValue = DefaultValue Then AddNewField = True
        Case Else
          Call ECASE("Unrecognised field type")
      End Select
    End If
    GoTo addnewfield_end
  End If
  
  On Error GoTo addnewfield_err
  If FieldSize = 0 Then
    Set fld = td.CreateField(FieldName, FieldType)
  Else
    Set fld = td.CreateField(FieldName, FieldType, FieldSize)
  End If
  If Not IsMissing(DefaultValue) Then fld.DefaultValue = DefaultValue
  If DAO_Attributes > 0 Then fld.Attributes = DAO_Attributes
  td.Fields.Append fld
  AddNewField = True
  
addnewfield_end:
  Call xReturn("AddNewField")
  Exit Function
  
addnewfield_err:
  AddNewField = False
  Resume addnewfield_end
End Function


'* Remove a new field from a Tabledef, No error messages raised
'*
'* td             TableDef object of Table to add field to
'* fieldname      name of field to add
'*
'* return value:
'* returns true if field removed successfully or field did not exist
Public Function KillField(td As TableDef, ByVal FieldName As String) As Boolean
  Dim fld As Field
  Dim i As Integer
  On Error Resume Next
  Call xSet("KillField")
  Set fld = td.Fields(FieldName)
  If fld Is Nothing Then
    KillField = True
    GoTo KillField_end
  End If
  On Error GoTo KillField_err
    
  Call KillIndexBeforeFieldDelete(td, FieldName)

  td.Fields.Delete FieldName
  KillField = True

KillField_end:
  Call xReturn("KillField")
  Exit Function

KillField_err:
  KillField = False
  Resume KillField_end
  Resume
End Function
Private Sub KillIndexBeforeFieldDelete(td As TableDef, ByVal FieldName As String)
'* Remove any indexes from a field before it is deleted
'*
'* td             TableDef object of Table to add field to
'* fieldname      name of field to find indexes on
'*
'* return value:
'* returns true if field removed successfully or field did not exist
  Dim i As Integer
  Dim sFieldName As String
  
  Call xSet("KillIndexBeforeFieldDelete")
  
  On Error GoTo KillIndex_err
  
  For i = 0 To td.Indexes.Count - 1
    sFieldName = Replace(td.Indexes(i).Fields, "+", "")
    If StrComp(sFieldName, FieldName, vbBinaryCompare) = 0 Then
      If Not KillIndex(td, td.Indexes(i).Name) Then Call Err.Raise(ERR_KILL_INDEX, "UpdateTable", "Unable to delete index" & td.Indexes(i).Name & " from field " & sFieldName)
      Exit For
    End If
  Next i
  

KillIndex_end:
  Call xReturn("KillIndexBeforeFieldDelete")
  Exit Sub

KillIndex_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "KillIndexBeforeFieldDelete"), Err.Description)
  
End Sub



'* Adds a new field to an index, No error messages raised
'*
'* idx            index to add field to
'* fieldname      name of field to add
'*
'* return value:
'* returns true if field added successfully
Public Function AddNewIndex(td As TableDef, ByVal Required As Boolean, ByVal Unique As Boolean, ByVal FieldName As String) As Boolean
  Dim fld As Field
  Dim idx As Index
  
  On Error GoTo AddNewIndex_err
  Call xSet("AddNewIndex")
  Set idx = td.CreateIndex(FieldName)
  idx.Required = Required
  idx.Unique = Unique
  idx.Fields.Append idx.CreateField(FieldName)
  Call td.Indexes.Append(idx)
  AddNewIndex = True
  
AddNewIndex_end:
  Call xReturn("AddNewIndex")
  Exit Function
  
AddNewIndex_err:
  AddNewIndex = False
  Resume AddNewIndex_end
End Function

Public Function AddNewPrimaryKey(td As TableDef, ParamArray FieldNames()) As Boolean
  Dim fld As Field, lmin As Long, lmax As Long, i As Long
  Dim idx As Index
  
  On Error GoTo AddNewPrimaryKey_err
  Call xSet("AddNewPrimaryKey")
  lmin = LBound(FieldNames)
  lmax = UBound(FieldNames)
  If lmin <= lmax Then
    ' kill indices on the above fields
    For i = lmin To lmax
      Call KillIndex(td, FieldNames(i))
    Next i
    Set idx = td.CreateIndex(FieldNames(lmin))
    idx.Primary = True
    For i = lmin To lmax
      idx.Fields.Append idx.CreateField(FieldNames(i))
    Next i
    Call td.Indexes.Append(idx)
    AddNewPrimaryKey = True
  End If
  
AddNewPrimaryKey_end:
  Call xReturn("AddNewPrimaryKey")
  Exit Function
  
AddNewPrimaryKey_err:
  AddNewPrimaryKey = False
  Resume AddNewPrimaryKey_end
End Function


'* Remove an index from a Tabledef, No error messages raised
'*
'* td             TableDef object of Table to add field to
'* indexname      name of field to add
'*
'* return value:
'* returns true if field removed successfully or field did not exist
Public Function KillIndex(td As TableDef, ByVal IndexName As String) As Boolean
  Dim idx As Index
  
  On Error Resume Next
  Call xSet("KillIndex")
  Set idx = td.Indexes(IndexName)
  If idx Is Nothing Then
    KillIndex = True
    GoTo KillIndex_end
  End If
  On Error GoTo KillIndex_err
  td.Indexes.Delete IndexName
  KillIndex = True
  
KillIndex_end:
  Call xReturn("KillIndex")
  Exit Function

KillIndex_err:
  KillIndex = False
  Resume KillIndex_end
End Function


'* Changes the name of a field in a Table
'*
'* td               TableDef object of Table to add field to
'* fromfield        Name of old fieldname
'* tofield          Name of new fieldname
'* [bShowErrors]    Flag to show error messages (default = true)
'*
'* return value:
'* True if success, false otherwise
Public Function ChangeField(td As TableDef, ByVal FromField As String, ByVal Tofield As String, Optional ByVal ShowErrors As Boolean = True) As Boolean
  
  Call xSet("ChangeField")
  On Error GoTo changefield_err
  
  td.Fields(FromField).Name = Tofield
  ChangeField = True
  
changefield_end:
  Call xReturn("ChangeField")
  Exit Function
  
changefield_err:
  If ShowErrors Then Call ErrorMessage(ERR_ERROR, Err, "ChangeField", "Change Field error", "Error changing the fieldname " & FromField)
  ChangeField = False
  Resume changefield_end
  
End Function
'* Adds a new query to the database or optionally updates an existing query
'*
'* db             Database object to add new query to
'* Query          querydef of the query to add
'* [bUpdate]      Update the query to the new one if it exists
'* return value:
'* returns true if query added or updated
Public Function AddNewQuery(Db As Database, Srcqd As QueryDef, Optional ByVal Update As Boolean = False) As Boolean
  Dim qd As QueryDef
  
  On Error GoTo AddNewQuery_err
  Call xSet("AddNewQuery")
  If QueryPresent(Db.QueryDefs, Srcqd.Name) Then
    If Update Then
      Set qd = Db.QueryDefs(Srcqd.Name)
      'If Srcqd.LastUpdated <> qd.LastUpdated Then qd.sql = Srcqd.sql
      qd.sql = Srcqd.sql
      AddNewQuery = True
    Else
      AddNewQuery = False
    End If
  Else
    Call Db.CreateQueryDef(Srcqd.Name, Srcqd.sql)
    AddNewQuery = True
  End If
  Db.QueryDefs.Refresh
  
AddNewQuery_end:
  Set qd = Nothing
  Call xReturn("AddNewQuery")
  Exit Function
  
AddNewQuery_err:
  AddNewQuery = False
  Resume AddNewQuery_end
End Function

'* Deletes a query from the database
'*
'* db             Database object to add delete query from
'* Query          querydef of the query to delete
'*
'* return value:
'* returns true if query removed or did not exist
Public Function KillQuery(Db As Database, ByVal QueryName As String) As Boolean
  On Error GoTo KillQuery_err
  xSet ("KillQuery")
  If QueryPresent(Db.QueryDefs, QueryName) Then
    Call Db.QueryDefs.Delete(QueryName)
  End If
  KillQuery = True
KillQuery_end:
  xReturn ("KillQuery")
  Exit Function
KillQuery_err:
    KillQuery = False
    Resume KillQuery_end
End Function

'* Adds a new table to the database. A new table must have at least one field.
'*
'* db             Database object to add new table to
'* tablename      Name of table to add
'* fieldname      name of field to add
'* fieldtype      type of field to add (choose from list)
'* [fieldsize]    size of field if applicable, set for dbText etc. (default=0)
'*
'* return value:
'* returns TabelDef of new table if success, Nothing otherwise
Public Function AddNewTable(Db As Database, ByVal TableName As String, ByVal FieldName As String, ByVal FieldType As DataTypeEnum, Optional ByVal FieldSize As Integer = 0) As TableDef
  Dim td As TableDef

  Call xSet("AddNewTable")
  On Error GoTo addnewtable_err
    
  Set td = Db.CreateTableDef(TableName)
  Call AddNewField(td, FieldName, FieldType, FieldSize)
  Db.TableDefs.Append td
  Set AddNewTable = Db.TableDefs(TableName)
  
addnewtable_end:
  Set td = Nothing
  Call xReturn("AddNewTable")
  Exit Function
  
addnewtable_err:
  Set AddNewTable = Nothing
  Resume addnewtable_end
End Function

'* Delete a table from a database.
'*
'* db             Database object from which the table will be removed
'* tablename      Name of table to be removed
'*
'* return value:
'* Returns true if Table removed or did not exist in database
Public Function KillTable(Db As Database, ByVal TableName As String) As Boolean
  
  On Error GoTo KillTable_err
  Call xSet("KillTable")
  If TablePresent(Db.TableDefs, TableName) Then
    Db.TableDefs.Delete TableName
  End If
  KillTable = True
  
KillTable_end:
  Call xReturn("KillTable")
  Exit Function
  
KillTable_err:
  KillTable = False
  Resume KillTable_end
End Function

Public Function IsSysTable(td As TableDef) As Boolean
  IsSysTable = IsSysTableEx(td)
End Function

Public Function ConnectAllFilterPrefix(DestDb As Database, SourceDb As Database, ByVal CopyQueries As Boolean, Optional ByVal TableFilter As String = "", Optional ByVal TablePrefix As String = "") As Boolean
  Call xSet("ConnectAllFilterPrefix")
  ConnectAllFilterPrefix = ConnectAllEx(DestDb, SourceDb, CopyQueries, TablePrefix, TableFilter)
  Call xReturn("ConnectAllFilterPrefix")
End Function

Public Function ConnectAllPrefix(DestDb As Database, SourceDb As Database, ByVal CopyQueries As Boolean, Optional ByVal TablePrefix As String = "") As Boolean
  Call xSet("ConnectAllPrefix")
  ConnectAllPrefix = ConnectAllEx(DestDb, SourceDb, CopyQueries, TablePrefix, "")
  Call xReturn("ConnectAllPrefix")
End Function

Public Function ConnectAll(DestDb As Database, SourceDb As Database, ByVal CopyQueries As Boolean) As Boolean
  Call xSet("ConnectAll")
  ConnectAll = ConnectAllEx(DestDb, SourceDb, CopyQueries, "", "")
  Call xReturn("ConnectAll")
End Function

Public Function RemoveLinkedTablesPrefix(Db As Database, Optional ByVal TablePrefix As String = "") As Boolean
  Call xSet("RemoveLinkedTablesPrefix")
  RemoveLinkedTablesPrefix = RemoveLinkedTablesEx(Db, TablePrefix)
  Call xReturn("RemoveLinkedTablesPrefix")
End Function

Public Function RemoveLinkedTables(Db As Database) As Boolean
  Call xSet("RemoveLinkedTables")
  RemoveLinkedTables = RemoveLinkedTablesEx(Db, "")
  Call xReturn("RemoveLinkedTables")
End Function

Public Sub InitSQLExplorer(ByVal SQLExplorer As Object)
  Set m_SQLExplorer = SQLExplorer
  If DebugMenuInterface Is Nothing Then Call RegisterDB(Nothing)
End Sub

'##D Register DB registers the current file/db as the global current database (for F12 etc..)
Public Sub RegisterDB(vDb As Variant)
  Dim dm As DebugMenu
  
  On Error GoTo RegisterDB_Err
  Call xSet("RegisterDB")
  Set dm = DebugMenuInterface
  If Not dm Is Nothing Then
    Set dm.m_DB = Nothing
    dm.m_DBName = ""
    Set dm.m_DBClass = Nothing
    Set dm = Nothing
  End If
  Set DebugMenuInterface = Nothing
  Set dm = New DebugMenu
  If VarType(vDb) = vbString Then
    dm.m_DBName = vDb
  End If
  If VarType(vDb) = vbObject Then
    If Not vDb Is Nothing Then
      If TypeOf vDb Is Database Then
        Set dm.m_DB = vDb
        dm.m_DBName = dm.m_DB.Name
      End If
    End If
  End If
  Set dm.m_DBClass = Me
  Set DebugMenuInterface = dm
  
RegisterDB_End:
  Call xReturn("RegisterDB")
  Exit Sub

RegisterDB_Err:
  Set DebugMenuInterface = Nothing
  Resume RegisterDB_End
End Sub

Public Function RepairCompactDB2(ByVal DatabasePath As String, ByVal mode As REPAIRCOMPACT_MODE, Optional ByVal ShowErrors As Boolean = True) As Boolean
  RepairCompactDB2 = RepairCompactDBEx(DatabasePath, mode, ShowErrors)
End Function

Public Function RepairCompactDB(ByVal DatabasePath As String, ByVal mode As REPAIRCOMPACT_MODE) As Boolean
  RepairCompactDB = RepairCompactDBEx(DatabasePath, mode, True)
End Function

Public Function CopyIndexes(DestIndexes As Collection, tdSrc As TableDef, Optional ByVal Remove As Boolean = False) As Boolean
  Dim i As Long, tn As String
  On Error GoTo CopyIndexes_Err
  Call xSet("CopyIndexes")
  
  CopyIndexes = True
  tn = tdSrc.Name
  For i = tdSrc.Indexes.Count - 1 To 0 Step -1
    If Remove Then
      CopyIndexes = CopyIndexes And CopyIndex(DestIndexes, tdSrc.Indexes(0), tdSrc.Indexes)
    Else
      CopyIndexes = CopyIndexes And CopyIndex(DestIndexes, tdSrc.Indexes(i), Nothing)
    End If
  Next i
  If Remove Then tdSrc.Indexes.Refresh
  
CopyIndexes_End:
  Call xReturn("CopyIndexes")
  Exit Function

CopyIndexes_Err:
  CopyIndexes = False
  Call ErrorMessage(ERR_ERROR, Err, "CopyIndexes", "Copy Indices from Table", "Error copying indices from table " & tn & ".")
  Resume CopyIndexes_End
End Function

Public Function CopyTableDefEx(tdDest As TableDef, tdSrc As TableDef, ByVal mode As COPYTABLE_MODE, ByVal PropFilter As PROPERTIES_FILTER, Optional ByVal RenameTable As Boolean = False, Optional ByVal UserPropsOnly As Boolean = False) As Boolean
  Dim fd As Field
  Dim cIdx As New Collection
  Dim prop As Property
  
  On Error GoTo CopyTableDefEx_err
  Call xSet("CopyTableDefEx")
  If Not tdDest.Updatable Then GoTo CopyTableDefEx_end
  CopyTableDefEx = True
  For Each fd In tdSrc.Fields
    If Not CopyField(tdDest, fd, mode, PropFilter, UserPropsOnly) Then Call Err.Raise(ERR_COPYTABLEDEF)
  Next fd
  
  If Not UserPropsOnly Then
    CopyTableDefEx = False
    If CopyIndexes(cIdx, tdSrc) Then
      If AddIndexes(tdDest, cIdx, False) Then
        If RenameTable Then tdDest.Name = tdSrc.Name
        CopyTableDefEx = True
      End If
    End If
  End If
  
CopyTableDefEx_end:
  Set fd = Nothing
  Set cIdx = Nothing
  xReturn ("CopyTableDefEx")
  Exit Function
  
CopyTableDefEx_err:
  CopyTableDefEx = False
  Resume CopyTableDefEx_end
End Function


' Copy tdSrc to tdDest
Public Function CopyTableDef(tdDest As TableDef, tdSrc As TableDef, ByVal mode As COPYTABLE_MODE, Optional ByVal RenameTable As Boolean = False, Optional ByVal UserPropsOnly As Boolean = False) As Boolean
  CopyTableDef = CopyTableDefEx(tdDest, tdSrc, mode, PROP_ALL, RenameTable, UserPropsOnly)
End Function

Public Function Transaction(ws As Workspace, ByVal Action As TRANSACTION_TYPE) As Boolean
  On Error GoTo Transaction_err
  
  If Action = TRANS_BEGIN Then
    ws.BeginTrans
  ElseIf Action = TRANS_COMMIT Then
    ws.CommitTrans
  ElseIf Action = TRANS_ROLLBACK Then
    ws.Rollback
  Else
    Err.Raise 380
  End If
  Transaction = True
Transaction_end:
  Exit Function
  
Transaction_err:
  Transaction = False
  Resume Transaction_end
End Function

Public Function TransactionRDO(rdoConn As Object, ByVal Action As TRANSACTION_TYPE) As Boolean
  On Error GoTo TransactionRDO_err
  Dim rdoCN As rdoConnection
  
  Set rdoCN = rdoConn
  If Action = TRANS_BEGIN Then
    rdoCN.BeginTrans
  ElseIf Action = TRANS_COMMIT Then
    rdoCN.CommitTrans
  ElseIf Action = TRANS_ROLLBACK Then
    rdoCN.RollbackTrans
  Else
    Err.Raise 380
  End If
  TransactionRDO = True
TransactionRDO_end:
  Exit Function
  
TransactionRDO_err:
  TransactionRDO = False
  Resume TransactionRDO_end
End Function

Public Function DAOtoDatatype(ByVal daoType As Long) As DATABASE_FIELD_TYPES
  DAOtoDatatype = DAOtoDatatypeEx(daoType)
End Function

Public Function VarTypetoDatatype(ByVal vbType As VbVarType) As DATABASE_FIELD_TYPES
  VarTypetoDatatype = VarTypetoDatatypeEx(vbType)
End Function


Public Function RDOtoDatatype(ByVal rdoType As Long) As DATABASE_FIELD_TYPES
  RDOtoDatatype = RDOtoDatatypeEx(rdoType)
End Function

Public Function IsNumberField(ByVal dbDataType As DATABASE_FIELD_TYPES) As Boolean
  IsNumberField = (dbDataType = TYPE_DOUBLE) Or (dbDataType = TYPE_LONG)
End Function

Public Function DataTypeName(ByVal dbDataType As DATABASE_FIELD_TYPES) As String
  Select Case dbDataType
    Case [_TYPE_INVALID]: DataTypeName = "INVALID"
    Case TYPE_LONG:    DataTypeName = "LONG"    ' 1
    Case TYPE_STR:     DataTypeName = "STRING"  ' 2
    Case TYPE_DOUBLE:  DataTypeName = "DOUBLE"  ' 3
    Case TYPE_DATE:    DataTypeName = "DATE"    ' 4
    Case TYPE_BOOL:    DataTypeName = "BOOLEAN" ' 5
    Case Else
      Call ECASE("Unrecognised data Type: " & CStr(dbDataType))
  End Select
End Function

Public Function GetRDOError(ErrorText As String) As Long
  Dim rd As rdoError, rddesc As String, p As Long
  Dim snum As String, sqlstate As String
  Const lasterror As String = "[SQL Server]"
  
  On Error GoTo GetRDOError_err
  ErrorText = "": GetRDOError = 0
  For Each rd In rdoEngine.rdoErrors
    GetRDOError = rd.Number
    snum = Format$(rd.Number, "00000")
    rddesc = rd.Description
    p = InStr(1, rddesc, lasterror, vbTextCompare)
    If p > 0 Then rddesc = Mid$(rddesc, p + Len(lasterror))
    ErrorText = ErrorText & "RDO Error number: " & Format$(rd.Number, "00000") & "  SQL State: " & rd.sqlstate & vbCrLf & rddesc & vbCrLf
  Next rd
  
GetRDOError_end:
  Exit Function
  
GetRDOError_err:
  Resume GetRDOError_end
End Function

Private Property Get SQLDriverName() As String
  Select Case DatabaseTarget
    Case DB_TARGET_JET
           SQLDriverName = ""
    Case DB_TARGET_SQLSERVER
           SQLDriverName = "SQL Server"
    Case DB_TARGET_ORACLE
           SQLDriverName = "Oracle ODBC Driver"
    Case Else
           Call ECASE("SQLDriverName - Unknown DatabaseTarget: " & CStr(DatabaseTarget))
  End Select
End Property

Public Function DSNAttributes(ByVal Description As String, ByVal ServerName As String, ByVal DBName As String, ByVal NetDriver As String) As String
  Dim s As String
  
  s = "Description=" & Description & vbCr
  If DatabaseTarget = DB_TARGET_ORACLE Then
    s = s & "DBQ=" & ServerName & vbCr
  Else
    s = s & "SERVER=" & ServerName & vbCr & _
            "Database=" & DBName
  End If
  If Len(NetDriver) > 0 Then s = s & vbCr & "Network=" & NetDriver
  DSNAttributes = s
End Function

Public Function DSNConnectString(ByVal DSN_Name As String, ByVal UserID As String, ByVal UserPassword As String) As String
  Dim s As String
  
  s = "DSN=" & DSN_Name
  If Len(UserID) > 0 Then s = s & ";UID=" & UserID
  If Len(UserPassword) > 0 Then s = s & ";PWD=" & UserPassword
  DSNConnectString = s
End Function

Public Function RegisterDataSource(ByVal DSNName As String, ByVal DSNAttributes As String, Optional DriverName As Variant) As Boolean
  Dim State As Long
  
  On Error GoTo RegisterDataSource_err
  If IsMissing(DriverName) Then DriverName = SQLDriverName
   State = -1
  Call DBEngine.RegisterDatabase(DSNName, DriverName, True, DSNAttributes)
  If State > 0 Then
    Call rdoEngine.rdoRegisterDataSource(DSNName, DriverName, True, DSNAttributes)
  End If
  
RegisterDataSource_end:
  RegisterDataSource = True
  Exit Function
  
RegisterDataSource_err:
  If State < 0 Then
    State = 1
    Resume Next
  End If
  RegisterDataSource = False
  Call ErrorMessage(ERR_ERROR, Err, "RegisterDataSource", "Registering Datasource: " & DSNName, "Unable to register data source:" & DSNName & vbCrLf & "Attributes:" & DSNAttributes & vbCrLf & "ODBC Driver:" & DriverName & vbCrLf)
  Resume RegisterDataSource_end
End Function

'DSN, userid & password read from ini file.
'Is there a problem with pwd & uid as strings?
'May want to create more than one connection?  Pass connection to function.
Public Function RDO_Connect(ByVal DSNConnectString As String, ByVal CursorDriver As CursorDriverConstants, Optional ByVal TimeOut As Long = 20) As rdoConnection
  Dim cn As rdoConnection
  Dim i As Long, errtype As errnumbersenum
  
  On Error GoTo RDO_Connect_Err
  Call xSet("RDO_Connect")
  Call SetCursor
  Set cn = New rdoConnection
  If Len(DSNConnectString) = 0 Then Err.Raise ERR_ODBC_CONNECTION, "RDO_Connect", "Attempt to establish a connection failed." & vbCrLf & "No data connection string specified."
retry_connect:
  With cn
    .LoginTimeout = TimeOut
    .Connect = DSNConnectString
    .CursorDriver = CursorDriver
    .EstablishConnection rdDriverNoPrompt, , rdAsyncEnable
  End With
  i = 1
  Do While cn.StillConnecting
    Call Sleep(500)
    i = i + 1
    If i > TimeOut Then Err.Raise ERR_ODBC_CONNECTIONTIMEOUT, "RDO_Connect", "Attempt to establish a connection timed out." & vbCrLf & "Connect: " & cn.Connect & vbCrLf & "Timeout = " & CStr(TimeOut)
  Loop
  i = Len(cn.Version)
  rdoEngine.rdoEnvironments(0).rdoConnections.Add cn
  
RDO_Connect_End:
  Set RDO_Connect = cn
  Call ClearCursor
  Call xReturn("RDO_Connect")
  Exit Function

RDO_Connect_Err:
  errtype = ERR_ERROR + ERR_ALLOWRETRY
  If Err.Number = ERR_ODBC_CONNECTION Then errtype = ERR_ERROR
  If ErrorMessage(errtype, Err, "RDO_Connect", "Connect to Database", "Unable to connect to datasource '" & DSNConnectString & "'.") Then Resume retry_connect
  Set cn = Nothing
  Resume RDO_Connect_End
End Function

Public Sub RDO_CloseConnection(cn As rdoConnection)
  Dim cName As String
  On Error GoTo CloseConnection_err
  If Not cn Is Nothing Then
    cName = cn.Name
    Call cn.Cancel
    Call cn.Close
    Set cn = Nothing
  End If
  
CloseConnection_err:
  Exit Sub
  
CloseConnection_end:
  Call ErrorMessage(ERR_ERROR, Err, "RDO_CloseConnection", "Close connection " & cName, "Unable to close RDO connection")
  Resume CloseConnection_end
End Sub

Public Function ReOpenConnection(ByVal cn As rdoConnection) As rdoConnection
  Dim sConnect As String, sName As String, nLoginTimeout As Long, nCursorDriver As Long
  
  sName = cn.Name
  sConnect = cn.Connect
  nLoginTimeout = cn.LoginTimeout
  nCursorDriver = cn.CursorDriver
  Call RDO_CloseConnection(cn)
  Set ReOpenConnection = RDO_Connect(sConnect, nCursorDriver, nLoginTimeout)
End Function

Public Sub DoDBEvents(Optional ByVal EventType As DB_EVENT_TYPES = IDLE_ONLY)
  Dim DoVBEnents As Boolean
  
  DoVBEnents = (EventType And VB_DOEVENTS)
  If DoVBEnents Then EventType = (EventType Xor VB_DOEVENTS)
  Call DBEngine.Idle(EventType)
  If DoVBEnents Then DoEvents
End Sub

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "Database library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
