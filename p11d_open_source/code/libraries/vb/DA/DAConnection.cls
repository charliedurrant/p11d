VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Set StatusBar(newval As Object)
  If TypeOf newval Is TCSStatus Then
    Set gStatusBar = New DAStatusBar
    Set gStatusBar.StatusBar = newval
    gStatusBar.StatMax = 0
    gStatusBar.StatMin = 0
  End If
End Property

Public Property Let StatMin(newval As Long)
  If Not gStatusBar Is Nothing Then gStatusBar.StatMin = newval
End Property

Public Property Let StatMax(newval As Long)
  If Not gStatusBar Is Nothing Then gStatusBar.StatMax = newval
End Property

Public Property Let ShowPopMessages(newval As Boolean)
  gShowPopMessages = newval
End Property

Public Function CloseConnection(Optional sAppName As String = "")
  Call ForceCloseDatabases(sAppName)
End Function

Private Function ForceCloseDatabases(Optional sAppName As String = "")
Dim DAdb As DADatabase
Dim i As Integer
  
  On Error Resume Next
  If Not gOpenDatabases Is Nothing Then
    If gOpenDatabases.Count Then
      For i = gOpenDatabases.Count To 1 Step -1
        Set DAdb = gOpenDatabases(i)
        If Len(sAppName) Then
          Call DAdb.CloseDb(sAppName)
        Else
          DAdb.ForceCloseDb
        End If
      Next i
    End If
    Set gOpenDatabases = Nothing
  End If
End Function

Public Sub LogRelations(ctrl As Object)
Dim tv As TreeView
Dim lv As ListView
Dim db As DADatabase
  
  On Error Resume Next
  If TypeOf ctrl Is TreeView Then
    Set tv = ctrl
    Set lv = Nothing
  ElseIf TypeOf ctrl Is ListView Then
    Set lv = ctrl
    Set tv = Nothing
  Else
    Exit Sub
  End If
  
  If Not lv Is Nothing Then
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    lv.ColumnHeaders.Add 1, "Object", "Object"
    lv.ColumnHeaders.Add 2, "Type", "Type"
    lv.ColumnHeaders.Add 3, "Relation", "Relation"
    lv.ColumnHeaders.Add 4, "Dirty Relation", "Dirty Relation?"
    If Not gOpenDatabases Is Nothing Then
      For Each db In gOpenDatabases
        db.LogRelations lv, LIST_REL
      Next db
    End If
  Else
    tv.Nodes.Clear
    If Not gOpenDatabases Is Nothing Then
      For Each db In gOpenDatabases
        db.LogRelations tv, TREE_TOP_REL
      Next db
    End If
  End If
  Set db = Nothing
  Set lv = Nothing
  Set tv = Nothing
End Sub

Public Sub LogObjects(frm As Object)
Dim DAdb As DADatabase

Dim tlv As ListView

  On Error Resume Next
  If Not TypeOf frm Is Form Then Exit Sub
  If Not (InCollection(frm.Controls, "lv") _
    And InCollection(frm.Controls, "lbl")) Then Exit Sub
  If Not (TypeOf frm.lv Is ListView _
    And TypeOf frm.lbl Is Label) Then Exit Sub
  frm.Caption = "List of objects loaded in memory"
  frm.lbl.Caption = ""
  Set tlv = frm.lv
  tlv.ListItems.Clear
  tlv.ColumnHeaders.Clear
  
  tlv.ColumnHeaders.Add 1, "Database", "Database"
  tlv.ColumnHeaders.Add 2, "Type", "Type"
  tlv.ColumnHeaders.Add 3, "Name", "Name"
  tlv.ColumnHeaders.Add 4, "Dirty", "Dirty"
  tlv.ColumnHeaders.Add 5, "Modes", "Modes"
  
  
  If Not gOpenDatabases Is Nothing Then
    For Each DAdb In gOpenDatabases
      DAdb.LogObjects frm
    Next DAdb
  End If
  Set DAdb = Nothing
  Set tlv = Nothing
End Sub

Private Sub Class_Initialize()
  On Error GoTo DAInit_err
  
  If gOpenDatabases Is Nothing Then Set gOpenDatabases = New Collection
  Set gStatusBar = New DAStatusBar

  gShowPopMessages = True

DAInit_end:
  Exit Sub
  
DAInit_err:
  Call ErrorMessage(ERR_ERROR, Err, "DAConnection_Initialise", "Initialise connections", "Error initialising a connection to a database.")
  Resume DAInit_end
End Sub

Public Function OpenDatabase(sAppName As String, Filename As String, idString As String, Optional bDBExclusive As Boolean = False, Optional bTBONewFixes As Boolean = True) As DADatabase
  Dim DAdb As DADatabase
  Dim cLocks As Collection
  
  On Error GoTo OpenDatabase_Err
  If gOpenDatabases Is Nothing Then Call Class_Initialize
    
  If Not InCollection(gOpenDatabases, Filename) Then
    Set DAdb = New DADatabase
    'If bDBExclusive Then DAdb.AccessMode = True
    'DAdb.AccessMode = True
    If DAdb.OpenDB(Filename, idString, , bTBONewFixes) Then
      If DAdb.Locks Is Nothing Then Set DAdb.Locks = New Collection
      Set cLocks = DAdb.Locks
      If Not InStrCollection(cLocks, sAppName) Then cLocks.Add sAppName, sAppName
      DAdb.inclockcount
      Debug.Print "File opened (" & Filename & ") by " & sAppName
      Set OpenDatabase = DAdb
    Else
      Set OpenDatabase = Nothing
      #If DEBUGVER Then
        logmessage ("OpenDatabase: An error occurred opening the database, " & Filename & ", for " & sAppName)
      #End If
    End If
  Else
    Set DAdb = gOpenDatabases(Filename)
    If DAdb.Locks Is Nothing Then Set DAdb.Locks = New Collection
    Set cLocks = DAdb.Locks
    If Not InStrCollection(cLocks, sAppName) Then cLocks.Add sAppName, sAppName
    DAdb.inclockcount
    Set OpenDatabase = DAdb
  End If

OpenDatabase_End:
'  If Not DAdb Is Nothing Then
'    Call DAdb.CloseDb(sAppName)
'    Set DAdb = Nothing
'  End If
  Exit Function
    
OpenDatabase_Err:
  Err.Raise ERR_OPEN_DATABASE, "OpenDatabase", "Error opening database '" & Filename & "'. " & vbCrLf & Err.Description
End Function

Private Sub Class_Terminate()
  Dim i As Long, errno As Long
    
  On Error GoTo DATerm_err
  Call CloseConnection
  
  Set gOpenDatabases = Nothing
  Set gStatusBar = Nothing
  Set gNotify = Nothing
  
DATerm_end:
  Exit Sub
  
DATerm_err:
  errno = Err.Number
  Call ErrorMessage(ERR_ERROR, Err, "DATerm", "Terminate Data Access", "Error terminating Data Access component")
  If errno = ERR_NODATABASES Then Resume DATerm_end
  Resume Next
End Sub

Public Property Let DebugMode(ByVal NewValue As Boolean)
  gDebugMode = NewValue
End Property

Public Property Get DebugMode() As Boolean
  DebugMode = gDebugMode
End Property

Public Function RepairCompact(sFile As String) As Boolean
  Dim DAdb As DADatabase

  On Error GoTo RepairCompact_Err
  Call xSet("RepairCompact")

  If Not FileExists(sFile, False) Then Err.Raise ERR_NODB
  If Not InCollection(gOpenDatabases, sFile) Then
    Set DAdb = New DADatabase
    DAdb.name = sFile
  Else
    Set DAdb = gOpenDatabases(sFile)
  End If
  DAdb.RepairCompact

RepairCompact_End:
  Call xReturn("RepairCompact")
  Exit Function

RepairCompact_Err:
  Call ErrorMessage(ERR_ERROR, Err, "RepairCompact", "Error in RepairCompact", "Undefined error.")
  Resume RepairCompact_End
End Function

Public Property Set NotifyInterface(ByVal newval As IBaseNotify)
  Set gNotify = newval
End Property
