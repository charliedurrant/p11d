VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mdbSrc As Database
Private mdbDest As Database
Private mUpdateTables As StringList
Private mDeleteTables As StringList
Private mUpdateQueries As StringList
Private mDeleteQueries As StringList
Private mStructureFilters As Collection
Private mDataFilters As Collection
Private mCheckFldFlags() As String
Private mCheckIdxFlags() As String
Private mLog As QString

Public Function SynchroniseDatabases(ByVal dbSource As Database, ByVal dbDest As Database, Optional ByVal iNotify As IBaseNotify) As Boolean
  Dim flt As syncFilter
  Dim dbChange As New dbChangeDetails
  
  On Error GoTo SynchroniseDatabases_Err
  Call xSet("SynchroniseDatabases")
    
  mUpdateQueries.RemoveAll
  mUpdateTables.RemoveAll
  mDeleteQueries.RemoveAll
  mDeleteQueries.RemoveAll
  If Not iNotify Is Nothing Then Set gNotify = iNotify
  Set mdbSrc = dbSource
  Set mdbDest = dbDest
  mLog = "<XML><START>" & Now & "</START>"
  Call CheckTables(mdbSrc, mdbDest, mCheckFldFlags, mCheckIdxFlags, mLog, dbChange)
  Call CheckAllQueries(mdbSrc, mdbDest, dbChange, mLog)
  Call syncStructure(mdbSrc, mdbDest, mStructureFilters, mLog, dbChange, mCheckFldFlags, mCheckIdxFlags)
  Call syncData(mDataFilters, mdbSrc, mdbDest)
  Call ProvideFeedback(-1, -1, "")
  mLog.Append ("<END>" & Now & "</END></XML>")

SynchroniseDatabases_End:
  Call xReturn("SynchroniseDatabases")
  Exit Function

SynchroniseDatabases_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SynchroniseDatabases", "Error in SynchroniseDatabases", "Undefined error.")
  Resume SynchroniseDatabases_End
  Resume
End Function

Private Function Kill()
  Dim i As Long
  On Error Resume Next
  
  Set mdbSrc = Nothing
  Set mdbDest = Nothing
  Set mUpdateTables = Nothing
  Set mDeleteTables = Nothing
  Set mUpdateQueries = Nothing
  Set mDeleteQueries = Nothing
  Set mLog = Nothing
  Set gNotify = Nothing
  If Not mStructureFilters Is Nothing Then
    For i = 1 To mStructureFilters.Count
      Set mStructureFilters(i) = Nothing
    Next i
    Set mStructureFilters = Nothing
  End If
  If Not mDataFilters Is Nothing Then
    For i = 1 To mDataFilters.Count
      Set mDataFilters(i) = Nothing
    Next i
    Set mDataFilters = Nothing
  End If
End Function

Private Sub Class_Initialize()
  
  Set mLog = New QString
  mCheckFldFlags = Split("Name", ";")
  mCheckIdxFlags = Split("Clustered;Foreign;IgnoreNulls;Primary;Required;Unique", ";")
  Set mUpdateQueries = New StringList
  Set mUpdateTables = New StringList
  Set mDeleteQueries = New StringList
  Set mDeleteTables = New StringList
  
End Sub

Private Sub Class_Terminate()
  Call Kill
End Sub

Public Property Let FieldPropertyChecks(NewVal As String)
  mCheckFldFlags = Split(NewVal, ";")
End Property

Public Property Get LogString() As String
  LogString = mLog.bstr
End Property

Public Function AddFilter(sFilterString As String, Optional lFilterType As FILTER_TYPE = INCLUDE_FILTER, Optional lFilterAction As Filter_Action = FULL_SYNC, Optional sException As String = "") As Boolean
  Dim flt As syncFilter
  
  On Error GoTo AddFilter_Err
  Call xSet("AddFilter")
    
  If mDataFilters Is Nothing Then Set mDataFilters = New Collection
  If mStructureFilters Is Nothing Then Set mStructureFilters = New Collection
  
  Set flt = New syncFilter
  If flt.SetupFilter(sFilterString, lFilterType, lFilterAction, sException) Then
    If flt.StructureFilter Then
      Call mStructureFilters.Add(flt)
      AddFilter = True
    End If
    If flt.DataFilter Then
      Call mDataFilters.Add(flt)
      AddFilter = True
    'Else
      'Err.Raise ERR_UNKNOWN_FILTER_TYPE, "AddFilter", "Unknown filter type"
    End If
  End If
  
AddFilter_End:
  Call xReturn("AddFilter")
  Exit Function

AddFilter_Err:
  AddFilter = False
  Call ErrorMessage(ERR_ERROR, Err, "AddFilter", "Error in AddFilter", "Undefined error.")
  Resume AddFilter_End
  Resume
End Function

