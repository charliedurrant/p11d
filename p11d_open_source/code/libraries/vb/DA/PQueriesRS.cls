VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParsedQueriesRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ParsedQueries As Collection

Public Sub Kill()
  Dim pq As ParsedQueryRS
  
  If Not m_ParsedQueries Is Nothing Then
    For Each pq In m_ParsedQueries
      Call pq.Kill
    Next pq
  End If
  Set m_ParsedQueries = Nothing
End Sub

Private Sub Class_Initialize()
  Set m_ParsedQueries = New Collection
End Sub

Private Function GetParsedQueryRSEx(ByVal rsType As RS_TYPE) As ParsedQueryRS
  On Error Resume Next
  Set GetParsedQueryRSEx = m_ParsedQueries.item(CStr(rsType))
End Function

Public Function GetParsedQueryRS(ByVal db As database, ByVal rsType As RS_TYPE, Optional ByVal SnapShot As Boolean = True) As recordset
  Dim pq As ParsedQueryRS
  
  Set pq = GetParsedQueryRSEx(rsType)
  If pq Is Nothing Then
    Set pq = New ParsedQueryRS
    Call pq.Setup(db, rsType, SnapShot)
    m_ParsedQueries.Add pq, CStr(pq.rsType)
  Else
    If pq.SnapShot <> SnapShot Then
      Call pq.Kill
      Call pq.Setup(db, rsType, SnapShot)
    End If
  End If
  Set GetParsedQueryRS = pq.rs
End Function
