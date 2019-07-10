VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbChangeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ChangedTables As Collection
Public ChangedQueries As Collection
Public OldTables As StringList
Public NewTables As StringList
Public OldQueries As StringList

Private Sub Class_Initialize()
  Set ChangedTables = New Collection
  Set ChangedQueries = New Collection
  Set OldTables = New StringList
  Set NewTables = New StringList
  Set OldQueries = New StringList
End Sub

Private Sub Class_Terminate()
  Set ChangedTables = Nothing
  Set ChangedQueries = Nothing
  Set OldTables = Nothing
  Set NewTables = Nothing
  Set OldQueries = Nothing
End Sub
