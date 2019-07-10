VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iAutoAuditRecord
Private mDB As Database
Public ShowChangesOnly As Boolean

Private Sub Class_Initialize()
  Set mDB = InitDB(gwsMain, AppPath & "\test.mdb", "Test Database")
End Sub

Private Sub Class_Terminate()
  Set mDB = Nothing
End Sub

Private Property Get iAutoAuditRecord_AuditLevel() As TCSBASE.GRIDAUDIT_LEVELS
  iAutoAuditRecord_AuditLevel = GRIDAUDIT_FULL
End Property

Private Sub IAutoAuditRecord_RecordUpdated(ByVal UserName As String, ByVal DateTime As Date, ByVal EditType As TCSBASE.GRIDEDIT_TYPE, ByVal Fields As Collection)
  Dim iFld As FieldDetails, sql0 As String, sql1 As String
  Dim v As Variant
  
  On Error GoTo RecordUpdate_err
  If Not Fields Is Nothing Then
    sql0 = "INSERT INTO AuditTable (User, AuditDateTime, UpdateType, TableName, FieldName, NewValue, OldValue) " & _
           "VALUES (" & StrSQL(UserName) & "," & DateTimeSQL(DateTime) & ","
    If EditType = GRID_EDIT Then
      sql0 = sql0 & StrSQL("EDIT") & ","
    ElseIf EditType = GRID_DELETE Then
      sql0 = sql0 & StrSQL("DELETE") & ","
    ElseIf EditType = GRID_ADDNEW Then
      sql0 = sql0 & StrSQL("ADDNEW") & ","
    Else
      Call ECASE("Invalid Edit Type in Auto Audit Class")
    End If
    For Each iFld In Fields
      If ShowChangesOnly Then If (iFld.Value = iFld.Tag) Or (IsNull(iFld.Value) And IsNull(iFld.Tag)) Then GoTo skip_field
      sql1 = StrSQL(iFld.Description) & "," & StrSQL(iFld.Name) & "," & _
             GetSQLValueAsString(iFld.Value, iFld.DataType) & "," & GetSQLValueAsString(iFld.Tag, iFld.DataType) & ")"
      Call mDB.Execute(sql0 & sql1)
skip_field:
    Next iFld
  End If
RecordUpdate_end:
  Exit Sub
  
RecordUpdate_err:
  Call ErrorMessage(ERR_ERROR, Err, "AuditRecordUpdated", "Error logging record update to audit trail", "Error in record update." & vbCrLf & "SQL: " & sql0 & sql1)
  Resume RecordUpdate_end
End Sub

Private Function GetSQLValueAsString(ByVal Value As Variant, ByVal dType As DATABASE_FIELD_TYPES)
  If Not (IsNull(Value) Or IsEmpty(Value)) Then Value = GetValueAsString(Value, dType)
  GetSQLValueAsString = StrSQL(Value)
End Function

