VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum DATABASE_TARGET
  DB_TARGET_JET = 0
  DB_TARGET_SQLSERVER
  DB_TARGET_ORACLE
End Enum
Private mDBTargetSet As Boolean
Private mDBTarget As DATABASE_TARGET
Private mDataConv As DataConversion
' MPS - Used in BoolSQL
Private mBooleanValueTrue As String
Private mBooleanValueFalse As String

Private Sub Class_Initialize()
  Set mDataConv = New DataConversion
  mDBTargetSet = False
End Sub

'##D Sets/Gets the current database target (SQL/Access/Oracle)
'##LD Setting the value of this property changes the behaviour of Database and SQL functions
'##C DF
Public Property Get DatabaseTarget() As DATABASE_TARGET
  DatabaseTarget = mDBTarget
End Property

Public Property Let DatabaseTarget(ByVal NewValue As DATABASE_TARGET)
  mDBTarget = NewValue
  mBooleanValueTrue = "-1"
  mBooleanValueFalse = "0"
  mDBTargetSet = True
End Property

'MPS - Used by BoolSQL
Public Property Get BooleanValueTrue() As String
  BooleanValueTrue = mBooleanValueTrue
End Property

Public Property Let BooleanValueTrue(ByVal NewValue As String)
  mBooleanValueTrue = NewValue
End Property

Public Property Get BooleanValueFalse() As String
  BooleanValueFalse = mBooleanValueFalse
End Property

Public Property Let BooleanValueFalse(ByVal NewValue As String)
  mBooleanValueFalse = NewValue
End Property

'##D Prepares a string to use in SQL
'##V vString source string
'##RV Valid SQL string
'##C DF
Public Function StrSQL(ByVal vString As Variant) As String
  If IsNull(vString) Then
    StrSQL = "Null"
  Else
    StrSQL = "'" & Replace(CStr(vString), "'", "''") & "'"
  End If
End Function

'##D Prepares a date to use in SQL
'##RV Valid SQL date as a string
'##C DF
Public Function DateSQL(ByVal vDate As Variant) As String
  If Not mDBTargetSet Then Err.Raise ERR_DBHELPER, "DateSQL", "Cannot access atecvbrt.DBHelper properties before setting the DatabaseTarget property"
  ' MPS - A non-date argument returns NULL
  If IsNull(vDate) Or (Not IsDate(vDate)) Then
    DateSQL = "NULL"
  Else
    If mDBTarget = DB_TARGET_JET Then
      DateSQL = "#" & Format$(vDate, "MM/DD/YYYY") & "#"
    ElseIf mDBTarget = DB_TARGET_SQLSERVER Then
      DateSQL = "'" & Format$(vDate, "YYYYMMDD") & "'"
    ElseIf mDBTarget = DB_TARGET_ORACLE Then
      DateSQL = "TO_DATE('" & Format$(vDate, "YYYYMMDD") & "','YYYYMMDD')"
    End If
  End If
End Function

'##D Prepares a date time to use in SQL
'##RV Valid SQL date as a string
'##C DF
Public Function DateTimeSQL(ByVal vDateTime As Variant) As String
  If Not mDBTargetSet Then Err.Raise ERR_DBHELPER, "DateTimeSQL", "Cannot access atecvbrt.DBHelper properties before setting the DatabaseTarget property"
  ' MPS - A non-date argument returns NULL
  If IsNull(vDateTime) Or (Not IsDate(vDateTime)) Then
    DateTimeSQL = "NULL"
  Else
    If mDBTarget = DB_TARGET_JET Then
      DateTimeSQL = "#" & Format$(vDateTime, "MM/DD/YYYY HH:NN:SS") & "#"
    ElseIf mDBTarget = DB_TARGET_SQLSERVER Then
      DateTimeSQL = "'" & Format$(vDateTime, "YYYYMMDD HH:NN:SS") & "'"
    ElseIf mDBTarget = DB_TARGET_ORACLE Then
      DateTimeSQL = "TO_DATE('" & Format$(vDateTime, "YYYYMMDD HH:NN:SS") & "','YYYYMMDD HH24:MI:SS')"
    End If
  End If
End Function

'##C DF
Public Function NumSQL(ByVal vNumber As Variant) As String
  ' MPS - A non-numeric argument returns NULL
  If IsNull(vNumber) Or (Not IsNumeric(vNumber)) Then
    NumSQL = "NULL"
  Else
    If VarType(vNumber) = vbBoolean Then
      If vNumber Then
        NumSQL = "-1"
      Else
        NumSQL = "0"
      End If
    Else
      NumSQL = CStr(vNumber)
    End If
  End If
End Function

'MPS - New function
Public Function BoolSQL(ByVal vValue As Variant) As String
  ' MPS - A non-boolean argument returns False value
  If mDataConv.IsBoolean(vValue) Then
    If mDataConv.CBoolean(vValue) Then
      BoolSQL = BooleanValueTrue
    Else
      BoolSQL = BooleanValueFalse
    End If
  Else
    BoolSQL = BooleanValueFalse
  End If
End Function

Public Property Get SQLDriverName() As String
  If Not mDBTargetSet Then Err.Raise ERR_DBHELPER, "SQLDriverName", "Cannot access atecvbrt.DBHelper properties before setting the DatabaseTarget property"
  Select Case DatabaseTarget
    Case DB_TARGET_JET
      SQLDriverName = "Jet"
    Case DB_TARGET_SQLSERVER
      SQLDriverName = "SQL Server"
    Case DB_TARGET_ORACLE
      SQLDriverName = "Oracle"
    Case Else
      Err.Raise ERR_DBHELPER, "SQLDriverName", "Unknown DatabaseTarget [" & CStr(DatabaseTarget) & "]"
  End Select
End Function

'##D Converts Value to String variable for use in SQL statements ( such as SELECT/INSERT )
'##LD Always converts to a string SQL value no matter what the underlying type
'##C GCF
Public Function GetSQLValueAsString(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  If Not (IsNull(v) Or IsEmpty(v)) Then v = mDataConv.GetValueAsString(v, dType)
  GetSQLValueAsString = StrSQL(v)
End Function

'##D Converts Value to SQL String variable for use in SQL statements ( such as SELECT/INSERT )
'##C GCF
Public Function GetSQLValue(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  If IsNull(v) Or IsEmpty(v) Then
    v = "Null"
  Else
    v = mDataConv.GetTypedValue(v, dType)
    Select Case dType
      Case TYPE_LONG, TYPE_DOUBLE
        v = NumSQL(v)
      Case TYPE_STR
        v = StrSQL(v)
      Case TYPE_DATE
        If mDataConv.IsDateTime(v) Then
          v = DateTimeSQL(v)
        Else
          v = DateSQL(v)
        End If
      Case TYPE_BOOL
        v = NumSQL(v)
    End Select
  End If
  GetSQLValue = v
End Function



