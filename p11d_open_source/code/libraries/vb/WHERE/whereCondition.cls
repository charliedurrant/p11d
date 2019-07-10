VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "whereCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String

Private mField As String
Private mOperator As String
Private mValue As Variant

Public Function Init(ByVal Field As FieldDetails, ByVal Op As TCSWHERE_CONDITIONS, ByVal Value As Variant) As Boolean
  mName = "(Unassigned)"
  mValue = GetTypedValueDefault(Value, Field.DataType, Null)
  If Not IsNull(mValue) Then
    mField = Field.Name
    mOperator = Op
  End If
End Function

Public Property Get Name() As String
  Name = mName
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = NewValue
End Property

Public Property Get Field() As String
  Field = mField
End Property

Public Property Get Operator() As String
  Operator = mOperator
End Property

Public Property Get Value() As String
  Value = mValue
End Property

Public Property Get ConditionSQL() As String
  Dim sql As String
  
  Select Case mOperator
    Case NUM_GREATER_THAN
      sql = SQLFieldName(mField) & ">" & NumSQL(mValue)
    Case NUM_LESS_THAN
      sql = SQLFieldName(mField) & "<" & NumSQL(mValue)
    Case NUM_EQUAL_TO
      sql = SQLFieldName(mField) & "=" & NumSQL(mValue)
    Case NUM_GREATER_OR_EQUAL
      sql = SQLFieldName(mField) & ">=" & NumSQL(mValue)
    Case NUM_LESS_OR_EQUAL
      sql = SQLFieldName(mField) & "<=" & NumSQL(mValue)
    Case NUM_NOT_EQUAL
      sql = SQLFieldName(mField) & "<>" & NumSQL(mValue)
    Case NUM_ISEMPTY
      sql = SQLFieldName(mField) & " Is Null"
    Case STR_CONTAINS
      sql = SQLFieldName(mField) & " Like " & StrSQL("*" & mValue & "*")
    Case STR_BEGINS
      sql = SQLFieldName(mField) & " Like " & StrSQL(mValue & "*")
    Case STR_ENDS
      sql = SQLFieldName(mField) & " Like " & StrSQL("*" & mValue)
    Case STR_EQUALS
      sql = SQLFieldName(mField) & "=" & StrSQL(mValue)
    Case STR_NOT_INCLUDE
      sql = SQLFieldName(mField) & "Not Like " & StrSQL("*" & mValue & "*")
    Case STR_ISEMPTY
      sql = SQLFieldName(mField) & " Is Null OR " & SQLFieldName(mField) & "=''"
    Case DT_ON
      sql = SQLFieldName(mField) & "=" & DateSQL(mValue)
    Case DT_AFTER
      sql = SQLFieldName(mField) & ">" & DateSQL(mValue)
    Case DT_BEFORE
      sql = SQLFieldName(mField) & "<" & DateSQL(mValue)
    Case DT_NOT_ON
      sql = SQLFieldName(mField) & "<>" & DateSQL(mValue)
    Case DT_ISEMPTY
      sql = SQLFieldName(mField) & " Is Null"
    Case BOOL_TRUE
      sql = SQLFieldName(mField) & "=True"
    Case BOOL_FALSE
      sql = SQLFieldName(mField) & "=False"
    Case Else
      Call ECASE("Unsupported field type")
  End Select
  ConditionSQL = "(" & sql & ")"
End Property

Public Property Get ConditionString() As String
  Dim sql As String
  
  sql = mField
  Select Case mOperator
    Case NUM_GREATER_THAN
      sql = sql & " is greater than " & mValue
    Case NUM_LESS_THAN
      sql = sql & " is less than " & mValue
    Case NUM_EQUAL_TO
      sql = sql & " is equal to " & mValue
    Case NUM_GREATER_OR_EQUAL
      sql = sql & " is greater than or equal to " & mValue
    Case NUM_LESS_OR_EQUAL
      sql = sql & " is less than or equal to " & mValue
    Case NUM_NOT_EQUAL
      sql = sql & " is not equal to " & mValue
    Case NUM_ISEMPTY
      sql = sql & " is empty"
    Case STR_CONTAINS
      sql = sql & " contains " & StrSQL(mValue)
    Case STR_BEGINS
      sql = sql & " begins with " & StrSQL(mValue)
    Case STR_ENDS
      sql = sql & " ends with " & StrSQL(mValue)
    Case STR_EQUALS
      sql = sql & " is equal to " & StrSQL(mValue)
    Case STR_NOT_INCLUDE
      sql = sql & " does not include " & StrSQL(mValue)
    Case STR_ISEMPTY
      sql = sql & " is empty"
    Case DT_ON
      sql = sql & " is on " & mValue
    Case DT_AFTER
      sql = sql & " is after " & mValue
    Case DT_BEFORE
      sql = sql & " is before " & mValue
    Case DT_NOT_ON
      sql = sql & " is not on " & mValue
    Case DT_ISEMPTY
      sql = sql & " is empty"
    Case BOOL_TRUE
      sql = sql & " is true"
    Case BOOL_FALSE
      sql = sql & " is false"
    Case Else
      Call ECASE("Unsupported field type")
  End Select
  ConditionString = sql
End Property

