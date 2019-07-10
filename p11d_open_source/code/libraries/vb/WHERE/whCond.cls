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
Public OutputDone As Boolean

Private mField As String
Private mOperator As TCSWHERE_CONDITIONS
Private mValue As Variant
Private mDataType As DATABASE_FIELD_TYPES

Public Function Init(ByVal Field As FieldDetails, ByVal op As TCSWHERE_CONDITIONS, ByVal Value As Variant) As Boolean
  mName = "(Unassigned)"
  mValue = GetTypedValueDefault(Value, Field.DataType, Null)
  If Not IsNull(mValue) Then
    mField = Field.Name
    mOperator = op
    mDataType = Field.DataType
    Init = True
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

Public Property Let Field(ByVal NewValue As String)
  mField = NewValue
End Property

Public Property Get Operator() As TCSWHERE_CONDITIONS
  Operator = mOperator
End Property

Public Property Let Operator(ByVal NewValue As TCSWHERE_CONDITIONS)
  mOperator = NewValue
End Property

Public Property Get Value() As String
  Value = mValue
End Property

Public Property Let Value(ByVal NewValue As String)
  mValue = GetTypedValueDefault(NewValue, mDataType, Null)
  If IsNull(mValue) Then Err.Raise ERR_LETVALUE, "Value", "Unable to set Field " & mField & " to value " & NewValue
End Property

Public Property Get DataType() As DATABASE_FIELD_TYPES
  DataType = mDataType
End Property

Public Property Let DataType(ByVal NewValue As DATABASE_FIELD_TYPES)
  mDataType = NewValue
End Property

Public Property Get ConditionSQL() As String
  Dim SQL As String
  
  Select Case mOperator
    Case NUM_GREATER_THAN
      SQL = SQLFieldName(mField) & ">" & NumSQL(mValue)
    Case NUM_LESS_THAN
      SQL = SQLFieldName(mField) & "<" & NumSQL(mValue)
    Case NUM_EQUAL_TO
      SQL = SQLFieldName(mField) & "=" & NumSQL(mValue)
    Case NUM_GREATER_OR_EQUAL
      SQL = SQLFieldName(mField) & ">=" & NumSQL(mValue)
    Case NUM_LESS_OR_EQUAL
      SQL = SQLFieldName(mField) & "<=" & NumSQL(mValue)
    Case NUM_NOT_EQUAL
      SQL = SQLFieldName(mField) & "<>" & NumSQL(mValue)
    Case NUM_ISEMPTY
      SQL = SQLFieldName(mField) & " Is Null"
    Case STR_CONTAINS
      SQL = SQLFieldName(mField) & " Like " & StrSQL("*" & mValue & "*")
    Case STR_BEGINS
      SQL = SQLFieldName(mField) & " Like " & StrSQL(mValue & "*")
    Case STR_ENDS
      SQL = SQLFieldName(mField) & " Like " & StrSQL("*" & mValue)
    Case STR_EQUALS
      SQL = SQLFieldName(mField) & "=" & StrSQL(mValue)
    Case STR_NOT_INCLUDE
      SQL = SQLFieldName(mField) & "Not Like " & StrSQL("*" & mValue & "*")
    Case STR_ISEMPTY
      SQL = SQLFieldName(mField) & " Is Null OR " & SQLFieldName(mField) & "=''"
    Case DT_ON
      SQL = SQLFieldName(mField) & "=" & DateSQL(mValue)
    Case DT_AFTER
      SQL = SQLFieldName(mField) & ">" & DateSQL(mValue)
    Case DT_BEFORE
      SQL = SQLFieldName(mField) & "<" & DateSQL(mValue)
    Case DT_NOT_ON
      SQL = SQLFieldName(mField) & "<>" & DateSQL(mValue)
    Case DT_ISEMPTY
      SQL = SQLFieldName(mField) & " Is Null"
    Case BOOL_TRUE
      SQL = SQLFieldName(mField) & "=True"
    Case BOOL_FALSE
      SQL = SQLFieldName(mField) & "=False"
    Case Else
      Call ECASE("Unsupported field type")
  End Select
  ConditionSQL = SQL
End Property

Public Property Get ConditionString() As String
  Dim SQL As String
  
  SQL = mField
  Select Case mOperator
    Case NUM_GREATER_THAN
      SQL = SQL & " is greater than " & mValue
    Case NUM_LESS_THAN
      SQL = SQL & " is less than " & mValue
    Case NUM_EQUAL_TO
      SQL = SQL & " is equal to " & mValue
    Case NUM_GREATER_OR_EQUAL
      SQL = SQL & " is greater than or equal to " & mValue
    Case NUM_LESS_OR_EQUAL
      SQL = SQL & " is less than or equal to " & mValue
    Case NUM_NOT_EQUAL
      SQL = SQL & " is not equal to " & mValue
    Case NUM_ISEMPTY
      SQL = SQL & " is empty"
    Case STR_CONTAINS
      SQL = SQL & " contains " & StrSQL(mValue)
    Case STR_BEGINS
      SQL = SQL & " begins with " & StrSQL(mValue)
    Case STR_ENDS
      SQL = SQL & " ends with " & StrSQL(mValue)
    Case STR_EQUALS
      SQL = SQL & " is equal to " & StrSQL(mValue)
    Case STR_NOT_INCLUDE
      SQL = SQL & " does not include " & StrSQL(mValue)
    Case STR_ISEMPTY
      SQL = SQL & " is empty"
    Case DT_ON
      SQL = SQL & " is on " & mValue
    Case DT_AFTER
      SQL = SQL & " is after " & mValue
    Case DT_BEFORE
      SQL = SQL & " is before " & mValue
    Case DT_NOT_ON
      SQL = SQL & " is not on " & mValue
    Case DT_ISEMPTY
      SQL = SQL & " is empty"
    Case BOOL_TRUE
      SQL = SQL & " is true"
    Case BOOL_FALSE
      SQL = SQL & " is false"
    Case Else
      Call ECASE("Unsupported field type")
  End Select
  ConditionString = SQL
End Property

