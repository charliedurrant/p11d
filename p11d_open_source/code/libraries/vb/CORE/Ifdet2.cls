VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Name As String
Public DefaultValue As Variant
Public DataType As DATABASE_FIELD_TYPES
Public Length As Long
Public Required As Boolean
Public AllowZeroLength As Boolean

' The following are additions to IFieldDetails
Public Value As Variant
Public Description As String
Public Visible As Boolean
Public Tag As Variant

'##CD FieldDetails class is used to as a generic representation of a field in a recordset

