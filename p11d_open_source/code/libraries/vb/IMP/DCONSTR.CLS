VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public DestField As String
Public Constraint As IMPORT_CONSTRAINT
Public DataType  As DATABASE_FIELD_TYPES
Public ConstraintList As StringList
Public Mask As String
Public minvalue As Variant
Public maxvalue As Variant
Public IgnoreNumericConversionErrors As Boolean

Private Sub Class_Initialize()
  Set ConstraintList = New StringList
End Sub
