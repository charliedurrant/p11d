VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMagneticMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get FieldSize(ByVal Item As Long) As Long
  'size of MM field
End Property
Public Property Get FieldOK(ByVal Item As Long) As Boolean
  'is the field ok example is
End Property
Public Property Get Required(ByVal Item As Long)
  'is the field required for MM if datatype is string then > "" if numeric then > 0
End Property
Public Property Let FieldSize(ByVal Item As Long, FieldSize As Long)
  'set Get FieldSize
End Property

