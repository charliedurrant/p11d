VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAutoAuditRecordADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub RecordUpdated(ByVal UserName As String, ByVal DateTime As Date, ByVal EditType As GRIDEDIT_TYPE, ByVal rs As Recordset)
  ' Called when a Record is Updated in an Auto Control
End Sub

Public Property Get AuditLevel() As GRIDAUDIT_LEVELS
  ' What audit level does the grid need to use
End Property

