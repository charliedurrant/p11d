VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRecentFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum RFL_ACTION
  RFL_OK
  [_RFL_FIRST_ITEM] = RFL_OK
  RFL_DISABLE
  RFL_REMOVE
  [_RFL_LAST_ITEM] = RFL_REMOVE
End Enum

Public Function Validate(ByVal PathAndFile As String) As RFL_ACTION
  ' Validate file
End Function

