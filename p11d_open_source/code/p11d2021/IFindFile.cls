VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFindFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Process(ByVal sDirectory As String, ByVal sFile As String, ByVal Index As Long, ByVal lEmployees As Long, ByVal sPAYEref As String)

End Sub
Public Sub Prenotify(ByVal Path As String, ByVal Count As Long)
'prenotify ie the count and the directory
End Sub
Public Sub Notify(ByVal lTotalNoOfFiles As Long)
End Sub
Public Sub PostNotify()
End Sub


