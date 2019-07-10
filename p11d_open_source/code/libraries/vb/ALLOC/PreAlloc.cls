VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PreAllocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum PREALLOC_OBJECTS
  PREALLOC_PARSER = 1
  PREALLOC_REP = 2
  PREALLOC_AUTO = 4
  PREALLOC_AUTO_ADO = 8
End Enum

Public Sub AllocObjects(ByVal PreAllocObjects As PREALLOC_OBJECTS)
  Call KillPreAllocated
  If (PreAllocObjects And PREALLOC_PARSER) = PREALLOC_PARSER Then
    Set PreParser = New Parser
  End If
  If (PreAllocObjects And PREALLOC_REP) = PREALLOC_REP Then
    Set PreRep = New Reporter
  End If
  If (PreAllocObjects And PREALLOC_AUTO) = PREALLOC_AUTO Then
    Set PreAuto = New AutoClass
  End If
  If (PreAllocObjects And PREALLOC_AUTO_ADO) = PREALLOC_AUTO_ADO Then
    Set PreADOAuto = New AutoClass
  End If
End Sub

Private Sub KillPreAllocated()
  Set PreParser = Nothing
  Set PreRep = Nothing
  If Not PreAuto Is Nothing Then
    Call PreAuto.Kill
    Set PreAuto = Nothing
  End If
  If Not PreADOAuto Is Nothing Then
    Call PreADOAuto.Kill
    Set PreADOAuto = Nothing
  End If
End Sub

Private Sub Class_Terminate()
  Call KillPreAllocated
End Sub
