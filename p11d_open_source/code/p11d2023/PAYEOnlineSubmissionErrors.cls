VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PAYEOnlineSubmissionErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPAYEOnlineOnError
Private m_errors As ObjectList
Private Sub Class_Initialize()
  Set m_errors = New ObjectList
End Sub
Private Sub IPAYEOnlineOnError_OnError(ByVal count As Long, Error As PAYEOnlineError)
  Call m_errors.Add(Error)
End Sub
Public Property Get Errors() As ObjectList
  Set Errors = m_errors
End Property
