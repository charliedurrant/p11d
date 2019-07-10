VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ReadFromFile() As Boolean

End Property
Public Property Let ReadFromFile(NewValue As Boolean)

End Property

Public Function FileOK(sPathAndFile As String) As Boolean
  
End Function
Public Sub OpenFile(sPathAndFile As String)
  
End Sub
Public Sub ReadFile()
  'actually reads tha data
End Sub
Public Property Get FileRead() As TCSFileread
  
End Property
Public Property Set FileRead(NewValue As TCSFileread)

End Property
