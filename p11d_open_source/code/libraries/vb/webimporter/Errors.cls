VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Number As Long
Public Description As String
Public Source As String
Public Logged As Boolean

Public Property Get ErrorDescription() As String
  ErrorDescription = ReplaceString(Description, vbCrLf, "<BR>")
  Logged = True
End Property

Public Property Get ErrorDetails() As String
  ErrorDetails = "<BR>Error Number: 0x" & Right$("00000000" & Hex$(Number), 8) & "<BR>" & _
                "<B>Error call stack</B><BR>" & ReplaceString(Source, ";", "<BR>")
  Logged = True
End Property

Private Sub Class_Initialize()
  Logged = False
End Sub

Public Property Get ErrorDescriptionXML() As String
  ErrorDescriptionXML = Description
  Logged = True
End Property

Public Property Get ErrorDetailsXML() As String
  ErrorDetailsXML = "<NUMBER>Error Number: 0x" & Right$("00000000" & Hex$(Number), 8) & "</NUMBER>"
  ErrorDetailsXML = ErrorDetailsXML & "<CALLSTACK><SOURCE>" & ReplaceString(Source, ";", "</SOURCE><SOURCE>") & "</SOURCE></CALLSTACK>"
  Logged = True
End Property




