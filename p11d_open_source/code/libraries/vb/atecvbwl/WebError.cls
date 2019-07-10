VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Logged As Boolean
Private m_Number As Long
Private m_Description As String
Private m_Source As String

Public Property Get Number() As Long
  Number = m_Number
End Property

Friend Property Let Number(ByVal NewValue As Long)
  m_Number = NewValue
End Property

Public Property Get Description() As String
  Description = m_Description
End Property

Friend Property Let Description(ByVal NewValue As String)
  m_Description = NewValue
End Property

Public Property Get Source() As String
  Source = m_Source
End Property

Friend Property Let Source(ByVal NewValue As String)
  m_Source = NewValue
End Property

Friend Property Get url(ByVal ErrorIndex As Long) As String
  ' URL
  url = "NUMBER" & ErrorIndex & "=" & Number & _
              "&DESCRIPTION" & ErrorIndex & "=" & URLEncode(ReplaceHTMLMetacharactersEx(Description)) & _
              "&SOURCE" & ErrorIndex & "=" & URLEncode(ReplaceHTMLMetacharactersEx(Source))
  m_Logged = True
End Property

Public Property Get DescriptionHTML() As String
  ' MPS
  DescriptionHTML = ReplaceString(ReplaceHTMLMetacharactersEx(m_Description), vbCrLf, "<BR>")
  m_Logged = True
End Property

Public Property Get DetailsHTML() As String
  ' MPS
  DetailsHTML = "<BR>Error Number: 0x" & Right$("00000000" & Hex$(m_Number), 8) & "<BR>" & _
                "<B>Error call stack</B><BR>" & ReplaceString(ReplaceHTMLMetacharactersEx(m_Source), ";", "<BR>")
  Logged = True
End Property

Public Property Get Logged() As Boolean
  Logged = m_Logged
End Property

Friend Property Let Logged(ByVal NewValue As Boolean)
  m_Logged = NewValue
End Property

' MPS
Public Property Get ErrorDescriptionXML() As String
  ErrorDescriptionXML = ReplaceHTMLMetacharactersEx(Description)
  Logged = True
End Property

'cadxx this was agreed to be lower case elements and attributes
Public Property Get ErrorDetailsXML() As String
  Dim v As Variant, src As String, i As Long
  
  ErrorDetailsXML = "<number>Error Number: 0x" & Right$("00000000" & Hex$(Number), 8) & "</number>"
  ' needs to parse each source individually
  v = Split(ReplaceHTMLMetacharactersEx(Source), ";")
  For i = 0 To UBound(v)
    src = src & "<source>" & XMLTextEx(CStr(v(i))) & "</source>"
  Next i
  ErrorDetailsXML = ErrorDetailsXML & "<callstack>"
  If Len(src) > 0 Then
    ErrorDetailsXML = ErrorDetailsXML & src
  Else
    ErrorDetailsXML = ErrorDetailsXML & "<source></source>"
  End If
  ErrorDetailsXML = ErrorDetailsXML & "</callstack>"
  Logged = True
End Property



