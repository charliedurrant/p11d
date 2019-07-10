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
  m_Number = Number
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

Friend Property Get URL(ByVal ErrorIndex As Long) As String
  URL = "NUMBER" & ErrorIndex & "=" & Number & _
              "&DESCRIPTION" & ErrorIndex & "=" & URLEncode(Description) & _
              "&SOURCE" & ErrorIndex & "=" & URLEncode(Source)
  m_Logged = True
End Property

Public Property Get DescriptionHTML() As String
  DescriptionHTML = ReplaceString(m_Description, vbCrLf, "<BR>")
  m_Logged = True
End Property

Public Property Get DetailsHTML() As String
  DetailsHTML = "<BR>Error Number: 0x" & Right$("00000000" & Hex$(m_Number), 8) & "<BR>" & _
                "<B>Error call stack</B><BR>" & ReplaceString(m_Source, ";", "<BR>")
  Logged = True
End Property

Public Property Get Logged() As Boolean
  Logged = m_Logged
End Property

Friend Property Let Logged(ByVal NewValue As Boolean)
  m_Logged = NewValue
End Property

Public Property Get ErrorDescriptionXML() As String
  ErrorDescriptionXML = Description
  Logged = True
End Property

Public Property Get ErrorDetailsXML() As String
  ErrorDetailsXML = "<NUMBER>Error Number: 0x" & Right$("00000000" & Hex$(Number), 8) & "</NUMBER>"
  ErrorDetailsXML = ErrorDetailsXML & "<CALLSTACK><SOURCE><![CDATA[" & ReplaceString(Source, ";", "]]></SOURCE><SOURCE><![CDATA[") & "]]></SOURCE></CALLSTACK>"
  Logged = True
End Property

