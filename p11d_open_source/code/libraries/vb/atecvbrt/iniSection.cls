VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iniSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public SectionName As String
Private m_inikeys As Dictionary

Private Sub Class_Initialize()
  Set m_inikeys = New Dictionary
  m_inikeys.CompareMode = TextCompare
End Sub

Private Sub kill()
  If Not m_inikeys Is Nothing Then
    Call m_inikeys.RemoveAll
    Set m_inikeys = Nothing
  End If
End Sub

Private Sub Class_Terminate()
  Call kill
End Sub

Public Sub AddKey(ByRef line As String)
  Dim p0 As Long, Key As String, Value As String
  
  p0 = InStr(1, line, "=", vbBinaryCompare)
  If p0 = 0 Then Err.Raise ERR_INIFILECACHE, "iniSection.AddKey", "Invalid ini file line: " & line
  Key = Trim$(Mid$(line, 1, p0 - 1))
  Value = Trim$(Mid$(line, p0 + 1))
  If Mid$(Value, 1, 1) = """" Then
    If Right$(Value, 1) = """" Then Value = Mid$(Value, 2, Len(Value) - 2)
  End If
  If m_inikeys.Exists(Key) Then Err.Raise ERR_INIFILECACHE, "iniSection.AddKey", "Duplicate key value in ini file line: " & line
  Call m_inikeys.Add(Key, Value)
End Sub

Public Function GetIniEntry(ByVal Key As String, ByRef Default As String) As String
  Dim ins As iniSection
  
  On Error GoTo GetIniEntry_err
  Key = Trim$(Key)
  If Not m_inikeys.Exists(Key) Then
    GetIniEntry = Default
  Else
    GetIniEntry = m_inikeys.Item(Key)
  End If
  Exit Function
  
GetIniEntry_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "iniSection.GetIniEntry"), Err.Description
End Function


