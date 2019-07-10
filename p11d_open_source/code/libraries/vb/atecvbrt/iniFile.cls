VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public iniFilename As String
Private m_inisections As Dictionary

Private Sub Class_Initialize()
  Set m_inisections = New Dictionary
  m_inisections.CompareMode = TextCompare
End Sub

Private Sub kill()
  If Not m_inisections Is Nothing Then
    Call m_inisections.RemoveAll
    Set m_inisections = Nothing
  End If
End Sub

Private Sub Class_Terminate()
  Call kill
End Sub

Public Function GetIniEntry(ByRef Section As String, ByRef key As String, ByRef Default As String) As String
  Dim ins As iniSection
  
  On Error GoTo GetIniEntry_err
  If Not m_inisections.Exists(Section) Then
    GetIniEntry = Default
  Else
    Set ins = m_inisections.Item(Section)
    GetIniEntry = ins.GetIniEntry(key, Default)
  End If
  Exit Function
  
GetIniEntry_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "iniFile.GetIniEntry"), Err.Description
End Function


Public Function AddSection(ByVal SectionName As String) As iniSection
  Dim ins As iniSection
  
  SectionName = Trim$(SectionName)
  If Right$(SectionName, 1) <> "]" Then Err.Raise ERR_INIFILECACHE, "iniFile.AddSection", "Invalid SectionName found [" & SectionName & "]"
  SectionName = Mid$(SectionName, 2, Len(SectionName) - 2)
  If m_inisections.Exists(SectionName) Then Err.Raise ERR_INIFILECACHE, "iniFile.AddSection", "Duplicate SectionName found [" & SectionName & "]"
  Set ins = New iniSection
  ins.SectionName = SectionName
  Call m_inisections.Add(ins.SectionName, ins)
  Set AddSection = ins
End Function


