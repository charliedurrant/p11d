VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iniFileCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_inifiles As Dictionary

Private Sub Class_Initialize()
  Set m_inifiles = New Dictionary
  m_inifiles.CompareMode = TextCompare
End Sub

Private Sub kill()
  If Not m_inifiles Is Nothing Then
    Call m_inifiles.RemoveAll
    Set m_inifiles = Nothing
  End If
End Sub

Private Sub Class_Terminate()
  Call kill
End Sub

Public Function GetIniEntry(ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal iniFilename As String) As String
  Dim ini As iniFile
  
  On Error GoTo GetIniEntry_err
  If Len(iniFilename) = 0 Then Err.Raise ERR_INIFILECACHE, "iniFileCache.GetIniEntry", "IniFile must be specified"
  If Not m_inifiles.Exists(iniFilename) Then
    Set ini = ReadIniFile(iniFilename)
  Else
    Set ini = m_inifiles.Item(iniFilename)
  End If
  GetIniEntry = ini.GetIniEntry(Section, Key, Default)
  Exit Function
  
GetIniEntry_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "iniFileCache.GetIniEntry"), Err.Description
End Function

Public Sub ClearCache()
  Call kill
  Set m_inifiles = New Dictionary
  m_inifiles.CompareMode = TextCompare
End Sub


Private Function ReadIniFile(ByRef iniFilename As String) As iniFile
  Dim fr As FileRead, buffer As String, ch As String
  Dim ini As iniFile, iniSection As iniSection
  On Error GoTo ReadIniFile_err
  
  Set ini = New iniFile
  ini.iniFilename = iniFilename
  Set fr = New FileRead
  If fr.OpenFile(iniFilename) Then
    Do While fr.GetLine(buffer)
      If Len(buffer) > 0 Then
        buffer = LTrim$(buffer)
        ch = Mid$(buffer, 1, 1)
        If ch = "[" Then
          Set iniSection = ini.AddSection(buffer)
        Else
          If iniSection Is Nothing Then Err.Raise ERR_INIFILECACHE, "iniFileCache.ReadIniFile", "No section found for line [" & buffer & "]"
          Call iniSection.AddKey(buffer)
        End If
      End If
    Loop
    Call fr.CloseFile
  End If
  Set fr = Nothing
  Call m_inifiles.Add(ini.iniFilename, ini)
  Set ReadIniFile = ini
  Exit Function
  
ReadIniFile_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "iniFileCache.ReadIniFile"), "Error reading iniFile [" & iniFilename & "]" & vbCrLf & Err.Description
  Resume
End Function


