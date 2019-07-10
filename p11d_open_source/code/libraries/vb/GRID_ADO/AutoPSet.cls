VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoParseSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_AutoCols As Collection
Private m_AutoClass As AutoClass
Private m_AutoCol As AutoCol
Implements IParseSettings

Private Sub IParseSettings_EndParseLine(objCurItem As Variant, pstatus As PARSESTATUS)
 
'  Dim pCol As ParseCol
'
'  Set pCol = objCurItem
'  If pstatus = PARSESTATUS_OK Then
'    If m_AutoCols.Count > 0 Then
'      Call m_AutoCols.Add(pCol, , , m_AutoCols.Count)
'    Else
'      Call m_AutoCols.Add(pCol)
'    End If
'  End If

End Sub

Private Property Set IParseSettings_ParseParameter(i As Long, RHS As Variant)
  Select Case i
    Case 1
      Set m_AutoCols = RHS
    Case 2
      Set m_AutoClass = RHS
    Case 3
      Set m_AutoCol = RHS
    Case Else
      Call ECASE("ParseParameter Set - invalid offset")
  End Select
End Property

Private Property Get IParseSettings_ParseParameter(i As Long) As Variant
  Select Case i
    Case 1
      Set IParseSettings_ParseParameter = m_AutoCols
    Case 2
      Set IParseSettings_ParseParameter = m_AutoClass
    Case 3
      Set IParseSettings_ParseParameter = m_AutoCol
    Case Else
      Call ECASE("ParseParameter Get - invalid offset")
  End Select
End Property

Private Property Get IParseSettings_ParserType() As PARSETYPE
  IParseSettings_ParserType = PARSETYPE_TOKEN
End Property

Private Function IParseSettings_StartParseLine() As Variant
  Set IParseSettings_StartParseLine = m_AutoCol
End Function

Private Property Get IParseSettings_TOKENCLOSE() As String
  IParseSettings_TOKENCLOSE = "}"
End Property

Private Property Get IParseSettings_TOKENOPEN() As String
  IParseSettings_TOKENOPEN = "{"
End Property

