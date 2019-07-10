VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepParseSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseSettings


Private Sub IParseSettings_EndParseLine(objCurItem As Variant, pstatus As PARSESTATUS)
  ' none
End Sub

Private Property Set IParseSettings_ParseParameter(i As Long, RHS As Variant)
  ' none
End Property

Private Property Get IParseSettings_ParseParameter(i As Long) As Variant
  ' none
End Property

Private Property Get IParseSettings_ParserType() As PARSETYPE
  IParseSettings_ParserType = PARSETYPE_TOKEN
End Property

Private Function IParseSettings_StartParseLine() As Variant
  Set IParseSettings_StartParseLine = Nothing
End Function

Private Property Get IParseSettings_TOKENCLOSE() As String
  IParseSettings_TOKENCLOSE = "}"
End Property

Private Property Get IParseSettings_TOKENOPEN() As String
  IParseSettings_TOKENOPEN = "{"
End Property
