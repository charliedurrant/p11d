VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportWizardParserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IParseSettings
Private mReportFields As Collection
Private mReportDetails As ReportDetails


Private Sub IParseSettings_EndParseLine(objCurItem As Variant, pstatus As PARSESTATUS)
  Dim rFld As ReportField
  
  Set rFld = objCurItem
  If pstatus = PARSESTATUS_OK Then
   Call mReportFields.Add(rFld, rFld.DataSetString & DATASET_SEP & rFld.Name)
  End If
End Sub

Private Property Get IParseSettings_ParserType() As PARSETYPE
  IParseSettings_ParserType = PARSETYPE_TOKEN
End Property

Private Function IParseSettings_StartParseLine() As Variant
  Set IParseSettings_StartParseLine = New ReportField
End Function

Private Property Get IParseSettings_TOKENCLOSE() As String
  IParseSettings_TOKENCLOSE = ")"
End Property

Private Property Get IParseSettings_TOKENOPEN() As String
  IParseSettings_TOKENOPEN = "("
End Property

Private Property Set IParseSettings_ParseParameter(i As Long, RHS As Variant)
  Select Case i
    Case 1
          Set mReportFields = RHS
    Case 2
          Set mReportDetails = RHS
    Case Else
         Call ECASE("ParseParameter Set - invalid offset")
  End Select
End Property

Private Property Get IParseSettings_ParseParameter(i As Long) As Variant
  Select Case i
    Case 1
          Set IParseSettings_ParseParameter = mReportFields
    Case 2
          Set IParseSettings_ParseParameter = mReportDetails
    Case Else
          Call ECASE("ParseParameter Get - invalid offset")
  End Select
End Property

