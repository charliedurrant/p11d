VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseCriterion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 5
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "CRITERION"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim rFld As ReportField
  Dim Crit As Criterion, Index As Long
  
  Set rFld = objCurItem
  Set Crit = New Criterion
  Crit.Comparison = CLng(params(1))
  If Crit.Comparison = CRITERION_COMPARISON_UNDATED Or Crit.Comparison = CRITERION_COMPARISON_DATED Then
    Crit.Value = ""
  Else
    Crit.Value = GetTypedValue(params(2), rFld.DataType)
  End If
  If (Crit.Comparison = CRITERION_COMPARISON_BW) Or (Crit.Comparison = CRITERION_COMPARISON_BWEQ) Then
    Crit.Value2 = GetTypedValue(params(3), rFld.DataType)
  End If
  Crit.CaseSensitive = CBoolean(params(4))
  Index = CLng(params(5))
  Call rFld.Criteria.AddIndex(Crit, Index)
  IParseItem_OnAction = PARSESTATUS_OK
End Function


