VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseYRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "YREL"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  
  On Error GoTo ParseYRel_err
  Call DisplayText(pre)
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    Printer.CurrentY = GetRelative(Printer.CurrentY, CSng(params(1)), ReportControl.PageHeight)
  ElseIf (ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE) Or (ReportControl.rTarget = RPT_CONFIG) Or (ReportControl.rTarget = RPT_PREPARE) Then
    ReportControl.Preview.CurrentY = GetRelative(ReportControl.Preview.CurrentY, CSng(params(1)), ReportControl.PageHeight)
  ElseIf IsExportHTML Then
    Call CloseDiv
    Call SetOpenDiv
    ReportControl.HTML.CurrentY = GetRelative(ReportControl.HTML.CurrentY, CSng(params(1)), ReportControl.PageHeight)
  End If
ParseYRel_end:
  Exit Function
  
ParseYRel_err:
  If ReportControl.BeginSectionY <> -1 Then Resume ParseYRel_end
  Err.Raise Err.Number, ErrorSource(Err, "ParseYRel"), Err.Description
End Function



