VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseLeftMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "LEFTMARGIN"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim lm As Single, pWidth As Single
  
  Call DisplayText(pre)
  lm = CSng(params(1))
  If (lm < 0) Or (lm > 49) Then Err.Raise ERR_MARGIN, "LeftMargin", "The left margin of the page must be between 1 and 49 percent"
  pWidth = ReportControl.PageWidth + ReportControl.LeftMargin - 1 + ReportControl.RightMargin
  lm = GetPercent(lm, pWidth)
  If lm = 0 Then lm = 1
  ReportControl.LeftMargin = lm
  ReportControl.PageWidth = pWidth - lm + 1 - ReportControl.RightMargin
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    If Printer.CurrentX <> ReportControl.LeftMargin Then Printer.CurrentX = ReportControl.LeftMargin
  ElseIf (ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE) Or (ReportControl.rTarget = RPT_CONFIG) Then
    If ReportControl.Preview.CurrentX <> ReportControl.LeftMargin Then ReportControl.Preview.CurrentX = ReportControl.LeftMargin
  End If
End Function


