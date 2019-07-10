VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseStaticX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "STATICX"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Call xSet("ParsePop")
  
  Call DisplayText(pre)
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    ReportControl.FirstX = Printer.CurrentX - ReportControl.LeftMargin
  ElseIf (ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE) Or (ReportControl.rTarget = RPT_CONFIG) Then
    ReportControl.FirstX = ReportControl.Preview.CurrentX - ReportControl.LeftMargin
  End If
  Call xReturn("ParsePop")
End Function




