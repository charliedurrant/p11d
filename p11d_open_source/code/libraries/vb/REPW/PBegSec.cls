VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseBeginSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "BEGINSECTION"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  
  If Len(pre) > 0 Then Err.Raise ERR_BEGINSECTION, "BeginSection", "The BEGINSECTION token must be used in a seperate out statement."
  If ReportControl.BeginSectionY <> -1 Then Err.Raise ERR_PARSEFONT, "BeginSection", "Cannot Start new report section as the report is currently in a section"
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    If ReportControl.PageTextPrinted Then
      ReportControl.BeginSectionX = Printer.CurrentX
      ReportControl.BeginSectionY = Printer.CurrentY
    End If
  ElseIf (ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE) Or (ReportControl.rTarget = RPT_CONFIG) Or (ReportControl.rTarget = RPT_PREPARE) Then
    If ReportControl.PageTextPrinted Then   'ReportControl.Preview.CurrentY > 1
      ReportControl.BeginSectionX = ReportControl.Preview.CurrentX
      ReportControl.BeginSectionY = ReportControl.Preview.CurrentY
    End If
  End If
  ReportControl.NoRecord = True
End Function




