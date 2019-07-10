VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseEndSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "ENDSECTION"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  
  If Len(pre) > 0 Then Err.Raise ERR_ENDSECTION, "EndSection", "The ENDSECTION token must be used in a seperate out statement."
  If ReportControl.BeginSectionY <> -1 Then
    If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
      Printer.CurrentX = ReportControl.BeginSectionX
      Printer.CurrentY = ReportControl.BeginSectionY
    ElseIf (ReportControl.rTarget = RPT_PREVIEW_DISPLAYPAGE) Or (ReportControl.rTarget = RPT_CONFIG) Or ((ReportControl.rTarget = RPT_PREPARE)) Then
      ReportControl.Preview.CurrentX = ReportControl.BeginSectionX
      ReportControl.Preview.CurrentY = ReportControl.BeginSectionY
    End If
    If ReportControl.BeginSectionY <> -1 Then
      ReportControl.BeginSectionX = -1!
      ReportControl.BeginSectionY = -1!
      ReportControl.OutputSection = True
    End If
  End If
  ReportControl.NoRecord = True
End Function




