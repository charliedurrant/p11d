VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseKEEPTOGETHER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "KEEPTOGETHER"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim fh As Single, CurrentY As Single
  
  If Len(pre) > 0 Then Err.Raise ERR_KEEPTOGETHER, "KeepTogether", "The KEEPTOGETHER token must be used in a seperate out statement."
  If (ReportControl.SuppressNewPageCalc) Or (ReportControl.rTarget > RPT_EXPORT) Then GoTo ParseKeepTogether_end:
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREPARE) Then
    Call CheckForNewPage(ReportControl.fStyle.FontHeight * CSng(params(1)), True, False)
  End If
ParseKeepTogether_end:
  ReportControl.NoRecord = True
End Function

