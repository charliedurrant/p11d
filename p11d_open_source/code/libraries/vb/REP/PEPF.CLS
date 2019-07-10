VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseEndPFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "ENDPFOOTER"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Function IParseItem_OnAction(psett As TCSPARSE.IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As TCSPARSE.PARSESTATUS
  Call DisplayText(pre)
  If ReportControl.rTarget > RPT_EXPORT Then
    ReportControl.NoOutput = False
  End If
End Function

