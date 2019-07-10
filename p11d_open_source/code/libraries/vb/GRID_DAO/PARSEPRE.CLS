VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParsePrecision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "PRECISION"
End Property

Private Function IParseItem_OnAction(psett As TCSPARSE.IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As TCSPARSE.PARSESTATUS
  Dim ACol As AutoCol
  
  On Error GoTo IParseItem_OnAction_err
  Set ACol = objCurItem
  ACol.Precision = CLng(params(0))
  IParseItem_OnAction = PARSESTATUS_OK
  
IParseItem_OnAction_end:
  Exit Function
  
IParseItem_OnAction_err:
  IParseItem_OnAction = PARSESTATUS_ABORTSTATEMENT
  Resume IParseItem_OnAction_end
End Function
