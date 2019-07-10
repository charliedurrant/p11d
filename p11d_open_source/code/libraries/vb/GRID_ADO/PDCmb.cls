VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDropCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'*  {DROPCOMBO=DBValue1,DBValue2,DBValue3,}
'*  GIVES A DROPDOWN BOX WITH THE GIVEN ITEMS IN IT
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "DROPCOMBO"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol
  Dim i As Long
  
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  Set acol = objCurItem
  acol.DropType = DROPDOWN_COMBO
  acol.BoundDataField = acol.DataField
  For i = 1 To paramcount
    Call acol.AddParam(params(i))
  Next i
  IParseItem_OnAction = PARSESTATUS_OK
End Function


