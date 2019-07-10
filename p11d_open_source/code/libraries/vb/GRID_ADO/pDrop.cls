VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'*  {DROP=DisplayValue,DBValue}
'*  GIVES A DROPDOWN BOX WITH THE GIVEN ITEMS IN IT
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "DROP"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol
  Dim i As Long, n As Long
  
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  Set acol = objCurItem
  n = 2 * (paramcount \ 2) ' Integer division
  If paramcount <> n Then Call Err.Raise(ERR_AUTOPARSE, "ParseDrop", "Combo Box drop downs must be in value pairs (DisplayName,Value). Fieldname=" & acol.DataField)
  acol.DropType = DROPDOWN_LIST
  acol.BoundDataField = acol.DataField
  For i = 1 To paramcount
    Call acol.AddParam(params(i))
  Next i
  IParseItem_OnAction = PARSESTATUS_OK
End Function

