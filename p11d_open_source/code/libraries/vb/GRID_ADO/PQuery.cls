VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'*  {QUERY="SQL STATEMENT OR QUERY NAME HERE!",QUERY STYLE,FIELD NAME,PARAM}
'*  QUERY STYLE:  0=LIST
'*                1=EDITABLE COMBO
'*  FIELD NAME:   WHERE THE CHOSEN ELEMENT OF THE QUERY SHOULD GO
'*                LEAVE BLANK  IF CURRENT FIELD USED
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "QUERY"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol
  Dim NoFieldName As Boolean
  Dim i As Long
  
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  Set acol = objCurItem
  If paramcount < 2 Then Call Err.Raise(ERR_AUTOPARSE, "ParseQuery", "QUERY format must have at least 2 parameters. Fieldname=" & acol.DataField)
  acol.DropType = DROPDOWN_BOUND
  acol.QueryName = params(1)
  acol.QueryStyle = params(2)
  NoFieldName = (paramcount = 2)
  If (paramcount > 2) Then NoFieldName = (Len(params(3)) = 0)
  If NoFieldName Then
    acol.BoundDataField = acol.DataField
  Else
    acol.BoundDataField = params(3)
  End If
  For i = 4 To paramcount
    If Len(Trim$(params(i))) Then Call acol.AddParam(params(i))
  Next i
  IParseItem_OnAction = PARSESTATUS_OK
End Function

