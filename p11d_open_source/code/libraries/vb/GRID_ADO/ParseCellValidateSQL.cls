VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseBeforeUpdateAddNewSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "BEFOREUPDATEADDNEWSQL"
End Property

Private Function IParseItem_OnAction(psett As TCSPARSE.IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol, BoolFormat As String
    
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  If (paramcount < 1) Then Call Err.Raise(ERR_AUTOPARSE, "ParseBeforeUpdateAddNewSQL", "Expected 1 parameter found " & CStr(paramcount))
  Set acol = objCurItem
  acol.BeforeUpdadeAddNewSQL = params(1)
  IParseItem_OnAction = PARSESTATUS_OK
End Function



