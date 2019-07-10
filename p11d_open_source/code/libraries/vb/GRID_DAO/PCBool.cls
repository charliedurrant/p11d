VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseCBoolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'* {CBOOLEAN}
'* SETS DATATYPE TO BE BOOLEAN
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "CBOOLEAN"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim aCol As AutoCol, BoolFormat As String
    
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  If Not ((paramcount = 0) Or (paramcount = 2)) Then Err.Raise ERR_AUTOPARSE, "ParseCBoolean", "Expected either 0 or 2 parameters found " & CStr(paramcount)
  Set aCol = objCurItem
  aCol.dbDataType = TYPE_LONG
  'aCol.dbDataType = TYPE_BOOL
  BoolFormat = "True/False"
  If paramcount = 2 Then BoolFormat = params(1) & "/" & params(2)
  aCol.IsCheckBox = False
  aCol.CustomFormat = True
  aCol.Format = BoolFormat
  aCol.SysNoEdit = True
  aCol.NoAddNew = True
  IParseItem_OnAction = PARSESTATUS_OK
End Function



