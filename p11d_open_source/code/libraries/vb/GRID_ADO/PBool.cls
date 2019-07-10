VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseBoolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'* {BOOLEAN}
'* SETS DATATYPE TO BE BOOLEAN
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "BOOLEAN"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol, BoolFormat As String
    
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  If (paramcount < 0) Or (paramcount > 2) Then Call Err.Raise(ERR_AUTOPARSE, "ParseBoolean", "Expected either 0,1 or 2 parameters found " & CStr(paramcount))
  Set acol = objCurItem
  acol.dbDataType = TYPE_BOOL
  BoolFormat = "True/False"
  If paramcount = 2 Then BoolFormat = params(1) & "/" & params(2)
  If paramcount = 1 Then BoolFormat = params(1)
  If Not StrComp(BoolFormat, "CheckBox", vbTextCompare) = 0 Then
    If Not ((StrComp(BoolFormat, "True/False", vbTextCompare) = 0) Or _
            (StrComp(BoolFormat, "Yes/No", vbTextCompare) = 0) Or _
            (StrComp(BoolFormat, "On/Off", vbTextCompare) = 0)) Then Call Err.Raise(ERR_AUTOPARSE, "ParseBoolean", "Expected either True/False, On/Off or Yes/No. Found " & BoolFormat)
    acol.IsCheckBox = False
    acol.Format = BoolFormat
  Else
    acol.IsCheckBox = True
    acol.Format = ""
  End If
  IParseItem_OnAction = PARSESTATUS_OK
End Function

