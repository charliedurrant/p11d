VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseCopyField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 2
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "COPYFIELD"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  ispec.DestField = params(1)
  ispec.CopyFieldKey = CLng(params(2))
  ispec.key = KEY_COPIEDFIELD
  'IParseItem_OnAction = PARSESTATUS_ABORTLINE
  IParseItem_OnAction = PARSESTATUS_OK
End Function

