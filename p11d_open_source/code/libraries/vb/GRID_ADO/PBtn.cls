VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "BUTTON"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim l As Long
  Dim acol As AutoCol
    
  Set acol = objCurItem
  If paramcount > 1 Then Call Err.Raise(ERR_AUTOPARSE, "ParseButton", "Button format can only have one parameter, (picture number), Column=" & acol.DataField)
  acol.HasButton = Not FormatRemove
  If paramcount = 1 Then l = CLng(params(paramcount))
  acol.ButtonPictureIndex = BUTTON_BASE_BMP + l
  IParseItem_OnAction = PARSESTATUS_OK
End Function




