VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseWingDings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "Wingdings"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

'* Fonts - {Fontname=PointSize,NBIUS}
'* Parameter 1
Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Call xSet("Wingdings")

  Call DisplayText(pre)
  If paramcount > 2 Then Call Err.Raise(ERR_PARSEFONT, "Wingdings", "Expected < 2 parameters found " & CStr(paramcount))
  Call ProcessFontFormat("Wingdings", paramcount, 1, params)
  
  Call xReturn("Wingdings")
End Function

