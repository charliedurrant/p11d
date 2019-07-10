VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseArial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "Arial"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

'* Fonts - {Fontname=PointSize,NBIUS}
'* Parameter 1
Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  
  Call xSet("ParseArial")
  
  Call DisplayText(pre)
  If paramcount > 2 Then Call Err.Raise(ERR_PARSEFONT, "ParseArial", "Expected < 2 parameters found " & CStr(paramcount))
  Call ProcessFontFormat("Arial", paramcount, 1, params)
  
  Call xReturn("ParseArial")
End Function
