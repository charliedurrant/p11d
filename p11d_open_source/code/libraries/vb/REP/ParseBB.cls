VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "WB"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

'* FillColor = RGB value
'* WB white background black text
Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Call xSet("WB")

  Call DisplayText(pre)
  If IsExportHTML Then
    'Call HTMLBox(CStr(params(3)), "#ffffff", "#000000", CSng(params(2)), CSng(params(1)), ALIGN_CENTER, False)
  Else
    Call ProcessForeColorRGB(0&)           'Black
    Call ProcessFillColorRGB(COLOR_WHITE)  'White
  End If
  
  Call xReturn("WB")
End Function



