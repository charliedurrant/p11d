VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseWBTEXTBOXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "WBTEXTBOXL"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 3
End Property

'* FillColor = RGB value
'* WB white background black text
Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ForeColorOld As Long, FillColorOld As Long
  Dim XWidth As Single, TWidth As Single, x0 As Single
  Dim YHeight  As Single, THeight  As Single, y0 As Single
  Call xSet("WBTEXTBOXL")

  Call DisplayText(pre)
  
  If IsExportHTML Then
    Call HTMLBox(CStr(params(3)), "#ffffff", "#000000", CSng(params(2)), CSng(params(1)), ALIGN_LEFT, False)
  Else
    ForeColorOld = ProcessForeColorRGB(0&) 'Black
    FillColorOld = ProcessFillColorRGB(COLOR_WHITE)
    
    Call DisplayTextBox(CStr(params(3)), CSng(params(1)), CSng(params(2)), False, ALIGN_LEFT)
    
    Call ProcessFillColorRGB(FillColorOld)
  End If
  Call xReturn("WBTEXTBOXL")
End Function

