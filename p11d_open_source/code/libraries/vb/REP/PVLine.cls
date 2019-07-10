VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseVLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "VLINE"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Call xSet("ParseVLine")
  
  Call DisplayText(pre)
  Select Case paramcount
    Case 0
      Call yLine(100!, False)
    Case 1
      Call yLine(CSng(params(1)), False)
    Case 2
      Call yLine(CSng(params(1)), InStr(1, params(2), "D", vbTextCompare) <> 0)
    Case Else
      Call Err.Raise(ERR_PARSELINE, "ParseVLine", "Expected < 2 parameters found " & CStr(paramcount))
  End Select
  
  
  Call xReturn("ParseVLine")
End Function
