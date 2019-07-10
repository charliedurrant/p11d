VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "BOX"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Call xSet("ParseBox")
  
  Call DisplayText(pre)
  If IsExportHTML Then
    Call HTMLBox("", ReportControl.HTML.FillColor, "#ffffff", CSng(params(2)), CSng(params(1)), ALIGN_GENERAL, InStr(1, params(3), "F", vbTextCompare) <> 0)
  Else
    Select Case paramcount
      Case 2
        Call xBox(CSng(params(1)), CSng(params(2)), False)
      Case 3
        Call xBox(CSng(params(1)), CSng(params(2)), InStr(1, params(3), "F", vbTextCompare) <> 0)
      Case Else
        Call Err.Raise(ERR_PARSEBOX, "ParseBox", "Expected 2 or 3 parameters found " & CStr(paramcount))
    End Select
  End If
  
  Call xReturn("ParseBox")
End Function


