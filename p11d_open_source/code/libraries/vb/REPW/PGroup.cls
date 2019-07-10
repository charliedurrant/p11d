VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "GROUP"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim rFld As ReportField
  
  Set rFld = objCurItem
  rFld.Group = True
  If (paramcount <> 0) And (paramcount <> 1) Then Call Err.Raise(ERR_REPWPARSE, "ParseGroup", "Expected either 0 or 1 parameters found " & CStr(paramcount))
  If paramcount = 1 Then
    If InStr(1, params(1), "T", vbTextCompare) > 0 Then rFld.GroupTotal = True
    If InStr(1, params(1), "P", vbTextCompare) > 0 Then rFld.GroupPageBreak = True
    If InStr(1, params(1), "H", vbTextCompare) > 0 Then rFld.GroupHeader = True
    If InStr(1, params(1), "C", vbTextCompare) > 0 Then rFld.GroupCount = True
  End If
  IParseItem_OnAction = PARSESTATUS_OK
End Function



