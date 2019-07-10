VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "FACTOR"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  
  If (ispec.vartype <> TYPE_DOUBLE) And (ispec.vartype <> TYPE_LONG) Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseFactor", "Unable to Parse token FACTOR as Import item is not a number." & vbCrLf & "Field: " & ispec.DestField)
  End If
  If ispec.varaction <> [_ACTION_NOACTION] Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseFactor", "Unable to Parse token FACTOR as Import item already has an action assigned to it." & vbCrLf & "Field: " & ispec.DestField)
  End If
  IParseItem_OnAction = PARSESTATUS_OK
  ispec.varaction = ACTION_FACTOR
  ispec.actionparam = CStr(params(1))
End Function



