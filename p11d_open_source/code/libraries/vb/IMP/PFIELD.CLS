VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1 '3
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "Field"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  If (paramcount <> 1) And (paramcount <> 3) Then Call Err.Raise(ERR_PARSETOKEN, "ParseField", "Unable to Parse token FIELD as 1 or 3 parameters expected" & vbCrLf & "Field: " & ispec.DestField)
  Set ispec = objCurItem
  
  ispec.key = KEY_FIELD
  ispec.DestField = params(1)
  If paramcount = 3 Then
    ispec.Offset = CLng(params(2))
    ispec.Length = CLng(params(3))
  End If
    
  IParseItem_OnAction = PARSESTATUS_OK
End Function
