VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseCollapseLike"
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
  IParseItem_ActionStr = "COLLAPSELIKE"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol
  Dim v As Variant, i As Long
    
  Set acol = objCurItem
  If (paramcount = 0) Or FormatRemove Then
    v = Empty
  Else
    ReDim v(1 To paramcount)
    For i = 1 To paramcount
      v(i) = params(i)
    Next i
  End If
  acol.CollapseLike = v
  IParseItem_OnAction = PARSESTATUS_OK
End Function
