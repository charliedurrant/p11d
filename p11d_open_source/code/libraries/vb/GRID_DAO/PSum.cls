VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "SUM"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim AClass As AutoClass
  Dim aCol As AutoCol
    
  Set aCol = objCurItem
  Set AClass = psett.ParseParameter(2)
  If AClass.CheckDataType And Not IsNumberField(aCol.dbDataType) Then Call Err.Raise(ERR_AUTOPARSE, "ParseSum", "Cannot sum on a non numeric field")
  If FormatRemove Then
    aCol.SumType = TYPE_NOSUM
  Else
    aCol.SumType = TYPE_SUM
  End If
  IParseItem_OnAction = PARSESTATUS_OK
End Function


