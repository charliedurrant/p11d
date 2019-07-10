VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseMean"
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
  IParseItem_ActionStr = "MEAN"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim acol As AutoCol
  Dim AClass As AutoClass
  
  Set acol = objCurItem
  Set AClass = psett.ParseParameter(2)
  If AClass.CheckDataType And Not IsNumberField(acol.dbDataType) Then Call Err.Raise(ERR_AUTOPARSE, "ParseMean", "Cannot calculate the mean of a non numeric field")
  If FormatRemove Then
    acol.SumType = TYPE_NOSUM
  Else
    acol.SumType = TYPE_MEAN
  End If
  IParseItem_OnAction = PARSESTATUS_OK
End Function




