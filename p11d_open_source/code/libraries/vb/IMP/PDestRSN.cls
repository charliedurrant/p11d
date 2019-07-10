VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDestRSName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'MPSAUG
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 2
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "DESTRSNAME"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim icls As ImportClass
  
  'MPSAUG
  Set icls = psett.ParseParameter(2)
  If icls.UseCFGDestNameIfAvailable Then
    icls.DestRSName = CStr(params(1))
  End If
  If icls.UseCFGUpdateTypeIfAvailable Then
    icls.UpdateType = CStr(params(2))
  End If
  IParseItem_OnAction = PARSESTATUS_INVALID
End Function

