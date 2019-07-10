VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseReqNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit  'MPS99
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "REQUIRENUMERIC"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As TCSPARSE.PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  If (ispec.vartype <> TYPE_DOUBLE) And (ispec.vartype <> TYPE_LONG) Then Call Err.Raise(ERR_PARSETOKEN, "ParseReqNum", "Unable to Parse token REQUIRENUMERIC as Import item is not a Number." & vbCrLf & "Field: " & ispec.DestField)
  ispec.RequireNumeric = True
  ispec.Constraint = (ispec.Constraint Xor CONSTRAINT_REQUIRENUMERIC) Or CONSTRAINT_REQUIRENUMERIC
  IParseItem_OnAction = PARSESTATUS_OK
End Function


