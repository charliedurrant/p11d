VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "MASK"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  If (ispec.vartype <> TYPE_STR) Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseMask", "Unable to Parse token MASK as Import item is not a string." & vbCrLf & "Field: " & ispec.DestField)
  End If
  If (ispec.Constraint And CONSTRAINT_MASK) = CONSTRAINT_MASK Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseMask", "Unable to Parse token MASK as this import item already has a MASK property." & vbCrLf & "Field: " & ispec.DestField)
  End If
  
  ispec.Constraint = ispec.Constraint + CONSTRAINT_MASK
  ispec.Mask = CStr(params(1))
  IParseItem_OnAction = PARSESTATUS_OK
End Function

