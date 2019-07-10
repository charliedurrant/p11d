VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseMinValue"
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
  IParseItem_ActionStr = "MINVALUE"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  If (ispec.vartype <> TYPE_DOUBLE) And (ispec.vartype <> TYPE_LONG) And (ispec.vartype <> TYPE_DATE) Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseMinValue", "Unable to Parse token MINVALUE as Import item is not a Number or a Date." & vbCrLf & "Field: " & ispec.DestField)
  End If
  If (ispec.Constraint And CONSTRAINT_MIN) = CONSTRAINT_MIN Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseMinValue", "Unable to Parse token MINVALUE as this import item already has a MINVALUE property." & vbCrLf & "Field: " & ispec.DestField)
  End If
  ispec.Constraint = ispec.Constraint + CONSTRAINT_MIN
  If ispec.vartype = TYPE_DATE Then
    ispec.minvalue = ConvertDate(params(1), CONVERT_FIXEDDATE, "DDMMYYYY")
  Else
    ispec.minvalue = params(1)
  End If
  IParseItem_OnAction = PARSESTATUS_OK
End Function


