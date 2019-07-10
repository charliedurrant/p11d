VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseIgnoreNumConvErrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "IGNORE_CONVERSION_ERRORS"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  If Not IsNumberField(ispec.vartype) Then Call Err.Raise(ERR_PARSETOKEN, "ParseIgnoreNumConvErrs", "Unable to Parse token IGNORE_CONVERSION_ERRORS as Import item is not a Number." & vbCrLf & "Field: " & ispec.DestField)
  ispec.Constraint = ispec.Constraint Or CONSTRAINT_IGNORE_CONVERSIONERRORS
  IParseItem_OnAction = PARSESTATUS_OK
End Function


