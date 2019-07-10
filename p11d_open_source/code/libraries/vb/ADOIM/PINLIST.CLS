VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseInList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "INLIST"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec, i As Long
  
  Set ispec = objCurItem
  If ispec.vartype <> TYPE_STR Then Call Err.Raise(ERR_PARSETOKEN, "ParseInList", "Unable to Parse token INLIST as Import item is not a String." & vbCrLf & "Field: " & ispec.DestField)
  If (ispec.Constraint And CONSTRAINT_INLIST) = CONSTRAINT_INLIST Then Call Err.Raise(ERR_PARSETOKEN, "ParseInList", "Unable to Parse token INLIST as this import item already has a INLIST property." & vbCrLf & "Field: " & ispec.DestField)
  If (ispec.Constraint And CONSTRAINT_EXCLUDELIST) = CONSTRAINT_EXCLUDELIST Then Call Err.Raise(ERR_PARSETOKEN, "ParseInList", "Unable to Parse token INLIST as this import item already has a EXLIST property." & vbCrLf & "Field: " & ispec.DestField)
  ispec.Constraint = ispec.Constraint + CONSTRAINT_INLIST
  For i = 1 To paramcount
    ispec.ConstraintList.Add (CStr(params(i)))
  Next i
  IParseItem_OnAction = PARSESTATUS_OK
End Function


