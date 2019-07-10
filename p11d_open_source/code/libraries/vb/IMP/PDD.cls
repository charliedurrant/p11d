VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDateDelimiter"
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
  IParseItem_ActionStr = "DATEDELIMITER"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  If ispec.vartype <> TYPE_DATE Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseDateDelimiter", "Unable to Parse token DATEDELIMITER as Import item is not of type Date." & vbCrLf & "Field: " & ispec.DestField)
  End If
    If (ispec.varaction <> [_ACTION_NOACTION]) And (ispec.varaction <> ACTION_REFORMATDELIMITED) Then
    Call Err.Raise(ERR_PARSETOKEN, "ParseDateDelimiter", "Unable to Parse token DATEDELIMITER as Import item already has an action assigned to it." & vbCrLf & "Field: " & ispec.DestField)
  End If
  If (paramcount < 1) Or (paramcount > 2) Then Err.Raise ERR_PARSETOKEN, "ParseDateDelimiter", "Unable to Parse token DATEDELIMITER expected 1 or 2 parameters found " & paramcount
  ispec.varaction = ACTION_REFORMATDELIMITED
  ispec.DateDelimiter = params(1)
  If paramcount > 1 Then ispec.actionparam = params(2)
  IParseItem_OnAction = PARSESTATUS_OK
End Function



