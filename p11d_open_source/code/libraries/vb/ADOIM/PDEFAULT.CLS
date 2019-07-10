VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDefault"
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
  IParseItem_ActionStr = "DEFAULT"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ispec As ImportSpec, DefSet As Boolean, d0 As Date
  Dim errtext As String
  
  Set ispec = objCurItem
  On Error GoTo default_err
  ispec.DefaultValue = GetTypedValue(params(1), ispec.vartype)
  IParseItem_OnAction = PARSESTATUS_OK
  Exit Function
  
default_err:
  errtext = Err.Description
  Err.Clear
  Call Err.Raise(ERR_PARSETOKEN, "ParseDefault", "Unable to Parse token DEFAULT. " & vbCrLf & errtext & vbCrLf & "Field: " & ispec.DestField)
End Function



