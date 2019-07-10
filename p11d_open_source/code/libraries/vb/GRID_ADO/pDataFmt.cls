VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDateFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'*  {DATEFORMAT=DD/MM/YY}
'*  SETS DISPLAY FORMAT FOR DATES

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "DATEFORMAT"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim AClass As AutoClass
  Dim acol As AutoCol
  
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  Set acol = objCurItem
  Set AClass = psett.ParseParameter(2)
  If AClass.CheckDataType And (acol.dbDataType <> TYPE_DATE) Then Call Err.Raise(ERR_AUTOPARSE, "ParseDateFormat", "Cannot set a date format for a non-date field. Fieldname=" & acol.DataField)
  acol.Format = params(1)
  IParseItem_OnAction = PARSESTATUS_OK
End Function
