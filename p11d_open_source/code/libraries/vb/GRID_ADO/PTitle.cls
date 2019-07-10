VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "TITLE"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

'* send a report format at the start of this column
Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim AClass As AutoClass

  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  If Not objCurItem Is Nothing Then Call Err.Raise(ERR_AUTOPARSE, "ParseTitle", "Title can only be set as a Global for the complete report using AddReportFormat")
  Set AClass = psett.ParseParameter(2)
  AClass.ReportHeader = params(1)
  IParseItem_OnAction = PARSESTATUS_OK
End Function

