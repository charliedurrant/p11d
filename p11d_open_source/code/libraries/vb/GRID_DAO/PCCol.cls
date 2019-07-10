VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseUnboundColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'* Calculated unbound column in grid
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "UNBOUNDCOLUMN"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim aCol As AutoCol, s As String
  Dim CalcType As DERIVED_VALUES_CALC
  
  '{CALCCOLUMN=5} {NOEDIT}{NOADDNEW}{ONUPDATE=xxxx}{ONADDNEW=xxx}
  If FormatRemove Then Err.Raise ERR_NOREMOVEFORMAT, IParseItem_ActionStr, "Cannot remove " & IParseItem_ActionStr & " format"
  Set aCol = objCurItem
  s = params(1)
  CalcType = DerivedValueType(s)
  aCol.OnUpdateCalcValueType = CalcType
  aCol.OnUpdateCalcValue = s
  aCol.SysNoEdit = True
  aCol.NoAddNew = True
  aCol.UnboundColumn = True
  IParseItem_OnAction = PARSESTATUS_OK
End Function


