VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseNewColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'*  {NEWCOLUMN}
'*  THIS WOULD START A NEW COLUMN ON THE DIALOG SCREEN
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "NEWCOLUMN"
End Property

Private Function IParseItem_OnAction(psett As TCSPARSE.IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim aCol As AutoCol
  Dim AClass As AutoClass
  
  On Error GoTo IParseItem_OnAction_err
  Set aCol = objCurItem
  aCol.NewColumn = True
  Set AClass = psett.ParseParameter(2)
  AClass.lCols = AClass.lCols + 1
  IParseItem_OnAction = PARSESTATUS_OK
  
IParseItem_OnAction_end:
  Exit Function
  
IParseItem_OnAction_err:
  IParseItem_OnAction = PARSESTATUS_ABORTSTATEMENT
  Resume IParseItem_OnAction_end
End Function

