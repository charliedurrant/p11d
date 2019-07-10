VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

'*  {TAB="TAB TITLE HERE"}
'*  THIS WILL CREATE A NEW TAB ON A DIALOG SCREEN WITH THE
'*  GIVEN TITLE
Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "TAB"
End Property

Private Function IParseItem_OnAction(psett As TCSPARSE.IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim aCol As AutoCol
  Dim AClass As AutoClass
  
  On Error GoTo IParseItem_OnAction_err
  Set aCol = objCurItem
  aCol.TabTitle = params(1)
  Set AClass = psett.ParseParameter(2)
  AClass.lTabs = AClass.lTabs + 1
  IParseItem_OnAction = PARSESTATUS_OK
  
IParseItem_OnAction_end:
  Exit Function
  
IParseItem_OnAction_err:
  IParseItem_OnAction = PARSESTATUS_ABORTSTATEMENT
  Resume IParseItem_OnAction_end
End Function

