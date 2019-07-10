VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "DATASOURCE"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim iwiz As ImportWizard
  Dim icls As ImportClass
  Dim dName As String
  Dim CurrentdName As String
  
  dName = params(1)
  Set icls = psett.ParseParameter(2)
  Set iwiz = psett.ParseParameter(3)
  If Not iwiz Is Nothing Then
    Call iwiz.SetDataSource(dName)
    If Not iwiz.CurrentDest Is Nothing Then CurrentdName = iwiz.CurrentDest.Name
    If StrComp(CurrentdName, dName, vbTextCompare) <> 0 Then Call Err.Raise(ERR_PARSETOKEN, "ParseDataSource", "Unable to Parse token DATASOURCE as the dataset cannot be found." & vbCrLf & "Dataset: " & dName)
  End If
  icls.ImportName = dName
  IParseItem_OnAction = PARSESTATUS_ABORTSTATEMENT
End Function


