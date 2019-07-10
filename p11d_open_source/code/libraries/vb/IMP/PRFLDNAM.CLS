VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseReadFieldNames"
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
  IParseItem_ActionStr = "READFIELDNAMES"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As TCSPARSE.PARSESTATUS
  Dim icls As ImportClass
  
  Set icls = psett.ParseParameter(2)
  icls.TakeFieldNamesFromFile = CBool(params(1))
  IParseItem_OnAction = PARSESTATUS_INVALID
End Function


