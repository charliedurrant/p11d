VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseFatal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "FATAL"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As TCSPARSE.PARSESTATUS
  Dim ispec As ImportSpec
  
  Set ispec = objCurItem
  ispec.constraint = ispec.constraint + CONSTRAINT_FATAL
  IParseItem_OnAction = PARSESTATUS_OK
End Function


