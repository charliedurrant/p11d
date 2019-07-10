VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IParseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get ActionStr() As String
' This function returns the Action String
' i.e the item to be parsed
End Property

Public Property Get ActionParamCount() As Long
' This function returns the number of Parameters expected for this Action
' if <0 then Variable number of parameters - number parsed sent in paramcount variable
' note: parameters are comma separated values (delimited)
End Property

Public Function OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  ' This is the function called when this action string is parsed
  ' objCurItem is the currently active Parse Object created by the
  ' NewParseLine function in the ParseSettings Class
  ' returns 0 if success
  ' returns 1 if Parsing of the current object is to be aborted
  ' returns 2 if all Parsing is to be aborted
End Function

