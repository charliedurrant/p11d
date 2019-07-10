VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IParseSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get TOKENOPEN() As String
  ' returns the start token delimiter
  ' default={
End Property

Public Property Get TOKENCLOSE() As String
  ' returns the end token delimiter
  ' default=}
End Property

Public Property Get ParserType() As PARSETYPE
  ' returns the type of parser e.g Token parser or @Function() parser
End Property

Public Property Get ParseParameter(i As Long) As Variant
  ' returns the ParseParameter i, if 0 is the pparam value otherwise treat as an array
End Property

Public Property Set ParseParameter(i As Long, pparam As Variant)
  ' sets a ParseParameter i which will be accessible to the Start/EndParseline functions
  ' through the Get ParseParameter(i As Long)
End Property

Public Function StartParseLine() As Variant
  ' returns the object created for each new parse line
End Function

Public Sub EndParseLine(objCurItem As Variant, pstatus As PARSESTATUS)
  ' after each line is parsed calls this function with the object parsed and a parse status
End Sub

