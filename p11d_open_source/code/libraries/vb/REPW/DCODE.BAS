Attribute VB_Name = "DebugCode"
Option Explicit

#If DEBUGVER Then
Public Sub OutputDebug(ByVal FunctionName As String, ByVal DebugString As String)
  Debug.Print (FunctionName & ":: " & DebugString)
End Sub
#End If
