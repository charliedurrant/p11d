Attribute VB_Name = "Errors"
Option Explicit

Public Function ErrorSourceEx(ByVal ErrObj As ErrObject, ByVal FunctionName As String) As String
  Dim Source As String

  If Not ErrObj Is Nothing Then Source = ErrObj.Source
  If Len(Source) = 0 Then
    ErrorSourceEx = FunctionName
  Else
    If InStr(1, Source, FunctionName, vbTextCompare) <> 1 Then
      ErrorSourceEx = FunctionName & ";" & Source
    Else
      ErrorSourceEx = Source
    End If
  End If
End Function

