Attribute VB_Name = "lows"
Option Explicit

Public Function isNonBlank(ByVal v As Variant) As Boolean
  isNonBlank = False
  If Not IsNull(v) Then
    isNonBlank = (Len(v) > 0)
  End If
End Function


