VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortAscending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISortFunction

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  If v0 < v1 Then
    ISortFunction_CompareItems = -1
  ElseIf v0 > v1 Then
    ISortFunction_CompareItems = 1
  Else
    ISortFunction_CompareItems = 0
  End If
End Function
