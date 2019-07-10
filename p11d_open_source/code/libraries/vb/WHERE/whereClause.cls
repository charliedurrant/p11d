VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "whereClause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As whereClause
Public Operator As TCSWHERE_LOGICAL_OPERATOR

Public LHTree As whereClause
Public RHTree As whereClause
Public Value As whereCondition

Public Sub Kill()
  If Not LHTree Is Nothing Then
    Call LHTree.Kill
    Set LHTree = Nothing
  End If
  If Not RHTree Is Nothing Then
    Call RHTree.Kill
    Set RHTree = Nothing
  End If
  Set Value = Nothing
  Set Parent = Nothing
End Sub

Public Function OutputLogic() As String
  Dim s As String
  
  If Operator = LOGICAL_AND Then
    s = "AND("
  ElseIf Operator = LOGICAL_OR Then
    s = "OR("
  End If
  If Not LHTree Is Nothing Then
    s = s & LHTree.OutputLogic
  End If
  If Not RHTree Is Nothing Then
    s = s & RHTree.OutputLogic
  End If
  If Not Value Is Nothing Then
    s = s & Value.Name
  End If
  OutputLogic = s
End Function


