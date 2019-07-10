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
End Sub


Public Function RebaseConditions(ByVal PrepareOnly As Boolean, Optional ByVal NextIndex As String = "A") As String
  If Not LHTree Is Nothing Then
    NextIndex = LHTree.RebaseConditions(PrepareOnly, NextIndex)
  End If
  If Not RHTree Is Nothing Then
    NextIndex = RHTree.RebaseConditions(PrepareOnly, NextIndex)
  End If
  If Not Value Is Nothing Then
    If PrepareOnly Then
      Value.Name = ""
    ElseIf Len(Value.Name) = 0 Then
      Value.Name = NextIndex
      NextIndex = NextClauseIndex(NextIndex)
    End If
  End If
  RebaseConditions = NextIndex
End Function

Public Property Get InternalConditions() As String
  Dim s As String
  
  If Not LHTree Is Nothing Then
    s = LHTree.InternalConditions
  End If
  If Not RHTree Is Nothing Then
    s = s & RHTree.InternalConditions
  End If
  If Not Value Is Nothing Then
    s = s & CONDITION_FN & Value.Name & CLAUSE_SEP & Value.Field & CLAUSE_SEP & Value.DataType & CLAUSE_SEP & Value.Operator & CLAUSE_SEP & VALUE_BEGIN & GetValueAsString(Value.Value, Value.DataType) & VALUE_END & CLAUSE_END
  End If
  InternalConditions = s
End Property

Public Property Get Internal() As String
  Dim s As String
  
  If Operator = LOGICAL_AND Then
    s = "AND("
  ElseIf Operator = LOGICAL_OR Then
    s = "OR("
  End If
  If Not LHTree Is Nothing Then
    s = s & LHTree.Internal & CLAUSE_SEP
  End If
  If Not RHTree Is Nothing Then
    s = s & RHTree.Internal & CLAUSE_END
  End If
  If Not Value Is Nothing Then
    s = s & Value.Name
  End If
  Internal = s
End Property

Public Property Get OutputSQL(ByVal UsePlaceholders As Boolean) As String
  Dim lhs As String, rhs As String
  Dim s As String
  
  If Not Value Is Nothing Then
    If UsePlaceholders Then
      OutputSQL = Value.Name
    Else
      OutputSQL = Value.ConditionSQL
    End If
  Else
    If Not LHTree Is Nothing Then
      If UsePlaceholders Then
        lhs = "(" & LHTree.OutputSQL(UsePlaceholders) & ")"
      Else
        If LHTree.Operator = LOGICAL_AND Then
          lhs = "(" & LHTree.OutputSQL(UsePlaceholders) & ")"
        Else
          lhs = LHTree.OutputSQL(UsePlaceholders)
        End If
      End If
    End If
    If Not RHTree Is Nothing Then
      If UsePlaceholders Then
        rhs = "(" & RHTree.OutputSQL(UsePlaceholders) & ")"
      Else
        rhs = RHTree.OutputSQL(UsePlaceholders)
      End If
    End If
    If UsePlaceholders Then
      If Operator = LOGICAL_AND Then
        OutputSQL = lhs & " AND " & rhs
      ElseIf Operator = LOGICAL_OR Then
        OutputSQL = lhs & " OR " & rhs
      End If
    Else
      If Operator = LOGICAL_AND Then
        If LHTree.LHTree Is Nothing Then
          OutputSQL = lhs & " AND " & rhs
        Else
          OutputSQL = "(" & lhs & ") AND " & rhs
        End If
      ElseIf Operator = LOGICAL_OR Then
        OutputSQL = lhs & " OR " & rhs
      End If
    End If
  End If
End Property

Public Sub OutputConditionList(lb As ListBox, ByVal PrepareOutput As Boolean)
  Dim s As String
  
  If Not LHTree Is Nothing Then
    Call LHTree.OutputConditionList(lb, PrepareOutput)
  End If
  If Not RHTree Is Nothing Then
    Call RHTree.OutputConditionList(lb, PrepareOutput)
  End If
  If Not Value Is Nothing Then
    If PrepareOutput Then
      Value.OutputDone = False
    ElseIf Not Value.OutputDone Then
      Call lb.AddItem(Value.Name & vbTab & Value.ConditionString)
      Value.OutputDone = True
    End If
  End If
End Sub


