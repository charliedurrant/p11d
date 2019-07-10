VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "whereExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mRootClause As whereClause

Public Sub RebaseConditions(Optional ByVal FirstIndex As Long = vbKeyA)
  If Not mRootClause Is Nothing Then Call mRootClause.RebaseConditions(FirstIndex)
End Sub

Public Function Logic() As String
  If Not mRootClause Is Nothing Then
    Logic = mRootClause.OutputLogic
  End If
End Function

Public Function SQL() As String
  If Not mRootClause Is Nothing Then
    SQL = mRootClause.OutputSQL
  End If
End Function

Public Sub MergeClauses(NewClause As whereClause, ByVal Operator As TCSWHERE_LOGICAL_OPERATOR)
  Dim NewRootClause As whereClause
  
  If Not mRootClause Is Nothing Then
    Set NewRootClause = New whereClause
    Set NewRootClause.RHTree = mRootClause
    Set NewRootClause.LHTree = NewClause
    NewRootClause.Operator = Operator
  Else
    Set NewRootClause = NewClause
  End If
  Set mRootClause = NewRootClause
End Sub
