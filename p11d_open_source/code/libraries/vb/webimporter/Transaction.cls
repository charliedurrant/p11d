VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ActiveTransactions As Integer

Public Enum TRANSACTION_TYPE
  TRANS_BEGIN = 1
  TRANS_ROLLBACK
  TRANS_COMMIT
End Enum


Public Function Transaction(Conn As Connection, ByVal Action As TRANSACTION_TYPE) As Boolean

On Error GoTo Transaction_err

  If Action = TRANS_BEGIN Then
    Conn.BeginTrans
    ActiveTransactions = ActiveTransactions + 1
  ElseIf Action = TRANS_COMMIT Then
    If ActiveTransactions > 0 Then
      Conn.CommitTrans
      ActiveTransactions = ActiveTransactions - 1
    End If
  ElseIf Action = TRANS_ROLLBACK Then
    If ActiveTransactions > 0 Then
        Conn.RollbackTrans
        ActiveTransactions = ActiveTransactions - 1
    End If
  Else
    'Err.Raise 380
  End If
  Transaction = True
Transaction_end:
  Exit Function
  
Transaction_err:
  Transaction = False
  Err.Raise ERR_INVALID_TRANSACTION, ErrorSource(Err, "Transaction.Transaction"), Err.Description & vbCrLf & "Error in transaction"
  Resume Transaction_end
End Function

