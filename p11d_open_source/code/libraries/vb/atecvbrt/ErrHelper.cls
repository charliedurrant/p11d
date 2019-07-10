VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Stk_ErrorNumber() As Long
Private Stk_ErrorDescription() As String
Private Stk_ErrorSource() As String

Private StackMaxTop As Long
Private StackTop As Long

Public Sub Push(ByVal ErrObj As ErrObject)
  StackTop = StackTop + 1
  If StackTop > StackMaxTop Then
    StackMaxTop = StackMaxTop + 10
    ReDim Preserve Stk_ErrorNumber(1 To StackMaxTop) As Long
    ReDim Preserve Stk_ErrorDescription(1 To StackMaxTop) As String
    ReDim Preserve Stk_ErrorSource(1 To StackMaxTop) As String
  End If
  If Not ErrObj Is Nothing Then
    Stk_ErrorNumber(StackTop) = ErrObj.Number
    Stk_ErrorDescription(StackTop) = ErrObj.Description
    Stk_ErrorSource(StackTop) = ErrObj.Source
  Else
    Stk_ErrorNumber(StackTop) = 0
    Stk_ErrorDescription(StackTop) = ""
    Stk_ErrorSource(StackTop) = ""
  End If
End Sub

Public Sub Pop(ByVal ErrObj As ErrObject)
  If StackTop > 0 Then
    If Not ErrObj Is Nothing Then
      Call ErrObj.Clear
      ErrObj.Number = Stk_ErrorNumber(StackTop)
      ErrObj.Description = Stk_ErrorDescription(StackTop)
      ErrObj.Source = Stk_ErrorSource(StackTop)
    End If
    StackTop = StackTop - 1
  End If
End Sub
