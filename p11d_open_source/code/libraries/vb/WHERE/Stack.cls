VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStack() As Variant
Private mStackMax As Long
Private mStackCur As Long
Private Const STACK_BASE As Long = 1
Private Const STACK_INCREMENT As Long = 256

Public Function IsEmpty() As Boolean
  IsEmpty = (mStackCur < STACK_BASE)
End Function

Private Function IsObject() As Boolean
  If mStackCur < STACK_BASE Then Err.Raise ERR_STACK_ISOBJECT, "IsObject", "Unable to determine whether the stack top is an object as the stack is empty"
  IsObject = (VarType(mStack(mStackCur)) = vbObject)
End Function

Public Sub Popv(vItem As Variant)
  If IsObject Then
    Set vItem = Me.Pop()
  Else
    vItem = Me.Pop
  End If
End Sub

Public Function Pop() As Variant
  If mStackCur < STACK_BASE Then Err.Raise ERR_STACK_POP, "Pop", "Unable to pop from stack as the stack is empty"
  If VarType(mStack(mStackCur)) = vbObject Then
    Set Pop = mStack(mStackCur)
    Set mStack(mStackCur) = Nothing
  Else
    Pop = mStack(mStackCur)
  End If
  mStack(mStackCur) = Empty
  mStackCur = mStackCur - 1
End Function

Public Function Push(ByVal vItem As Variant)
  mStackCur = mStackCur + 1
  If mStackCur > mStackMax Then
    mStackMax = mStackMax + STACK_INCREMENT
    ReDim Preserve mStack(1 To mStackMax)
  End If
  If VarType(vItem) = vbObject Then
    Set mStack(mStackCur) = vItem
  Else
    mStack(mStackCur) = vItem
  End If
End Function

Private Sub Class_Terminate()
  Dim v As Variant
  Do While Not Me.IsEmpty
    Call Popv(v)
    If VarType(v) = vbObject Then Set v = Nothing
  Loop
  mStackCur = 0
End Sub
