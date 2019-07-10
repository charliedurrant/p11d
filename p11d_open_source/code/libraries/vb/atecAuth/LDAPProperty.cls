VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LDAPProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mName As String
Private mDType As DATABASE_FIELD_TYPES
Private mMultiValued As Boolean
Private mValues As Variant
Private mValidValue As Boolean

Public Property Get Name() As String
  Name = mName
End Property

Friend Property Let Name(ByVal RHS As String)
  mName = RHS
End Property

Public Property Get MultiValued() As Boolean
  MultiValued = mMultiValued
End Property

Friend Property Let MultiValued(ByVal RHS As Boolean)
  mMultiValued = RHS
End Property

Public Property Get DType() As DATABASE_FIELD_TYPES
  DType = mDType
End Property

Friend Property Let DType(ByVal RHS As DATABASE_FIELD_TYPES)
  mDType = RHS
End Property

Public Property Get Values() As Variant
  Values = mValues
End Property

Friend Property Let Values(ByVal RHS As Variant)
  mValues = RHS
  mValidValue = True
End Property

Friend Sub ClearValue()
  mMultiValued = False
  mValues = vbEmpty
  mValidValue = False
End Sub

