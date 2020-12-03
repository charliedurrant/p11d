VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RelocationDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Item As Variant
Public value As Variant
Public MadeGood As Variant
Public MadeGoodNet As Variant

Public Qualify As Variant
Public Comments As Variant
Public benefit As Variant
Public Key As Long
Public IsTaxDeducted As Boolean

Public Function Calculate() As Variant
  On Error GoTo Calculate_ERR
  
  Call xSet("Calculate")
'MP DB ToDo - reason for setting below vars to Error first?
  Calculate = S_ERROR
  benefit = S_ERROR
  
  MadeGoodNet = Min(value, MadeGood)
  Calculate = value - MadeGoodNet
  benefit = Calculate
  
Calculate_END:
  Call xReturn("Calculate")
  Exit Function
Calculate_ERR:
  Calculate = S_ERROR
  Resume Calculate_END
End Function

