VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AUTO_DOC_COMMENTS
  [_ADC_FIRST_ITEM] = 1
  ADC_DESCRIPTION = [_ADC_FIRST_ITEM]
  ADC_CATEGORY
  [_ADC_LAST_ITEM] = ADC_CATEGORY
End Enum

Public CommentType As AUTO_DOC_COMMENTS
Public Text As String

