VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Private mCaption As String
 
Public Width As String
Public Sum As Boolean
Public Count As Boolean
Public Group As Integer
Public GroupLastValue As String

Private mSumToDate As Long
Private mCountToDate As Integer
Private mCharsToDate As Integer ' used to calculate relative widths for columns automatically

Private Sub Class_Initialize()
  Name = ""
  mCaption = ""
  Width = ""
  Sum = False
  Count = False
  Group = 0
  GroupLastValue = ""
End Sub

Public Property Let Caption(sNewVal As String)
  mCaption = sNewVal
End Property


Public Property Get Caption() As String
  If Len(mCaption) = 0 Then
    Caption = Name
  Else
    Caption = mCaption
  End If
End Property

' Retrieve the final figure
Public Property Get SumTotal() As Long
  SumTotal = mSumToDate
End Property

' Retrieve the final figure
Public Property Get CountTotal() As Integer
  CountTotal = mCountToDate
End Property

' Retrieve the final figure
Public Property Get CharsTotal() As Integer
  CharsTotal = mCharsToDate
End Property

Public Sub SumAdd(ByVal lAdd As Long)
  mSumToDate = mSumToDate + lAdd
End Sub

Public Sub CountAdd(ByVal iAdd As Integer)
  mCountToDate = mCountToDate + iAdd
End Sub

Public Sub CharsAdd(ByVal iAdd As Integer)
  mCharsToDate = mCharsToDate + iAdd
End Sub

