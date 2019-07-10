VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fldChangeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFieldName As String
Private mbRequireNewField As Boolean
Private mChangedProperties As StringList
Private mNewFieldChecks As StringList

Public Function Setup(sFieldName As String)
  
  Set mChangedProperties = New StringList
  Set mNewFieldChecks = New StringList
  Call mNewFieldChecks.Add("Attributes")
  Call mNewFieldChecks.Add("Size")
  Call mNewFieldChecks.Add("Type")
  mFieldName = sFieldName
    
End Function

Public Sub AddChangedProperty(sPropertyName As String)
  If mNewFieldChecks.IsPresent(sPropertyName) Then
    mbRequireNewField = True
  End If
  mChangedProperties.Add (sPropertyName)
End Sub

Public Property Get RequireNewField() As Boolean
  RequireNewField = mbRequireNewField
End Property
Public Property Get ChangedProperties() As StringList
  Set ChangedProperties = mChangedProperties
End Property
Public Property Get Name() As String
  Name = mFieldName
End Property

Private Sub Class_Terminate()
  Set mChangedProperties = Nothing
  Set mNewFieldChecks = Nothing
End Sub
