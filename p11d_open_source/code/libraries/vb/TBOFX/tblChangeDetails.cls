VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TblChangeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTableName As String
Private mChangedFields As Collection
Private mNewFields As StringList
Private mOldFields As StringList

Public Function Setup(sTableName As String) As Boolean
  
  Set mChangedFields = New Collection
  Set mNewFields = New StringList
  Set mOldFields = New StringList
 
  If Not Len(sTableName) = 0 Then
    mTableName = sTableName
    Setup = True
  End If
  
End Function

Public Function AddChangedField(sFieldName As String) As fldChangeDetails
  On Error GoTo AddChangedField_err
  
  If InCollection(mChangedFields, sFieldName) Then
    Set AddChangedField = New fldChangeDetails
    AddChangedField.Setup (sFieldName)
    Call mChangedFields.Add(AddChangedField, sFieldName)
  Else
    Set AddChangedField = mChangedFields(sFieldName)
  End If
    
AddChangedField_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "AddChangedField"), "Error adding a field to the collection of changed fields." & vbCrLf & Err.Description)
End Function

Public Sub AddNewField(sFieldName As String)
  mNewFields.Add (sFieldName)
End Sub
Public Sub AddOldField(sFieldName As String)
  mOldFields.Add (sFieldName)
End Sub
Public Property Get TableName() As String
  TableName = mTableName
End Property
Public Property Get ChangedFields() As Collection
  Set ChangedFields = mChangedFields
End Property
Public Property Get NewFields() As StringList
  Set NewFields = mNewFields
End Property
Public Property Get OldFields() As StringList
  Set OldFields = mOldFields
End Property

Private Sub Class_Terminate()
  Set mChangedFields = Nothing
  Set mNewFields = Nothing
  Set mOldFields = Nothing
End Sub
