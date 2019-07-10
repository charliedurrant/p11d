VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ImportSpec"
Attribute VB_Ext_KEY = "Member0" ,"ImportSpec"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

Public Function Add() As ImportSpec
  Dim objNewMember As New ImportSpec
  mCol.Add objNewMember
  Set Add = objNewMember
  Set objNewMember = Nothing
End Function

Public Function AddObject(ByVal NewMember As ImportSpec) As ImportSpec
  mCol.Add NewMember
  Set AddObject = NewMember
End Function

Public Property Get Item(ByVal Index As Long) As ImportSpec
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(Index)
End Property

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Sub Remove(ByVal Index As Long)
  Call mCol.Remove(Index)
End Sub

Public Sub RemoveAll()
  Do While mCol.Count > 0
    Call mCol.Remove(1)
  Loop
End Sub

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  Do While mCol.Count > 0
    Call mCol.Remove(1)
  Loop
  Set mCol = Nothing
End Sub

