VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSetCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mDataSets As Collection

Private Sub Class_Initialize()
  Set mDataSets = New Collection
End Sub

Public Function Item(Index As Variant) As ReportDataSet
Attribute Item.VB_UserMemId = 0
  Set Item = mDataSets.Item(Index)
End Function

Public Property Get Count() As Long
  Count = mDataSets.Count
End Property

Public Sub Remove(Index As Variant)
  Call mDataSets.Remove(Index)
End Sub

Public Sub RemoveAll()
  Call ClearCollection(mDataSets)
End Sub

Public Sub Add(Item As ReportDataSet, Key As String)
  Call mDataSets.Add(Item, Key)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mDataSets.[_NewEnum]
End Function

