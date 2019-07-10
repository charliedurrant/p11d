VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_ObjectHash As Dictionary
Private m_bRefreshKeys As Boolean
Private m_AllKeys As Variant

Private Sub Class_Initialize()
  Set m_ObjectHash = New Dictionary
  m_ObjectHash.CompareMode = BinaryCompare
  m_bRefreshKeys = True
End Sub

Public Function Add(ByVal Item As Object, ByVal sKey As String) As Object
  m_bRefreshKeys = True
  Call m_ObjectHash.Add(UCase$(sKey), Item)
  Set Add = Item
End Function

Public Function Item(sKey As String, Optional ByVal RaiseErrors As Boolean = True) As Object
  Dim sKeyUC As String
  On Error GoTo Item_err
  sKeyUC = UCase$(sKey)
  If m_ObjectHash.Exists(sKeyUC) Then Set Item = m_ObjectHash.Item(sKeyUC)
  Exit Function
  
Item_err:
  If RaiseErrors Then Err.Raise 380, "ObjectHash", "Invalid item key [" & sKey & "]"
End Function

Public Property Get Count() As Long
  Count = m_ObjectHash.Count
End Property

Public Function ItemByIndex(ByVal Index As Long) As Object
  If (Index < 1) Or (Index > m_ObjectHash.Count) Then Err.Raise 380, "ObjectHash", "Invalid item index " & Index
  Index = Index - 1
  If m_bRefreshKeys Then m_AllKeys = m_ObjectHash.Keys
  Set ItemByIndex = m_ObjectHash.Item(m_AllKeys(Index))
End Function

Public Property Get Increment() As Long
  Increment = 0
End Property

Public Property Let Increment(ByVal NewValue As Long)
  Err.Raise 380, "ObjectHash", "Increment value is not available"
End Property

Public Sub RemoveAll()
  Call m_ObjectHash.RemoveAll
  m_AllKeys = Empty
  m_bRefreshKeys = True
End Sub


