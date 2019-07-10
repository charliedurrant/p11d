VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_MAXINDEX As Long
Private m_CURINDEX As Long
Private m_INCREMENT As Long
Private Const DEF_INCREMENT As Long = 100
Private m_Objects() As Object

Private Sub Class_Initialize()
  m_INCREMENT = DEF_INCREMENT
End Sub

Private Sub Class_Terminate()
  Call RemoveAll
End Sub

Public Function Add(Item As Object) As Long
  m_CURINDEX = m_CURINDEX + 1
  If m_CURINDEX > m_MAXINDEX Then
    m_MAXINDEX = m_MAXINDEX + m_INCREMENT
    If IsArrayEx2(m_Objects) Then
      ReDim Preserve m_Objects(1 To m_MAXINDEX) As Object
    Else
      ReDim m_Objects(1 To m_MAXINDEX) As Object
    End If
  End If
  Set m_Objects(m_CURINDEX) = Item
  Add = m_CURINDEX
End Function

Public Function AddIndex(Item As Object, ByVal Index As Long) As Long
re_inc:
  If Index > m_MAXINDEX Then
    m_MAXINDEX = m_MAXINDEX + m_INCREMENT
    If IsArrayEx2(m_Objects) Then
      ReDim Preserve m_Objects(1 To m_MAXINDEX) As Object
    Else
      ReDim m_Objects(1 To m_MAXINDEX) As Object
    End If
    GoTo re_inc
  End If
  If Not m_Objects(Index) Is Nothing Then Err.Raise 380, "AddIndex", "Item already allocated"
  Set m_Objects(Index) = Item
  If Index > m_CURINDEX Then m_CURINDEX = Index
  AddIndex = m_CURINDEX
End Function

Public Function ItemIndex(ObjectRef As Object) As Long
  Dim i As Long
  
  For i = 1 To m_CURINDEX
    If m_Objects(i) Is ObjectRef Then
      ItemIndex = i
      Exit Function
    End If
  Next i
  ItemIndex = 0
End Function

Public Function Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0
  If (Index < 1) Or (Index > m_CURINDEX) Then Err.Raise 380, "ObjectList", "Invalid item index " & Index
  Set Item = m_Objects(Index)
End Function

Public Property Get Count() As Long
  Count = m_CURINDEX
End Property

Public Property Get CountValid() As Long
  Dim i As Long
  
  For i = 1 To m_CURINDEX
    If Not m_Objects(i) Is Nothing Then
      CountValid = CountValid + 1
    End If
  Next i
End Property

Public Sub Remove(vItem As Variant)
  Dim i As Long, Index As Long, obj As Object
  
  If VarType(vItem) = vbObject Then
    Set obj = vItem
    Index = ItemIndex(obj)
  Else
    Index = CLng(vItem)
  End If
  If (Index < 1) Or (Index > m_CURINDEX) Then Err.Raise 380, "ObjectList", "Index error"
  Set m_Objects(Index) = Nothing
  If Index = m_CURINDEX Then m_CURINDEX = m_CURINDEX - 1
End Sub

Public Sub RemoveAll()
  Dim i As Long
  For i = m_CURINDEX To 1 Step -1
    Set m_Objects(i) = Nothing
  Next i
  m_CURINDEX = 0
End Sub

Public Sub Compact()
  Dim i As Long, j As Long
  
  j = m_CURINDEX
  For i = 1 To m_CURINDEX
    If i >= j Then Exit For
    If m_Objects(i) Is Nothing Then
      For j = j To 1 Step -1
        If Not (m_Objects(j) Is Nothing) Or (j <= i) Then Exit For
      Next j
      If j > i Then
        Set m_Objects(i) = m_Objects(j)
        Set m_Objects(j) = Nothing
        j = j - 1
      End If
    End If
  Next i
  Call CompactTop ' set curindex to top value
End Sub

Public Sub CompactTop()
  Dim i As Long
  
  ' set curindex to top value
  For i = m_CURINDEX To 1 Step -1
    If Not (m_Objects(i) Is Nothing) Then Exit For
  Next i
  m_CURINDEX = i
End Sub

Public Property Get Increment() As Long
  Increment = m_INCREMENT
End Property

Public Property Let Increment(ByVal NewValue As Long)
  If NewValue < 1 Then Err.Raise 380, "ObjectList", "Increment value must be > 0"
  m_INCREMENT = NewValue
End Property

Public Sub Sort(SortFn As ISortFunction)
  If m_CURINDEX > 1 Then Call SortAny(m_Objects, 1, m_CURINDEX, SortFn, QUICK_SORT)
End Sub


