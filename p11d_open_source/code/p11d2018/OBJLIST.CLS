VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_MAXINDEX As Long
Private m_CURINDEX As Long
Private m_INCREMENT As Long
Private Const DEF_INCREMENT As Long = 20
Private m_Objects() As Object

Private Sub Class_Initialize()
  m_INCREMENT = DEF_INCREMENT
End Sub

Private Sub Class_Terminate()
  Call RemoveAll
End Sub

Public Sub RemoveAll()
  Dim i As Long
  For i = m_CURINDEX To 1 Step -1
    Set m_Objects(i) = Nothing
  Next i
  m_CURINDEX = 0
End Sub

Public Sub Remove(ByVal Index As Long)
Dim i As Long
  If (Index < 1) Or (Index > m_CURINDEX) Then Call Err.Raise(380) 'apf
  Set m_Objects(Index) = Nothing
End Sub

Public Sub Trim()
  Dim i As Long
  
  For i = m_CURINDEX To 1 Step -1
    If Not (m_Objects(i) Is Nothing) Then Exit For
  Next i
  m_CURINDEX = i
End Sub

Public Function Add(Item As Object) As Long
  Call xSet("ObjectList_Add")
  
  m_CURINDEX = m_CURINDEX + 1
  If m_CURINDEX > m_MAXINDEX Then
    m_MAXINDEX = m_MAXINDEX + Increment
    ReDim Preserve m_Objects(1 To m_MAXINDEX) As Object
  End If
  Set m_Objects(m_CURINDEX) = Item
  Add = m_CURINDEX
  Call xReturn("ObjectList_Add")
End Function

Public Function Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0
  Call xSet("ObjectList_Get")
  If (Index < 1) Or (Index > m_CURINDEX) Then Err.Raise 380 'apf
  Set Item = m_Objects(Index)
  Call xReturn("ObjectList_Get")
End Function

Public Property Get count() As Long
  count = m_CURINDEX
End Property

Public Property Get Increment() As Long
  Increment = m_INCREMENT
End Property

Public Property Let Increment(newvalue As Long)
  If newvalue < 1 Then Call Err.Raise(380, "ObjectList", "Increment value must be > 0")
  m_INCREMENT = newvalue
End Property



