VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IndexList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_MAXINDEX As Long
Private m_CURINDEX As Long
Private m_INCREMENT As Long

Private Const DEF_INCREMENT As Long = 20
Private m_Indices() As Long

Private Sub Class_Initialize()
  m_INCREMENT = DEF_INCREMENT
End Sub

Private Sub Class_Terminate()
  Call RemoveAll
End Sub

Public Property Get Increment() As Long
  Increment = m_INCREMENT
End Property

Public Property Let Increment(ByVal NewValue As Long)
  If NewValue < 1 Then Err.Raise 380, "IndexList", "Increment value must be > 0"
  m_INCREMENT = NewValue
End Property

Public Function Add(ByVal Index As Long) As Long
  Add = FindIndex(Index)
  If Add < 0 Then
    m_CURINDEX = m_CURINDEX + 1
    If m_CURINDEX > m_MAXINDEX Then
      m_MAXINDEX = m_MAXINDEX + m_INCREMENT
      If IsArray(m_Indices) Then
        ReDim Preserve m_Indices(1 To m_MAXINDEX) As Long
      Else
        ReDim m_Indices(1 To m_MAXINDEX) As Long
      End If
    End If
    m_Indices(m_CURINDEX) = Index
    Add = m_CURINDEX
  End If
End Function

Public Function IsPresent(ByVal Index As Long) As Boolean
  IsPresent = FindIndex(Index) <> -1
End Function

Public Function Item(ByVal Index As Long) As Long
  If (Index < 1) Or (Index > m_CURINDEX) Then Err.Raise 380, "IndexList invalid index [" & Index & "]"
  Item = m_Indices(Index)
End Function

Public Function ItemIndex(ByVal Index As Long) As Long
  ItemIndex = FindIndex(Index)
End Function

Private Function FindIndex(ByVal Index As Long) As Long
  Dim i As Long
  For i = 1 To m_CURINDEX
    If Index = m_Indices(i) Then
      FindIndex = i: Exit Function
    End If
  Next i
  FindIndex = -1
End Function

Public Property Get Count() As Long
  Count = m_CURINDEX
End Property

Public Sub RemoveAll()
  m_CURINDEX = 0
End Sub


