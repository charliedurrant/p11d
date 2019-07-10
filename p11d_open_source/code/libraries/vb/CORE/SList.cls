VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_MAXINDEX As Long
Private m_CURINDEX As Long
Private m_INCREMENT As Long

Private Const DEF_INCREMENT As Long = 20
Private m_Strings() As String

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
  If NewValue < 1 Then Err.Raise 380, "StringList", "Increment value must be > 0"
  m_INCREMENT = NewValue
End Property

Public Function Add(ByVal Item As String) As Long
  Add = FindIndex(Item)
  If Add < 0 Then
    m_CURINDEX = m_CURINDEX + 1
    If m_CURINDEX > m_MAXINDEX Then
      m_MAXINDEX = m_MAXINDEX + m_INCREMENT
      If IsArrayEx2(m_Strings) Then
        ReDim Preserve m_Strings(1 To m_MAXINDEX) As String
      Else
        ReDim m_Strings(1 To m_MAXINDEX) As String
      End If
    End If
    m_Strings(m_CURINDEX) = Item
    Add = m_CURINDEX
  End If
End Function

Public Function Item(ByVal Index As Long) As String
  If (Index < 1) Or (Index > m_CURINDEX) Then Err.Raise 381, "StringList"
  Item = m_Strings(Index)
End Function

Public Function IsPresent(ByVal String1 As String) As Boolean
  IsPresent = FindIndex(String1) <> -1
End Function

Public Function ItemIndex(ByVal Item As String) As Long
  ItemIndex = FindIndex(Item)
End Function

Private Function FindIndex(ByVal String1 As String) As Long
  Dim i As Long
  For i = 1 To m_CURINDEX
    If StrComp(String1, m_Strings(i), vbTextCompare) = 0 Then
      FindIndex = i: Exit Function
    End If
  Next i
  FindIndex = -1
End Function

Public Property Get Count() As Long
  Count = m_CURINDEX
End Property

Public Sub Remove(vItem As Variant)
  Dim i As Long, Index As Long
  
  If VarType(vItem) = vbString Then
    Index = FindIndex(CStr(vItem))
  Else
    Index = CLng(vItem)
  End If
  If (Index < 1) Or (Index > m_CURINDEX) Then Err.Raise 380, "StringList", "Index error"
  If m_CURINDEX <> Index Then m_Strings(Index) = m_Strings(m_CURINDEX)
  m_CURINDEX = m_CURINDEX - 1
End Sub

Public Sub RemoveAll()
  m_CURINDEX = 0
End Sub
