VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecentFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_value  As String
Private m_max As Long
Private m_recentFiles As StringList

Private Sub Class_Initialize()
  Set m_recentFiles = New StringList
  Max = 6
End Sub

Public Property Get Max() As Long
  Max = m_max
End Property
Public Property Let Max(ByVal value As Long)
  m_max = value
  If (m_recentFiles.Count > Me.Max) Then
    While (m_recentFiles.Count > Me.Max)
      Call m_recentFiles.Remove(m_recentFiles.Item(m_recentFiles.Count))
    Wend
  Else
    While (m_recentFiles.Count > Me.Max)
      m_recentFiles.Add ("")
    Wend
  End If
End Property
Private Sub Remove(ByVal value As String)
  If m_recentFiles.IsPresent(value) Then
    Call m_recentFiles.Remove(value)
  End If
End Sub
Public Sub Add(ByVal value As String)
  Call Remove(value)
  If (m_recentFiles.Count + 1 > Me.Max) Then
    m_recentFiles.Remove (m_recentFiles.Item(1))
  End If
  Call m_recentFiles.Add(value)
End Sub
Public Function MenuEntries() As StringList
  Dim sl As StringList
  Dim i As Long
  
  Set sl = New StringList
  
  For i = m_recentFiles.Count To 1 Step -1
    Call sl.Add(m_recentFiles.Item(i))
  Next
  
  Set MenuEntries = sl
End Function
Public Function Serialize() As String
  Dim sb As QString
  Dim i As Long
  
  Set sb = New QString
  
  For i = 1 To m_recentFiles.Count
    sb.Append (m_recentFiles.Item(i))
    If i < m_recentFiles.Count Then
      sb.Append (";")
    End If
  Next
  Serialize = sb.bstr
  
End Function
Public Sub Deserialize(ByVal value As String)
  Dim v As Variant
  Dim i As Long
  
  value = Trim$(value)
  If (Len(value) = 0) Then Exit Sub
  
  v = Split(value, ";")
  
  For i = LBound(v) To UBound(v)
    Call m_recentFiles.Add(v(i))
  Next
  
End Sub



