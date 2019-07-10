VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RequestParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cadxx new
Option Explicit
Private m_ASPPage As String
Private m_Dictionary As Dictionary
Public Sub Add(ByVal Key As String, Value As Variant)
  Key = LCase$(Key)
  Call Dictionary.Add(Key, LCase$(Value))
End Sub
Public Function Exists(ByVal Key As String) As Boolean
  Key = LCase$(Key)
  Exists = Dictionary.Exists(Key)
End Function
Public Function Item(ByVal Key As String, Optional Default As Variant) As Variant
Attribute Item.VB_UserMemId = 0
  Key = LCase$(Key)
  If Not IsMissing(Default) Then
    If Not Exists(Key) Then
      Item = Default
    Else
      GoTo Exists_Key
    End If
  Else
Exists_Key:
    Item = Dictionary(Key)
  End If
End Function
Public Property Let ASPPage(ByVal NewValue As String)
  m_ASPPage = NewValue
End Property
Public Property Get ASPPage() As String
  ASPPage = m_ASPPage
End Property
Public Property Get Count() As Long
  Count = Dictionary.Count
End Property
Private Property Get Dictionary() As Dictionary
  If m_Dictionary Is Nothing Then Set m_Dictionary = New Dictionary
  Set Dictionary = m_Dictionary
End Property
Public Sub Update(ByVal Key As String, ByVal NewValue As Variant, Optional AddItem As Boolean = False)
  If AddItem Then
    If Not Exists(Key) Then
      Call Add(Key, NewValue)
    Else
      Dictionary(Key) = NewValue
    End If
  Else
    Dictionary(Key) = NewValue
  End If
End Sub
Public Sub Remove(ByVal Key As String)
  Call Dictionary.Remove(Key)
End Sub
Public Property Get QueryString()
  Dim s As String
  Dim i As Long

  s = m_ASPPage & "?"
  For i = 0 To m_Dictionary.Count - 1
    s = s & m_Dictionary.Keys(i) & "=" & m_Dictionary.Item(m_Dictionary.Keys(i))
    If i < m_Dictionary.Count - 1 Then
      s = s & "&"
    End If
  Next
  QueryString = s
End Property


Private Sub Class_Initialize()
  m_ASPPage = "Default.asp"
End Sub
