VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavigatorRSField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name As String
Private m_FieldName As String
Private m_field As field
Private m_AsNode As Boolean

Public Sub Init(ByVal fld As field, ByVal Name As String, ByVal FieldName As String, ByVal AsNode As Boolean)
  Set m_field = fld
  m_Name = Name
  m_FieldName = FieldName
  m_AsNode = AsNode
End Sub

Public Property Get Name() As String
  Name = m_Name
End Property

Public Property Get FieldName() As String
  FieldName = m_FieldName
End Property

Public Property Get AsNode() As Boolean
  AsNode = m_AsNode
End Property

Public Property Set field(ByVal NewValue As field)
  Set m_field = NewValue
End Property

Public Property Get field() As field
  Set field = m_field
End Property

