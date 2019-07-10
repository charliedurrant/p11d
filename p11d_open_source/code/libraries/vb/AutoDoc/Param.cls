VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Description As String
Public VarType As String
Public IsParamArray As Boolean
Public IsOptional As Boolean
Public IsArray As Boolean
Public IsByval As Boolean
Public DefaultValue As String
Private m_Text As String

Public Property Get Text() As String
  If Len(m_Text) > 0 Then
    Text = m_Text
  Else
    m_Text = ""
    If IsOptional Then m_Text = m_Text & "Optional "
    If IsByval Then m_Text = m_Text & "ByVal "
    If IsParamArray Then m_Text = m_Text & "ParamArray "
    m_Text = m_Text & Name
    If IsArray Then m_Text = m_Text & "()"
    If IsParamArray Then
      m_Text = m_Text & " as Variant"
    Else
      m_Text = m_Text & " as " & VarType
      If Len(DefaultValue) > 0 Then m_Text = m_Text & " = " & DefaultValue
    End If
    Text = m_Text
  End If
End Property


