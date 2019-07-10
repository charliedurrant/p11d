VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Depth As Long
Private m_CarrigeReturn As String
Public Function ElementOpen(Name As String, Optional Attributes As String = "", Optional ByVal ShortClose As Boolean = False) As String
  ElementOpen = lows.ElementOpen(Name, Attributes, ShorClose)
  If ShortClose Then m_Depth = m_Depth - 1
  If g_Debugging Then
    ElementOpen = s
  Else
    ElementOpen = DepthString() & s
  End If
End Function
Private Static Function DepthString() As String
  If m_Depth > 0 Then DepthString = String$(2 * (m_Depth - 1), " ")
End Function
Public Function ElementClose(ByVal Name As String) As String
  Dim s As String
  s = lows.ElementClose(Name)
  If g_Debugging Then
    s = DepthString() & s
    m_Depth = m_Depth - 1
  End If
  ElementClose = s
End Function


