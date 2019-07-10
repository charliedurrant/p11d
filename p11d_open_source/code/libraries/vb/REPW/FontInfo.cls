VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name As String
Public Size As Single
Public SizeRequested As Single
Public bold As Boolean
Public Italic As Boolean
Public UnderLine As Boolean
Public StrikeThrough As Boolean

Public Property Get Key() As String
  Key = FontKey(Name, SizeRequested, bold, Italic, UnderLine, StrikeThrough)
End Property
